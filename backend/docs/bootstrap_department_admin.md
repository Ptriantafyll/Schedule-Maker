# Bootstrap a Department Admin

## Purpose

The department-admin bootstrap command creates a trusted `department_admin` account (and its associated hospital `Department` if it does not already exist) directly through the backend service layer.

This provides an immediate, secure mechanism to provision initial departmental admin accounts prior to end-user invitation workflows.

A department admin is authorized to manage doctor rosters, schedules, and shift preferences strictly within their own department.

---

## What the command does

The command:

1. Connects to the database selected by `DATABASE_URL`.
2. Creates any missing database tables via `init_db()`.
3. Checks if a department with `--department-name` exists:
   - If found, reuses the existing department.
   - If not found, creates a new department. If `--department-code` is not provided, automatically derives a short uppercase code (up to the first 4 characters of the department name).
4. Prompts for the password without displaying it in the terminal.
5. Requires the password to be entered twice for confirmation.
6. Validates password complexity against the system password policy.
7. Hashes the password using bcrypt before persistence.
8. Creates a User with:
   - `role = department_admin`;
   - `department_id = <target_department_id>`;
   - `doctor_id = null`.
9. Enforces ACID atomicity: if user creation fails midway (e.g. duplicate email), any newly created department is rolled back cleanly.
10. Emits a structured security audit event (`department_admin.bootstrap`).

It does not:
- Accept a role argument (always `department_admin`).
- Accept a doctor ID (must be `null`).
- Print the password or password hash.
- Accept passwords via command-line flags (to prevent leaking into shell history).

---

## Prerequisites

- Run the command from the `backend` directory.
- Install the project environment with `uv`.
- Set `DATABASE_URL` before running if targeting an instance other than the default local `hospital_schedule.db`.

---

## Create a Department Admin

From `backend`:

```powershell
uv run --no-sync python -m scripts.bootstrap_department_admin `
    --department-name "Cardiology" `
    --department-code "CARD" `
    --email "admin@cardio.hospital.com" `
    --full-name "Dr. Sarah Connor"
```

If `--department-code` is omitted, the script automatically generates a code (e.g., `"CARD"` for `"Cardiology"`):

```powershell
uv run --no-sync python -m scripts.bootstrap_department_admin `
    --department-name "Cardiology" `
    --email "admin@cardio.hospital.com" `
    --full-name "Dr. Sarah Connor"
```

The command prompts securely:

```text
Password:
Confirm password:
```

On success:

```text
Department: Cardiology (CODE: CARD, id: 01928374-...)
Department admin created successfully: Dr. Sarah Connor, admin@cardio.hospital.com
```

The command exits with status code `0`.

---

## Select a Different Database

Set `DATABASE_URL` in the same PowerShell session before running the command:

```powershell
$env:DATABASE_URL = "sqlite:///hospital_schedule.db"

uv run --no-sync python -m scripts.bootstrap_department_admin `
    --department-name "Emergency" `
    --department-code "ER" `
    --email "admin@er.hospital.com" `
    --full-name "Dr. John Dorian"
```

---

## Verify the Account via API

1. Start the API server:

```powershell
uv run --no-sync uvicorn src.main:app
```

2. Authenticate from PowerShell:

```powershell
$securePassword = Read-Host "Password" -AsSecureString
$credential = [pscredential]::new("unused", $securePassword)
$plainPassword = $credential.GetNetworkCredential().Password

$login = Invoke-RestMethod `
    -Method Post `
    -Uri "http://127.0.0.1:8000/api/v1/auth/login" `
    -ContentType "application/x-www-form-urlencoded" `
    -Body @{
        username = "admin@cardio.hospital.com"
        password = $plainPassword
    }

$plainPassword = $null
$headers = @{
    Authorization = "Bearer $($login.access_token)"
}

Invoke-RestMethod `
    -Method Get `
    -Uri "http://127.0.0.1:8000/api/v1/auth/me" `
    -Headers $headers
```

The returned profile should confirm:

```json
{
  "email": "admin@cardio.hospital.com",
  "full_name": "Dr. Sarah Connor",
  "role": "department_admin",
  "department_id": "<valid-department-uuid>",
  "doctor_id": null
}
```

---

## Failure Behavior

The command exits with status code `1` when:
- The password is empty.
- The two password entries do not match.
- The password fails complexity validation (under 10 characters, missing uppercase/lowercase/digit/symbol, or common password).
- The email address already belongs to an existing user.

Invalid or forbidden CLI arguments (e.g. `--role`, `--password`, `--doctor-id`) are rejected by `argparse` with status code `2`.
