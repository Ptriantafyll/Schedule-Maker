# Bootstrap a Super-Admin

## Purpose

The bootstrap command creates a trusted `super_admin` account directly through
the backend service layer. This is the only supported way to create a
super-admin before the protected control-plane provisioning API is available.

A super-admin is a global control-plane account. It can access explicitly
authorized administration operations, but it does not belong to a department
and must not perform routine tenant scheduling work.

## What the command does

The command:

1. Connects to the database selected by `DATABASE_URL`.
2. Creates any missing database tables.
3. Prompts for the password without displaying it.
4. Requires the password to be entered twice.
5. Hashes the password before persistence.
6. Creates a User with:
   - `role = super_admin`;
   - `department_id = null`;
   - `doctor_id = null`.
7. Rejects an email address that already belongs to any User.

It does not:

- accept a role argument;
- accept a department or Doctor ID;
- print the password or password hash;
- create or modify a Department;
- update an existing User.

## Prerequisites

- Run the command from the `backend` directory.
- Install the project environment with `uv`.
- Set `DATABASE_URL` before running the command if the target is not the
  default local `hospital_schedule.db`.

## Create a super-admin

From `backend`:

```powershell
uv run --no-sync python -m scripts.bootstrap_super_admin `
    --email "admin@example.com" `
    --full-name "System Administrator"
```

The command prompts securely:

```text
Password:
Confirm password:
```

The password is intentionally not accepted as a command-line argument because
command-line values can be retained in shell history or process listings.

On success:

```text
Super admin created successfully
```

The command exits with status code `0`.

## Select a different database

Set `DATABASE_URL` in the same PowerShell session before running the command:

```powershell
$env:DATABASE_URL = "sqlite:///bootstrap_test.db"

uv run --no-sync python -m scripts.bootstrap_super_admin `
    --email "admin@example.com" `
    --full-name "System Administrator"
```

For the normal local development database, either omit `DATABASE_URL` or set:

```powershell
$env:DATABASE_URL = "sqlite:///hospital_schedule.db"
```

Always confirm the target database before creating a privileged account.

## Verify the account

Start the API:

```powershell
uv run --no-sync uvicorn src.main:app
```

Then authenticate from another PowerShell terminal:

```powershell
$securePassword = Read-Host "Password" -AsSecureString
$credential = [pscredential]::new("unused", $securePassword)
$plainPassword = $credential.GetNetworkCredential().Password

$login = Invoke-RestMethod `
    -Method Post `
    -Uri "http://127.0.0.1:8000/api/v1/auth/login" `
    -ContentType "application/x-www-form-urlencoded" `
    -Body @{
        username = "admin@example.com"
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

The returned profile should have:

```text
role: super_admin
department_id: null
doctor_id: null
```

Do not print, log, or persist the access token.

## Failure behavior

The command exits with status code `1` when:

- the password is empty;
- the two password entries do not match;
- the email address already exists;
- account creation loses a concurrent duplicate-email race.

Invalid or missing CLI arguments are rejected by `argparse`.

To display the supported arguments:

```powershell
uv run --no-sync python -m scripts.bootstrap_super_admin --help
```

## Creating another super-admin

Additional super-admins must be created by running this same trusted local
command again with a different email address. There is no public API endpoint
for creating or promoting a super-admin.
