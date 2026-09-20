# Authentication Smoke Test

## Purpose

`scripts/smoke_test_authentication.ps1` validates the authentication,
authorization, and tenant-isolation boundaries against a real running API and
a disposable SQLite database.

The script supports:

- automated mode for a repeatable pass/fail check;
- manual mode for interactive threat testing through Swagger or PowerShell.

It does not use or modify `hospital_schedule.db`.

## Prerequisites

- Run the script from the `backend` directory.
- Use PowerShell 7 or later.
- Install the project environment with `uv`.
- Choose a local port that is not already in use.

## Automated mode

Run:

```powershell
.\scripts\smoke_test_authentication.ps1
```

Use another port when required:

```powershell
.\scripts\smoke_test_authentication.ps1 -Port 8877
```

A successful run ends with:

```text
Authentication smoke test passed: 147 checks.
```

The process exits with status code `0` on success and `1` on failure.

## What automated mode checks

### Authentication and route protection

- Public health access.
- Login for super-admin, department-admin, Doctor, and viewer roles.
- Authenticated `/auth/me`.
- `401` and `WWW-Authenticate: Bearer` for anonymous protected requests.
- Removed public signup, User lookup, and Department creation routes.

### Role authorization

- Doctors and viewers cannot perform department-admin writes.
- Super-admin cannot perform routine tenant writes.
- Department admin cannot use the global Department list.
- Viewer cannot read raw Doctor unavailability.

### Tenant isolation

The script creates Department A and Department B with matching Team, Position,
Shift, and Doctor names. It then verifies:

- Department A lists exclude Department B IDs.
- Department A requests for known Department B UUIDs return `404`.
- Cross-tenant Doctor, Shift, ShiftAssignment, pre-assignment, and
  Doctor-Position writes are rejected.
- Rejected writes do not create records in either tenant.
- Foreign Shift and foreign Doctor assignment failures are indistinguishable.

### Ownership and response privacy

- Doctor A can create their own unavailability.
- Doctor A receives `403` for Doctor B's unavailability.
- A Department A admin receives `404` for Doctor B's unavailability.
- `/doctors/roster` omits contact email.
- Department-admin Doctor detail includes contact email.

### Derived tenant ownership

- Team, Position, and Doctor request bodies reject `department_id`.
- Successful Team, Position, and Doctor creation stores Department A from the
  authenticated admin.
- Shift ownership is derived through a Position already scoped to Department
  A.
- Position and Shift detail routes work with UUIDs.

### Department provisioning and invitation onboarding

- Super-admin atomic department provisioning with initial department-admin invitation.
- Department-admin public signup via one-time invitation token.
- Invited admin account verified with correct `department_admin` role and department assignment.
- Department admin creates staff invitations for doctors.
- Doctor public signup auto-provisions and links doctor entity record.
- Rejection of already consumed invitation tokens (`400 Bad Request`).
- Department admin creates and revokes viewer invitations.
- Rejection of revoked invitation tokens (`400 Bad Request`).
- Department-scoped invitation listing.

### Refresh session and token rotation lifecycle (RFC 6819)

- Login returns session credentials (`access_token`, `refresh_token`, and `csrf_token`).
- Token rotation: `POST /api/v1/auth/refresh` rotates token $T_1 \to T_2$ and issues a fresh access token.
- Verified access to protected endpoints using newly rotated access token.
- Reuse detection: replaying stale token $T_1$ triggers `401 Unauthorized` ("reuse detected").
- Lineage revocation: reuse detection immediately invalidates the entire session family (subsequent presentation of $T_2$ returns `401`).
- Logout terminates active refresh session; post-logout refresh attempt returns `401`.

### Endpoint rate limiting and abuse protection

- Throttling boundary: 5 warmup signup attempts from simulated IP (`X-Forwarded-For: 198.51.100.42`).
- 6th request triggers `HTTP 429 Too Many Requests`.
- Verified presence of positive integer `Retry-After` response header.
- Client IP isolation: requests from an independent client IP (`198.51.100.43`) succeed normally and receive standard `404` without being throttled.


## Disposable environment and cleanup

Every run:

1. Creates a uniquely named `authentication_smoke_<guid>.db`.
2. Overrides `DATABASE_URL`, `SECRET_KEY`, `LOG_LEVEL`, and the internal smoke
   password only for the script process.
3. Seeds two departments and the required accounts and resources.
4. Starts a dedicated Uvicorn process on the requested port.
5. Runs the checks or waits for manual testing.
6. Stops that exact Uvicorn process.
7. Deletes the exact database, journal, WAL, SHM, and temporary log files.
8. Restores the original process environment and working directory.

Cleanup also runs when a normal script error occurs. In manual mode, press
Enter to finish cleanly instead of closing the terminal.

## Manual mode

Run:

```powershell
.\scripts\smoke_test_authentication.ps1 -Manual
```

The script securely prompts for one password used by all disposable test
accounts. It does not print the password or any access token.

After startup, it prints:

- the API and Swagger URLs;
- the test account emails and roles;
- Department A and Department B UUIDs;
- Team, Position, Shift, Doctor, and ShiftAssignment UUIDs for both tenants;
- the manual threat-check sequence.

The API remains available until you press Enter in the script terminal.

### Authenticate with Swagger

1. Open the printed Swagger URL, normally
   `http://127.0.0.1:8765/docs`.
2. Use `POST /api/v1/auth/login`, or the Swagger authorization dialog.
3. Enter one of the printed account emails and the password entered when the
   script started.
4. Repeat authentication when switching roles.

### Authenticate with PowerShell

In a second PowerShell terminal:

```powershell
$baseUri = "http://127.0.0.1:8765"
$securePassword = Read-Host "Smoke-test password" -AsSecureString
$credential = [pscredential]::new("unused", $securePassword)
$plainPassword = $credential.GetNetworkCredential().Password

$login = Invoke-RestMethod `
    -Method Post `
    -Uri "$baseUri/api/v1/auth/login" `
    -ContentType "application/x-www-form-urlencoded" `
    -Body @{
        username = "<printed-account-email>"
        password = $plainPassword
    }

$plainPassword = $null
$headers = @{
    Authorization = "Bearer $($login.access_token)"
}
```

Do not print or persist `$login.access_token`.

### Example manual requests

List the authenticated tenant's Teams:

```powershell
Invoke-RestMethod `
    -Method Get `
    -Uri "$baseUri/api/v1/teams/" `
    -Headers $headers
```

Request a known Department B Team while authenticated as Department A:

```powershell
Invoke-WebRequest `
    -Method Get `
    -Uri "$baseUri/api/v1/teams/<department-B-team-id>" `
    -Headers $headers `
    -SkipHttpErrorCheck
```

Expected result:

```text
404
```

Attempt to create a Shift under Department B's Position:

```powershell
Invoke-WebRequest `
    -Method Post `
    -Uri "$baseUri/api/v1/shifts/" `
    -Headers $headers `
    -ContentType "application/json" `
    -Body (@{
        name = "Forbidden Cross-Tenant Shift"
        doctors_per_shift = 1
        grants_day_off = $false
        position_id = "<department-B-position-id>"
    } | ConvertTo-Json) `
    -SkipHttpErrorCheck
```

Expected result:

```text
404
```

## Manual threat-check checklist

Using the printed IDs:

1. Log in as Department A admin.
2. List Users, Doctors, Teams, Positions, Shifts, and ShiftAssignments.
3. Confirm no Department B ID appears.
4. Request each known Department B resource and confirm `404`.
5. Attempt cross-tenant Doctor, Shift, assignment, pre-assignment, and
   Doctor-Position writes.
6. Repeat the relevant list calls and confirm no record was created.
7. Log in as Doctor A and access Doctor A's unavailability successfully.
8. Access Doctor B's unavailability as Doctor A and confirm `403`.
9. Retrieve `/api/v1/doctors/roster` and confirm `email` is absent.
10. Log in as Department A admin, retrieve Doctor A detail, and confirm
    `email` is present.
11. Log in as super-admin and confirm routine tenant writes return `403`.
12. Attempt to sign up using an already-consumed or revoked invitation token and confirm `400 Bad Request`.
13. Rotate a refresh token, replay the older token, and confirm the entire session family is revoked with `401 Unauthorized` ("reuse detected").
14. Send rapid requests to a rate-limited endpoint (e.g. `POST /api/v1/auth/signup`) and confirm `HTTP 429 Too Many Requests` with a valid `Retry-After` header.

Press Enter in the original script terminal after completing the checks.

## Parameters

| Parameter | Purpose |
|---|---|
| `-Port` | Selects the local API port. The default is `8765`. |
| `-Manual` | Starts an interactive manual-testing session instead of running the automated checks. |
| `-ManualPassword` | Accepts an existing `SecureString`; otherwise manual mode prompts for one. Never pass plaintext. |
| `-ManualWaitSeconds` | Keeps manual mode open for a fixed duration instead of waiting for Enter. Primarily useful for diagnostics. |

Example using an existing `SecureString`:

```powershell
$password = Read-Host "Smoke-test password" -AsSecureString

.\scripts\smoke_test_authentication.ps1 `
    -Manual `
    -ManualPassword $password `
    -Port 8877
```

## Troubleshooting

### PowerShell version error

The script requires PowerShell 7:

```powershell
$PSVersionTable.PSVersion
```

Run it with `pwsh` if the current shell is Windows PowerShell:

```powershell
pwsh -NoProfile -File .\scripts\smoke_test_authentication.ps1
```

### Port already in use

Select another port:

```powershell
.\scripts\smoke_test_authentication.ps1 -Port 8877
```

### API startup or request failure

The script prints the last server error lines before cleanup. Fix the reported
startup or API error and rerun the script; do not disable the failing
assertion merely to make the smoke test pass.
