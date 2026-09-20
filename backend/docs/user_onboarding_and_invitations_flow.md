# User Onboarding and Invitation Architecture

This document describes the end-to-end user onboarding lifecycle, role hierarchy, invitation mechanisms, and multi-department scalability design in Schedule-Maker.

---

## 1. Core Philosophy: Why No Open Public Signup?

In a hospital scheduling system:
- Access to physician rosters, doctor contact information, and shift schedules contains sensitive operational and personal data.
- Allowing generic public registration (`POST /users/signup`) introduces serious security vulnerabilities: arbitrary callers could elevate themselves to administrators, attach themselves to real doctors' identities, or access other departments' private rosters.
- Therefore, **Schedule-Maker strictly enforces an invitation-only onboarding model**. No user account can be created without an explicit, cryptographically secured invitation issued by a verified authority.

---

## 2. Role Hierarchy and Authorization Boundaries

The platform operates on a clear three-tier administrative hierarchy:

```
                      ┌─────────────────────────┐
                      │    Super Administrator   │
                      │  (System Owner / DevOps)│
                      └────────────┬────────────┘
                                   │
                    Provisions Departments & Invites
                                   │
                                   ▼
                      ┌─────────────────────────┐
                      │  Department Admin(s)    │
                      │ (Clinical Chief / Lead) │
                      └────────────┬────────────┘
                                   │
                  Manages Roster & Invites Staff
                                   │
                   ┌───────────────┴───────────────┐
                   ▼                               ▼
      ┌─────────────────────────┐     ┌─────────────────────────┐
      │         Doctors         │     │         Viewers         │
      │  (Attending Physicians) │     │ (Nurses / Coordinators) │
      └─────────────────────────┘     └─────────────────────────┘
```

### Roles and Their Scopes

1. **Super Admin (`SUPER_ADMIN`)**:
   - **Identity**: System owner / IT administration (tenantless; `department_id = None`, `doctor_id = None`).
   - **Scope**: Global system instance.
   - **Capabilities**:
     - Provisions new clinical departments (e.g. Cardiology, Radiology).
     - Issues the initial Department Administrator invitation for a department.
     - Reissues or revokes department-admin invitations.
     - Oversees database health and system configurations.
   - **Restrictions**: Cannot perform routine department operations (cannot add shifts, schedule doctors, or edit departmental rules).

2. **Department Admin (`DEPARTMENT_ADMIN`)**:
   - **Identity**: Clinical chief or administrative manager of a specific department (`department_id` is required; optional `doctor_id` if the chief is also an active practicing doctor on shift).
   - **Scope**: Strictly isolated to their assigned `department_id`.
   - **Capabilities**:
     - Manages departmental roster: Doctors, Teams, Positions, and Shifts.
     - Runs the OR-Tools scheduling engine to generate monthly duty rosters.
     - Invites Doctors and Viewers to access the portal.
     - Lists and revokes pending doctor/viewer invitations within their department.
   - **Restrictions**: Zero access to other departments' data (enforced at API route dependency, database, and repository layers). Cannot create new departments or invite other department admins.

3. **Doctor (`DOCTOR`)**:
   - **Identity**: A physician with an active profile on the department roster (`department_id` and `doctor_id` both required).
   - **Scope**: Self-service and department view.
   - **Capabilities**:
     - Views published department schedules and rosters.
     - Submits personal availability and shift preferences.
   - **Restrictions**: Cannot edit department configuration, manage other doctors, or invite users.

4. **Viewer (`VIEWER`)**:
   - **Identity**: Hospital staff needing schedule visibility (head nurses, medical coordinators, interns) (`department_id` required; `doctor_id = None`).
   - **Scope**: Read-only schedule visibility within their department.

---

## 3. End-to-End Onboarding Flow (Approach A)

### Phase 0: Super Admin Bootstrap
The Super Admin account is created directly via the secure CLI tool during initial deployment:
```powershell
python -m scripts.bootstrap_super_admin --email "admin@hospital.org" --full-name "System Administrator"
```
The command prompts for a secure password via standard terminal masking. No public HTTP endpoint exists to create a super admin.

---

### Phase 1: Department Provisioning & Admin Invitation
When a new clinical department (e.g. Cardiology) is established in the hospital:

1. **Request Out-of-Band**: The hospital management or department chief agrees on onboarding with the Super Admin.
2. **Super Admin Provisions the Department**:
   - Super Admin calls `POST /api/v1/admin/provision-department` (or via Admin UI) with:
     ```json
     {
       "department_name": "Cardiology",
       "department_code": "CARD",
       "admin_email": "house@hospital.org",
       "admin_name": "Dr. Gregory House"
     }
     ```
3. **Atomic Execution**:
   - The database creates the `Department` record.
   - Generates a high-entropy URL-safe token (32 bytes).
   - Hashes the token with SHA-256 and stores an `Invitation` record bound to `department_id`, `admin_email`, and `role: DEPARTMENT_ADMIN`.
   - Both department and invitation are committed atomically.
4. **Copyable Link**: The response returns the raw token **once**:
   ```json
   {
     "department_id": "c1a2...",
     "admin_email": "house@hospital.org",
     "invitation_link": "https://app.hospital.org/signup?token=inv_sec_9Fk2..."
   }
   ```
5. The Super Admin sends this link to Dr. House via official hospital email or internal messaging.

---

### Phase 2: Department Admin Activation / Signup
1. Dr. House opens the link: `https://app.hospital.org/signup?token=inv_sec_9Fk2...`
2. The UI reads the token and asks Dr. House to set his password:
   ```json
   POST /api/v1/auth/signup
   {
     "invitation_token": "inv_sec_9Fk2...",
     "password": "CorrectHorseBatteryStaple123!"
   }
   ```
3. **Server Validation**:
   - Hashes `invitation_token` and matches the stored `token_hash`.
   - Verifies the invitation is unused (`used_at IS NULL`), unexpired, and unrevoked.
   - Verifies the department is active.
   - Derives identity immutably:
     - `email = invitation.email` (`house@hospital.org`)
     - `full_name = invitation.full_name` (`Dr. Gregory House`)
     - `role = UserRole.DEPARTMENT_ADMIN`
     - `department_id = invitation.department_id`
     - `doctor_id = None`
   - Hashes password with bcrypt and creates the `User` account via `stage_user_account`.
   - Marks invitation `used_at = now()`.
   - Commits transaction atomically.
4. Dr. House is logged in and redirected to his Cardiology dashboard.

---

### Phase 3: Doctor & Viewer Invitation
Once inside the Cardiology dashboard:

1. **Add Doctor to Roster First**:
   - Dr. House creates doctors in the department roster (`POST /api/v1/doctors`):
     - Name: `Dr. Allison Cameron`
     - Email: `cameron@hospital.org`
     - Team: `ER Team A`
2. **Issue Invitation**:
   - Next to Dr. Cameron's name in the roster, Dr. House clicks **"Invite to Portal"**.
   - Backend creates an `Invitation`:
     - `role: UserRole.DOCTOR`
     - `department_id: Cardiology`
     - `doctor_id: Dr. Cameron's UUID`
     - `email: cameron@hospital.org`
   - Returns the one-time invitation link.
3. Dr. House shares the link with Dr. Cameron.

---

### Phase 4: Doctor / Viewer Activation / Signup
1. Dr. Cameron opens the link and submits:
   ```json
   POST /api/v1/auth/signup
   {
     "invitation_token": "inv_sec_7xL1...",
     "password": "DoctorSecurePassword456!"
   }
   ```
2. The server consumes the token, binds the user account directly to Dr. Cameron's existing doctor record, and activates her login.
3. When Dr. Cameron logs in, the portal automatically loads her assigned shifts and lets her submit unavailability.

---

## 4. Cryptographic Security Model

To protect against database leakage and brute-force attacks:

| Element | Security Specification |
|---------|------------------------|
| **Token Entropy** | Generated using `secrets.token_urlsafe(32)` (~256 bits of cryptographic entropy). Unguessable offline or online. |
| **Token Storage** | **Raw tokens are never stored in the database.** The database stores `token_hash = sha256(raw_token)`. If a database dump is leaked, attackers cannot consume pending invitations. |
| **Exposure Window** | The raw token is returned over HTTPS to the creator **exactly once** in the creation response. It is never logged or returned by listing endpoints. |
| **Expiration** | Default expiration of 7 days (`expires_at = utcnow() + 7 days`). Expired tokens are rejected. |
| **Single-Use** | Upon successful signup, `used_at` is stamped. Subsequent requests with the same token fail with `400 Bad Request`. |
| **Revocation** | Admins can revoke a pending invite at any time (`revoked_at = utcnow()`), immediately rendering it invalid. |

---

## 5. Multi-Department Architecture & Future Scaling

### How Multi-Department Isolation Works Today
All data records (doctors, shifts, teams, positions, user accounts) contain a `department_id` foreign key.
- FastAPI dependency injection (`require_department_scope`, `require_department_admin`) automatically filters every query by the authenticated user's `department_id`.
- Multiple departments live side-by-side in the same database without data cross-contamination.

```
Database: hospital_schedule.db
 ├── Department: Cardiology (id: 1111)
 │    ├── Doctor: Dr. House
 │    ├── Team: Cardio Team A
 │    └── Shift: Night Shift Cardio
 └── Department: Radiology (id: 2222)
      ├── Doctor: Dr. Wilson
      ├── Team: Rad Imaging Team
      └── Shift: Day Scan Shift
```

### Future Extension: Multi-Hospital Healthcare Networks
If a healthcare organization expands to manage multiple physical hospital facilities (e.g. "St. Jude General Hospital" and "Westside Clinic"), our architecture scales cleanly:

```
                      ┌─────────────────────────┐
                      │    Healthcare Group     │
                      └────────────┬────────────┘
                                   │
                   ┌───────────────┴───────────────┐
                   ▼                               ▼
      ┌─────────────────────────┐     ┌─────────────────────────┐
      │   Hospital 1 (Main)     │     │   Hospital 2 (Westside) │
      └────────────┬────────────┘     └────────────┬────────────┘
                   │                               │
            ┌──────┴──────┐                 ┌──────┴──────┐
            ▼             ▼                 ▼             ▼
       Cardiology     Radiology        Cardiology     Pediatrics
```

- **Data Model Evolution**: A `Hospital` table can be introduced, adding a parent `hospital_id` to `Department`.
- **Zero Roster Impact**: Because all scheduling logic (shifts, doctors, constraints) is strictly bound to `department_id`, adding a `Hospital` parent requires **zero modifications** to existing CP-SAT solver algorithms, shifts, or doctor models.
