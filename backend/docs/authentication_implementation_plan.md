# Authentication Remediation Plan

## Problem and proposed approach

The backend has working password hashing and basic JWT creation, but its
security boundary is incomplete:

- Public signup can choose the `admin` role and arbitrary resource links.
- Almost all domain routes are anonymous.
- Authentication is split between passive middleware and an unused OAuth2
  dependency.
- Authorization checks roles globally instead of limiting access by
  department and doctor ownership.
- There is no invitation workflow, refresh-session storage, rotation,
  revocation, or logout lifecycle.
- Authentication tests do not currently prove these boundaries.

The implementation should first close anonymous and privilege-escalation
paths, then establish department-scoped authorization, and only then add
doctor invitation signup and rotating refresh sessions.

## Working agreement

- The user writes every repository change.
- The assistant reviews each test change and implementation change.
- Do not move to the next phase until the current phase's tests and behavior
  have been reviewed.
- Follow TDD for each phase:
  1. Write the failing security or behavior tests.
  2. Send the test diff for review.
  3. Write the minimum implementation that satisfies those tests.
  4. Send the implementation diff and test output for review.
- Shared test fixtures should be extracted only when the first auth tests need
  them, rather than through a large unrelated test refactor.

## Confirmed design decisions

- Scope is authentication, authorization, user identity, and related tests.
- Doctors may self-register only through a one-time invitation issued by
  their department administrator.
- `admin` means a department administrator scoped to one department.
- `super_admin` is a global control-plane role.
- Super-admins may create and manage departments and department-admin
  accounts, and may perform audited global reads.
- Super-admins must not perform routine doctor, availability, shift, or
  schedule mutations.
- Super-admin accounts are created only by a local administrative bootstrap
  command and never through an API endpoint.
- Departments and their first department administrators are provisioned
  through protected super-admin API endpoints.
- Bearer tokens are extracted by FastAPI dependencies.
- `AuthContextMiddleware` will be removed.
- Access tokens remain JWTs.
- Refresh tokens are opaque random values, stored only as hashes, and rotated
  on every use.
- Refresh tokens use a Secure, HttpOnly cookie for web and native clients.
  Native Flutter clients must use a persistent cookie jar.
- Cookie-backed refresh and logout requests use CSRF protection.
- The development database is disposable, so no data migration is required
  for this phase.
- Breaking authentication endpoint changes are acceptable.
- Secret storage, key rotation, and the current loose secret configuration are
  explicitly deferred at the user's request.

## Target route layout

### Public endpoints

| Method | Endpoint | Purpose |
|---|---|---|
| `GET` | `/health` | Health check |
| `POST` | `/api/v1/auth/login` | Verify credentials, return access token, set refresh and CSRF cookies |
| `POST` | `/api/v1/auth/signup` | Consume a doctor invitation and create a doctor account |
| `POST` | `/api/v1/auth/refresh` | Rotate the refresh session and issue a new access token |
| `POST` | `/api/v1/auth/logout` | Revoke the refresh-session family and clear cookies |

`signup`, `refresh`, and `logout` are public only in the sense that they do
not require a valid access token. They still require their invitation or
cookie/CSRF credentials.

### Authenticated endpoints

| Area | Doctor | Viewer | Department admin | Super-admin |
|---|---|---|---|---|
| `/auth/me` | Own account | Own account | Own account | Own account |
| Department details | Own department | Own department | Own department | Create/manage and read globally |
| Teams, positions, shifts | Read own department | Read own department | Read/write own department | Audited global read; no routine write |
| Published shift assignments | Read own department | Read own department | Read/write own department | Audited global read; no routine write |
| Doctor roster | Read own department as required by scheduling UX | Read own department | Read/write own department | Audited global read; no routine write |
| Doctor unavailability | Read/write own only | No raw access | Read/write doctors in own department | Audited global read only |
| Doctor pre-assignments | Read own or department schedule view | Read own department if required | Read/write own department | Audited global read only |
| User records | Own account through `/auth/me` | Own account through `/auth/me` | List/read users in own department | Read globally and manage department admins |
| Invitations | None | None | Create/list/revoke doctor invitations in own department | Create/list/revoke department-admin invitations |

### Department creation

Department creation is a protected super-admin control-plane operation.
Provisioning a department should atomically create the department and a
one-time invitation for its first department administrator. The super-admin
must not set or learn that administrator's password. Department
administrators must not create or enumerate other departments.

## Error and information-disclosure rules

- Missing or invalid access credentials return `401` with
  `WWW-Authenticate: Bearer`.
- An authenticated role attempting an operation that its role can never
  perform returns `403`.
- A resource outside the caller's department should normally appear as `404`
  because scoped repository queries should not return it.
- A doctor attempting to modify a different doctor ID may return `403`
  because the ownership violation is known after authentication.
- A super-admin attempting a routine scheduling mutation returns `403`.
- Login uses the same response for an unknown email and an incorrect password.
- Never log passwords, access tokens, refresh tokens, invitation tokens, raw
  cookies, authorization headers, or CSRF values.

## Phase 1: Establish the authentication test foundation

### Files

- Add `src/auth/tests/`.
- Add or extend `backend/conftest.py`.
- Update `src/user/tests/test_user.py`.

### Test fixtures to introduce

- In-memory SQLModel session that registers all application models.
- FastAPI client with the session dependency overridden.
- Factories for departments, doctors, users, and password-hashed users.
- Helpers that create a valid access token and Authorization header.
- Factories for users in two distinct departments.

Do not create several role-specific fixtures when one user factory with clear
arguments will keep tests simpler.

### First failing tests

1. A protected route returns `401` without a bearer token.
2. A malformed or tampered bearer token returns `401`.
3. A deleted user represented by an otherwise valid token returns `401`.
4. A doctor calling an admin-only operation returns `403`.
5. A department admin can call the same operation.
6. A department admin calling a super-admin control-plane operation returns
   `403`.
7. A super-admin can call a control-plane operation.
8. A super-admin cannot perform a routine scheduling mutation.
9. The public signup request cannot submit `role`, `department_id`, or
   `doctor_id`.
10. The existing user route test uses the real route rather than
   `/api/v1/users`.
11. Passwords created through the intended service/controller path are bcrypt
   hashes and verify against the original plaintext.

### Review gate

Send only the new fixtures and failing tests for review. Confirm that at least
one test fails for each current security defect rather than failing because
of an unrelated import or fixture problem.

### Done when

- Auth tests collect successfully.
- Each test has one clear security expectation.
- Broad `pytest.raises(Exception)` is not used for auth behavior.

## Phase 2: Simplify access-token authentication

### Files

- `src/auth/dependencies.py`
- `src/auth/security.py`
- `src/auth/schemas.py`
- `src/auth/controllers.py`
- `src/auth/routes.py`
- `src/user/controllers.py`
- `src/user/routes.py`
- `src/main.py`
- Remove `src/auth/middleware.py` after references are gone.

### Changes

1. Make `oauth2_scheme` the single source of bearer tokens.
2. Change `get_current_user` to receive the token from
   `OAuth2PasswordBearer`, decode it, validate its typed payload, load the
   active user, and return the database user.
3. Catch only expected JWT validation exceptions.
4. Stop reading `request.state.user`.
5. Remove `AuthContextMiddleware` registration and then remove the middleware
   module.
6. Keep access-token claims minimal:
   - subject user ID;
   - issued-at time;
   - expiration time;
   - unique token ID;
   - issuer;
   - audience;
   - token type equal to `access`.
7. Do not trust token copies of role, email, department, or doctor ID for
   authorization. Load current values from the database.
8. Add a `require_super_admin` dependency that checks the current database
   role.
9. Move login behavior from the user feature into the auth feature.
10. Mount the auth router from `src/main.py`.
11. Replace the current `/auth/me` declaration with a read endpoint returning
    a safe user-profile schema rather than a token schema.
12. Move login to `/api/v1/auth/login`.
13. Remove the old `/api/v1/users/login` route because breaking changes are
    allowed.

Do not implement refresh tokens in this phase. First make access-token
authentication coherent and testable.

### Tests

- Successful and failed login.
- Expired token.
- Tampered token.
- Missing required claim.
- Incorrect issuer, audience, or token type.
- Invalid subject UUID.
- Deleted user.
- Correct `/auth/me` response with no password hash.

### Review gate

Send the dependency/security changes and targeted auth test output for review
before applying authentication to every domain router.

### Done when

- No authentication code depends on middleware request state.
- `/auth/login` and `/auth/me` behave correctly.
- Expected invalid tokens produce `401`; unexpected failures are not silently
  converted to anonymous requests.

## Phase 3: Add a safe super-admin bootstrap path

### Files

- Add a command under `scripts/`, for example a super-admin bootstrap command.
- Reuse user/domain services instead of writing directly to database fields
  where possible.
- Add focused command/service tests.

### Required behavior

1. Accept super-admin identity details.
2. Read the password interactively without placing it in command history.
3. Derive the role as `super_admin`; never accept an arbitrary role.
4. Ensure a super-admin has neither `department_id` nor `doctor_id`.
5. Refuse a duplicate email address with a clear non-zero result.
6. Allow additional super-admins only through the same local command, never
   through an HTTP endpoint.
7. Be explicit that rerunning the command is not an idempotent success unless
   the existing records exactly match a deliberately supported recovery case.
8. Do not print password hashes or tokens.

Because the database is disposable, update model registration and recreate
the local database after the later schema phases rather than adding migrations
now.

### Review gate

Review the command's role derivation and password-handling behavior before
public signup is removed.

### Done when

- A trusted super-admin can be created without an HTTP privilege-escalation
  path.
- No API caller can create or promote a super-admin.

## Phase 4: Make the API authenticated by default

### Files

- `src/user/routes.py`
- `src/department/routes.py`
- `src/doctor/routes.py`
- `src/team/routes.py`
- `src/position/routes.py`
- `src/shift/routes.py`
- `src/main.py`

### Changes

1. Remove the existing public signup endpoint until the invitation flow is
   ready.
2. Remove arbitrary public user lookup by email.
3. Require `get_current_user` for every domain route.
4. Require `require_admin` for all domain writes except a doctor's own
   unavailability changes.
5. Restrict user listing and lookup to department admins.
6. Reserve department creation and department-admin provisioning for
   `require_super_admin`; do not enable those operations until the invitation
   flow is complete.
7. Keep only the public endpoints listed in the target route layout.
8. Avoid relying solely on router-level dependencies where the handler needs
   the current user for scope checks; inject the current user into those
   handlers explicitly.

### Tests

For each router, add a small authorization matrix rather than duplicating
every CRUD test:

- anonymous request -> `401`;
- doctor on admin write -> `403`;
- same-department admin -> reaches business logic;
- department admin on a super-admin endpoint -> `403`;
- super-admin on a routine scheduling write -> `403`;
- public health and login remain accessible.

Existing feature route tests must start using authenticated clients. Do not
weaken the new default-deny behavior merely to keep old anonymous tests green.

### Review gate

Review route dependencies feature by feature. Confirm there is no mounted
domain route without an authentication or role dependency.

### Done when

- Anonymous access is impossible outside the explicit public endpoints.
- Public signup can no longer create any account.

## Phase 5: Implement department and object-level authorization

Role checks alone are insufficient. Complete this phase one feature at a time
and stop for review after each subphase.

### Shared approach

1. Route dependencies return the current database user.
2. Controllers determine the allowed operation and pass an explicit
   `department_id` or ownership key to repositories.
3. Repositories query within that scope instead of retrieving globally and
   filtering afterward.
4. Cross-department IDs therefore resolve to `None` and become `404`.
5. Do not put HTTP-specific logic inside repositories.
6. Tenant-owned create bodies do not accept `department_id`; controllers
   derive it from the authenticated department admin.
7. Use explicit scoped repository methods. Do not make a missing or optional
   department ID mean global access.

### Phase 5A: User and department scope

**Files:** `src/user/*`, `src/department/*`, related tests.

- Admin user queries return only users in the current admin's department.
- Super-admin user queries use explicit global control-plane repository
  methods rather than bypassing scope flags.
- `/auth/me` is the only normal self-profile endpoint.
- Department reads return only the caller's department.
- Only super-admin endpoints may list all departments.
- Super-admins may create, disable, and inspect departments and department
  administrators, but not routine scheduling data mutations.
- Validate that department admins have a department ID.
- Validate that super-admins have no department or doctor link.

### Phase 5B: Doctor ownership and department scope

**Files:** `src/doctor/*`, related tests.

- Roster queries are limited to the caller's department.
- Full doctor list/detail responses, including contact email, are
  department-admin-only.
- Department members use a separate reduced `/doctors/roster` response that
  excludes contact email.
- Admins may manage doctors only in their own department.
- Doctors may read and change only their own raw unavailability.
- A doctor targeting another doctor's private data receives `403`.
- Viewers cannot read raw unavailability.
- Pre-assignments and position assignments validate all related resources
  within one department.
- A doctor account without a valid `doctor_id` is rejected as an invalid
  account state, not granted broader access.

### Phase 5C: Teams

**Files:** `src/team/*`, related tests.

- Reads are limited to the current department.
- Only the department admin may create or change teams.
- Team create bodies do not accept a department ID; derive it from the admin.
- Team names are unique within a department, including soft-deleted rows.

### Phase 5D: Positions

**Files:** `src/position/*`, related tests.

- Reads are limited to the current department.
- Only the department admin may create or change positions.
- Position create bodies do not accept a department ID; derive it from the
  admin.
- Position detail routes use UUIDs rather than names.
- Position names are unique within a department, including soft-deleted rows.

### Phase 5E: Shifts and assignments

**Files:** `src/shift/*`, related tests.

- Shift definitions and assignments are scoped through the position's
  department.
- Only the department admin may create shifts or assignments.
- Assigned doctors must belong to the same department and be active.
- List endpoints return only assignments from the current department.
- Shift detail routes use UUIDs rather than names.
- Shift names are unique within a position, including soft-deleted rows.

### Required cross-tenant tests

For each feature:

1. Create department A and department B.
2. Authenticate as a user from department A.
3. Attempt to read and mutate department B's object.
4. Assert the request does not reveal or modify the object.
5. Verify the database record remains unchanged after rejected writes.

### Review gate

Review each subphase independently. The assistant should trace every resource
ID from route to repository and verify that all parent resources are scoped.

### Done when

- Authentication from one department cannot observe or alter another
  department's records.
- Doctors cannot alter another doctor's protected data.

## Phase 6: Strengthen user-model and repository contracts

Detailed execution plan:

```text
docs/authentication_phase_6_progress.md
```

### Files

- `src/user/models.py`
- `src/user/schemas.py`
- `src/user/repository.py`
- `src/user/controllers.py`
- Add `src/user/services.py`.
- Related tests and database initialization.

### Changes

1. Separate public input DTOs from internal account-creation data.
2. Stop mutating `UserCreate.password` into a hash.
3. Make repository creation accept an explicitly named hashed-password value
   or a fully prepared internal model.
4. Normalize login email consistently before lookup and storage.
5. Enforce case-insensitive uniqueness for normalized emails.
6. Make `doctor_id` unique for active doctor accounts so one doctor cannot be
   linked to multiple accounts.
7. Validate role relationships:
   - doctor -> active doctor and matching department;
   - admin -> active department and optional active doctor in the same
     department;
   - super-admin -> no department or doctor link;
   - viewer -> explicit department and no doctor link unless a use case says
     otherwise.
8. Decide whether the currently unused viewer role should remain. If retained,
   keep it read-only and provide no public provisioning path in this phase.

### Confirmed Phase 6 decisions

- Store one trimmed lowercase canonical login email.
- Keep login emails globally reserved after soft deletion.
- Allow a Doctor to have no User or one active User; allow a replacement User
  after the prior account is soft-deleted.
- Allow a Department admin to hold the single active User link for a Doctor in
  the same Department.
- Keep Department-admin authentication active if its optional Doctor link
  later becomes invalid; do not trust that link for Doctor identity.
- Do not assign Doctor links through Phase 7 Department-admin invitations;
  defer explicit linking and unlinking to account management.
- Retain the Department-scoped read-only viewer role.
- Reject malformed accounts and accounts whose required links are missing,
  deleted, or mismatched during current-user loading with `401`.
- Add a reusable account service and a non-committing staging operation so
  Phase 7 invitation consumption can create the User and consume the
  invitation atomically.
- Enforce role/link nullability with a database `CHECK` constraint and enforce
  one active User per Doctor with a partial unique index.

### Tests

- Repository callers cannot accidentally persist plaintext.
- Email case variants cannot create duplicate accounts.
- One doctor cannot receive two active accounts.
- Invalid role/department/doctor combinations are rejected.
- Soft-deleted linked resources are rejected.

### Review gate

Review model invariants and repository signatures before adding invitations,
because the invitation flow will depend on them.

### Done when

- Account creation has one safe internal path.
- User-role relationships cannot enter an ambiguous state through normal
  application flows.

## Phase 7: Add scoped account invitations and super-admin provisioning

### Files

- Add `src/auth/models.py`.
- Extend `src/auth/schemas.py`.
- Add invitation persistence to `src/auth/repository.py`.
- Add invitation business logic to `src/auth/controllers.py`.
- Add protected invitation routes.
- Register auth models during database initialization.
- Add invitation tests.

#### Invitation model

Store:

- invitation ID;
- token hash, never the raw token;
- invitation role (`DEPARTMENT_ADMIN`, `DOCTOR`, `VIEWER`);
- intended full name / display name;
- optional target doctor ID (required for `DOCTOR`, null for others);
- derived department ID;
- intended login email;
- creator admin user ID;
- expiration time;
- used time;
- revoked time;
- normal creation/update timestamps.

Use a high-entropy opaque token (`secrets.token_urlsafe(32)`). A fast cryptographic digest (`SHA-256`) is stored in the database so database leaks cannot compromise unconsumed invitations.

### Department-admin staff invitation flow (Doctors and Viewers)

1. Authenticate and require department admin.
2. For Doctor invitation:
   - Load doctor through department-scoped query; reject deleted doctors or doctors in other departments.
   - Reject if doctor already has an active user account or active pending invite.
   - Derive `full_name` from `doctor.name` and role as `DOCTOR`.
3. For Viewer invitation:
   - Accept intended `email` and `full_name`.
   - Set `doctor_id = None` and role as `VIEWER`.
4. Normalize and reserve the intended login email.
5. Create the invitation and return the raw token / link exactly once.
6. Do not log the token.
7. Add list and revoke operations that reveal metadata but never the token.
8. Email delivery is out of scope for now (documented in `docs/backlog.md` BL-002); tokens are shared out-of-band via copyable link.

### Super-admin department provisioning flow

1. Authenticate and require super-admin.
2. Accept department details (`name`, `code`) and the intended first administrator's login email and display name.
3. Create the department and an invitation with role `DEPARTMENT_ADMIN` atomically.
4. Never accept a caller-selected role.
5. Do not accept a password for the future department administrator.
6. Return the raw invitation token exactly once and never log it.
7. Permit super-admins to list, revoke, and reissue department-admin invitations for new or existing departments.
8. Keep routine team, doctor, shift, and schedule writes unavailable to the super-admin.

### Public signup flow

1. Accept only invitation token and password (`POST /api/v1/auth/signup`).
2. Hash the presented invitation token (`SHA-256`) and load one active, unused, unexpired, unrevoked invitation.
3. Recheck the doctor (if doctor invite), department, email uniqueness, and absence of an existing active doctor account inside the creation transaction.
4. Derive identity immutably:
   - role from the trusted invitation;
   - doctor ID from a doctor invitation, or None for admin/viewer;
   - department ID from the invitation;
   - login email from invitation;
   - display/full name from invitation (or doctor record if doctor invite).
5. Call safe `user_services.stage_user_account` and mark the invitation used atomically.
6. Return the created safe user and/or login authentication token.

The doctor contact email and user login email are not required to be equal.
The invitation explicitly binds the intended login identity to the doctor.

### Tests

- Department admin A cannot invite a doctor or viewer from department B.
- Doctor or viewer cannot issue invitations.
- Department admins cannot issue admin or super-admin invitations.
- Super-admins can issue department-admin invitations but not doctor invitations.
- No invitation may create a super-admin.
- Raw tokens are never stored.
- Expired, used, revoked, malformed, and unknown invitations fail.
- Caller-supplied role or resource IDs are rejected by schema validation.
- Concurrent or repeated invitation consumption creates at most one account.
- Signup cannot change the invitation email, doctor, department, or role.

### Review gate

First review the model and failing tests, then the admin invitation flow, then
the public signup transaction.

### Done when

- Doctor, viewer, and department-admin registration cannot select identity, role, or tenant scope.
- Each invitation is single-use, cryptographically secured, and auditable.
- Departments cannot be left committed without their first admin invitation when provisioning fails.

## Phase 8: Add refresh-session persistence

### Files

- `src/auth/models.py`
- `src/auth/repository.py`
- `src/auth/schemas.py`
- Database registration and tests.

### Refresh-session model

Store:

- session ID;
- user ID;
- refresh-token hash;
- session-family ID;
- CSRF-token hash;
- expiration time;
- last-used time;
- revoked time and reason;
- replacement session ID;
- creation metadata.

Keep old rotated records long enough to detect reuse. Add indexes for token
hash, user ID, family ID, and expiration cleanup.

### Repository operations

- Create session.
- Find session by refresh-token hash.
- Rotate one active session in one transaction.
- Revoke one session family.
- Revoke all sessions for a user.
- Delete expired history according to an explicit retention policy.

### Tests

- Only token hashes are stored.
- Expired and revoked sessions are inactive.
- Rotation revokes/replaces exactly one session.
- Reuse of a replaced token can still be detected.
- Family revocation affects all related sessions.

### Review gate

Review persistence and transactional behavior before exposing refresh routes.

### Done when

- The database can represent rotation, reuse detection, logout, and global
  user-session revocation without storing raw credentials.

## Phase 9: Implement login, refresh rotation, logout, and CSRF

### Files

- `src/auth/security.py`
- `src/auth/controllers.py`
- `src/auth/routes.py`
- `src/auth/schemas.py`
- Auth tests.

### Login behavior

1. Verify normalized email and password.
2. Issue a short-lived access JWT.
3. Generate a high-entropy opaque refresh token.
4. Store only its hash in a new refresh session.
5. Generate a CSRF value and store its hash with the session.
6. Set the refresh token in an HttpOnly cookie.
7. Set or return the companion CSRF value so the client can send it in an
   `X-CSRF-Token` header.
8. Return the access token and safe user/session metadata.

### Cookie requirements

- `HttpOnly` for the refresh cookie.
- `Secure` outside explicit local HTTP development.
- Restrictive `SameSite`.
- Narrow path covering only auth refresh/logout endpoints.
- Explicit maximum age matching refresh-session expiration.
- No broad parent-domain cookie unless deployment requires it.

### Refresh behavior

1. Read the refresh cookie.
2. Require and validate the CSRF header.
3. Hash and load the refresh session.
4. Reject missing, expired, revoked, or mismatched sessions.
5. Reload the active user.
6. Rotate refresh token and CSRF value atomically.
7. Mark the old session as replaced.
8. Set replacement cookies and return a new access token.
9. If a previously replaced token is reused, revoke the entire family and
   reject the request.

### Logout behavior

1. Validate refresh cookie and CSRF value when present.
2. Revoke the current session family.
3. Clear refresh and CSRF cookies even when the session is already invalid.
4. Do not require the access token to still be valid, because logout must work
   after access-token expiry.

### Tests

- Cookie flags and path.
- Refresh succeeds once and rotates both credentials.
- Old refresh token fails after rotation.
- Reuse revokes the complete family.
- Missing/wrong CSRF value fails.
- Expired or deleted user fails and revokes the session.
- Logout revokes and clears cookies.
- Multiple device families remain independent until all-user revocation.

### Review gate

Review login cookie behavior, then refresh rotation/reuse detection, then
logout. Include database assertions, not only HTTP status assertions.

### Done when

- Access-token expiry does not force re-login while an active refresh session
  remains.
- A stolen old refresh token cannot be reused silently.
- Refresh credentials are not exposed to Flutter application code.

## Phase 10: Restrict CORS and add authentication audit events

### Files

- `src/main.py`
- `src/utils/logger.py`
- Auth controllers/dependencies.
- Tests.

### CORS

1. Load an explicit set of trusted Flutter Web origins.
2. Keep credential support because refresh cookies are used.
3. Limit methods to those used by the API.
4. Limit headers to required values such as `Authorization`,
   `Content-Type`, and `X-CSRF-Token`.
5. Test a trusted origin and an untrusted origin.

### Audit events

Record structured events for:

- login success and failure;
- invitation creation, revocation, and consumption;
- super-admin bootstrap and department provisioning;
- audited super-admin global reads;
- refresh success and reuse detection;
- logout;
- role denial;
- cross-tenant access denial;
- user or session revocation.

Include request ID, user ID when known, role, department ID, event result, and
safe reason codes. Do not include raw identifiers when they are unnecessary
for operational investigation.

### Review gate

Review log fields specifically for sensitive-data leakage and excessive PII.

### Done when

- Browser origins are explicitly controlled.
- Security-relevant actions can be investigated without logging credentials.

## Phase 11: Abuse protection and password policy

### Changes

1. Add a password length policy suitable for bcrypt and reject known weak
   values according to the chosen product policy.
2. Rate-limit login, invitation consumption, and refresh failures.
3. Use generic external error messages while retaining safe internal reason
   codes.
4. Define temporary lockout or progressive delay behavior without allowing an
   attacker to permanently lock another account.
5. Consider organizational SSO and MFA as a later replacement or enhancement,
   especially for department administrators.

### Tests

- Weak passwords are rejected.
- Repeated login failures trigger the chosen control.
- Successful authentication resets or reduces failure state.
- Rate limits do not reveal whether an email exists.

### Review gate

Review denial-of-service and account-enumeration implications before enabling
lockout behavior.

### Done when

- Online password guessing is bounded.
- Password rules are explicit and tested.

## Phase 12: Cleanup and final validation

### Cleanup

1. Delete obsolete auth middleware and unused auth imports.
2. Remove old user login/signup routes and stale tests.
3. Remove or rewrite `src/auth/controllers.py` and
   `src/auth/repository.py` stubs as applicable.
4. Ensure `TokenPayload` matches the actual access-token claims.
5. Remove the `doctor__id` typo rather than retaining compatibility.
6. Ensure all auth models are registered before database creation.
7. Recreate the disposable local SQLite database after schema changes.
8. Update API documentation and Flutter integration notes.
9. Keep secret cleanup explicitly deferred, but record it as unresolved risk.

### Validation

Run targeted tests after every phase and the full backend suite at the end.
Use the repository's local workflow:

- `uv run --no-sync pytest src/auth/tests src/user/tests`
- feature-specific authorization tests while scoping each feature;
- `uv run --no-sync pytest`

Also inspect the generated OpenAPI document and verify:

- public endpoints are intentionally public;
- protected endpoints show bearer authentication;
- obsolete routes are absent;
- request schemas cannot accept role or tenant identity during signup;
- cookie and CSRF behavior is documented for Flutter Web and native clients.

### Final manual threat checks

1. Anonymous user attempts every mounted domain route.
2. Doctor attempts every admin mutation.
3. Department A admin uses department B IDs in paths and bodies.
4. Doctor A tries to alter doctor B's unavailability.
5. Attacker reuses a rotated refresh token.
6. Untrusted browser origin attempts credentialed refresh.
7. Deleted user tries access and refresh tokens.
8. Department admin attempts a super-admin control-plane operation.
9. Super-admin attempts a routine scheduling mutation.
10. API caller attempts to create or promote a super-admin.

## Deferred items

These are intentionally not implementation todos in this plan:

- Replacing and rotating the current loose signing secret.
- Moving secrets into a production secret manager.
- Full key-rotation support.
- Organizational SSO and MFA implementation.
- Alembic migrations, because the current development database is disposable.
- Email delivery for invitation tokens.

They remain important before production deployment, especially secret
management and administrator MFA.
