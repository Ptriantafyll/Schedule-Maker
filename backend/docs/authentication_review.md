# Authentication and Authorization Review

**Reviewed:** 2026-08-27  
**Scope:** `src/auth`, `src/user`, API route protection, JWT handling,
password handling, tenant isolation, CORS, and authentication tests.

## Verdict

The authentication setup is **not production-ready** for a hospital
scheduling API.

The project has useful foundations: bcrypt password hashing, signed JWT access
tokens, expiration timestamps, generic login failures, and database lookup of
the authenticated user. However, authentication is optional for almost every
route, authorization is not scoped by department or doctor, public signup can
create administrators, and the JWT signing key has a predictable fallback.

## Confirmed security issues

### High: public signup can create administrator accounts

**Locations:**

- `src/user/routes.py:27-30`
- `src/user/schemas.py:12-26`
- `src/user/controllers.py:16-31`
- `src/user/repository.py:13-22`
- `src/auth/dependencies.py:49-63`

`POST /api/v1/users/signup` is public. Its request model lets the caller choose
`role`, `department_id`, and `doctor_id`, and these values are persisted
without authorization or ownership checks.

An unauthenticated caller can therefore create an `admin` account and attach
it to an arbitrary department or doctor. Any route that trusts
`require_admin` will then accept that account.

**Improvements:**

1. Remove `role`, `department_id`, and `doctor_id` from the public signup
   schema.
2. Prefer disabling public signup entirely for this application.
3. Require an authenticated administrator to provision users within the
   administrator's own department.
4. Use a separate, auditable bootstrap process for the first administrator.
5. Review accounts already created through the public endpoint.

### High: middleware parses tokens but does not enforce authentication

**Locations:**

- `src/auth/middleware.py:14-25`
- `src/main.py:58-65`
- Feature route files under `src/*/routes.py`

`AuthContextMiddleware` treats a missing or invalid token as an anonymous
request and always calls the next handler. This is acceptable only when every
sensitive route explicitly requires `get_current_user` or a role dependency.
That is not the case.

At review time, only the department list route used `require_admin`. Sensitive
doctor, availability, pre-assignment, shift-assignment, team, position, and
user routes remained callable without authentication.

**Impact:** Anonymous callers can read user and doctor information and modify
scheduling data.

**Improvements:**

1. Make authentication default-deny at the API router or `include_router`
   level.
2. Explicitly mark only health and login endpoints as public.
3. Apply role-specific dependencies to write operations.
4. Prefer one standard FastAPI authentication dependency over optional
   middleware plus separate dependencies.

### High: department and object-level authorization is missing

**Locations:**

- `src/auth/dependencies.py:38-63`
- `src/user/models.py:24-29`
- `src/doctor/repository.py`
- `src/team/repository.py`
- `src/position/repository.py`
- `src/shift/repository.py`

The current dependency verifies only whether the database user has an allowed
global role. It does not enforce `department_id`, `doctor_id`, or ownership of
the requested object. Repository list and lookup functions are also global.

Requiring login alone would therefore not prevent a doctor or department
administrator from reading or changing another department's records.

**Improvements:**

1. Define a permission matrix for `admin`, `doctor`, and `viewer`.
2. Pass the authenticated user's department into service or repository
   queries.
3. Scope every lookup and list query by department unless the user is an
   explicitly defined global super-administrator.
4. Allow doctors to modify only their own availability and preferences.
5. Validate that every path and body resource belongs to the caller's allowed
   department before mutation.

### High: predictable fallback JWT signing key

**Location:** `src/auth/security.py:11-15`

If `SECRET_KEY` is not exported in the process environment, the application
silently uses:

```text
dev-secret-key-change-in-production-123456
```

An attacker who knows this source value can forge tokens for an existing user
ID. Public user-listing endpoints make those IDs easier to obtain.

**Improvements:**

1. Remove the fallback and fail application startup when the secret is absent.
2. Load settings through a validated configuration object.
3. Store the production key in a deployment secret manager.
4. Rotate the current key and invalidate previously issued tokens.
5. Use a fixed, validated signing algorithm rather than unrestricted
   environment configuration.

## Important design and implementation gaps

### Authentication router is not mounted

`src/auth/routes.py` defines `/auth/me`, but `src/main.py` does not include the
authentication router. The endpoint is therefore unreachable.

The endpoint also declares `response_model=Token` while returning a `User`
object and uses `POST` for a read operation. It should normally be
`GET /auth/me` with a safe user-profile response schema.

### OAuth2 integration is incomplete and duplicated

`oauth2_scheme` is declared in `src/auth/dependencies.py` but is unused.
Middleware manually parses the `Authorization` header instead. The login
endpoint accepts JSON, while `OAuth2PasswordBearer` conventionally points to an
OAuth2 password-form token endpoint.

**Recommendation:** Choose one approach. A simpler FastAPI design is to read
the bearer token through `oauth2_scheme` inside `get_current_user`, decode it
there, and apply that dependency globally or per router.

### Token validation is too permissive

`decode_access_token` verifies the signature and algorithm, but does not
require expected claims such as:

- `sub`
- `exp`
- `iat`
- `iss`
- `aud`
- token type or purpose

The login payload also contains the typo `doctor__id`, while the declared
`TokenPayload` does not model that field.

**Recommendation:** Define one typed claim contract and require all security
claims during decoding. Validate issuer, audience, token purpose, and subject
format.

### Broad exception handling hides token failures

`src/auth/middleware.py` catches every `Exception` and silently converts it to
an anonymous request. This can hide configuration and programming defects in
addition to expected invalid-token errors.

**Recommendation:** Catch only expected JWT exceptions and let unexpected
failures surface through normal error handling.

### Account and role relationships are not validated

All roles may have optional department and doctor links, and signup does not
check whether supplied IDs exist, are active, belong together, or are valid
for the selected role.

Suggested invariants include:

- A doctor account must reference exactly one active doctor.
- The referenced doctor and department must match.
- A department administrator must reference an active department.
- A viewer's permitted scope must be explicit.
- A global super-administrator, if required, should be a separate role.
- A doctor record should not be linked to multiple active user accounts.

### Password policy and stronger identity controls are absent

`UserCreate` and `UserLogin` accept unrestricted strings for passwords. There
is no minimum length, compromised-password check, login throttling, temporary
lockout, MFA, or external identity provider integration.

For hospital administration, prefer organizational SSO and MFA, especially
for administrators. If local passwords remain supported, enforce a modern
length policy and rate-limit login attempts by account and client.

### Token lifecycle is minimal

There is no refresh-token separation, revocation, token version, signing-key
rotation strategy, or forced logout after account or permission changes.

Reloading the user from the database on each protected request is a good start
because deleted users are rejected. Extend this with a token version or
session identifier when immediate revocation is required.

### CORS remains unrestricted

**Location:** `src/main.py:51-56`

The API permits every origin, method, and header while credentials are
enabled.

Use a configured allowlist for trusted Flutter Web origins. Bearer-header
authentication is not classic cookie CSRF, but cookie-based authentication
would also require CSRF and `SameSite` protections.

### Security audit logging is incomplete

The logging implementation avoids request bodies and tokens, which is good.
However, it does not record the authenticated subject, role, department,
login outcome, account provisioning, or privileged mutations.

Add structured, non-sensitive security events with request IDs. Do not log
passwords, JWTs, raw authorization headers, or unnecessary patient/staff data.

### User persistence API makes plaintext mistakes easy

`create_user_controller` replaces `UserCreate.password` with a bcrypt hash,
then passes the same object to `create_user`. The repository treats the
`password` field as already hashed. Tests and other internal callers can call
the repository directly and store plaintext in `hashed_password`.

Use a separate internal command model or an explicit `hashed_password`
parameter so the repository contract cannot be misunderstood.

## Correct foundations

The following parts are reasonable and should be retained:

- Passwords created through the controller are salted and hashed with bcrypt.
- Login uses the same error for a missing user and an incorrect password.
- Issued access tokens include `exp` and `iat`.
- JWT decoding restricts verification to the configured algorithm.
- Protected requests reload the user from the database.
- Deleted users are rejected by `get_current_user`.
- API response schemas do not expose `hashed_password`.

These controls are useful, but they do not compensate for public account
provisioning and mostly unprotected routes.

## Authentication test coverage

There is no dedicated `src/auth/tests` suite. The current user tests do not
adequately validate authentication or authorization.

### Missing tests

1. Successful login with a bcrypt-hashed password.
2. Incorrect password and unknown-user responses.
3. Deleted-user login rejection.
4. JWT creation and required claims.
5. Expired, malformed, tampered, wrong-algorithm, wrong-issuer, and
   wrong-audience tokens.
6. Missing bearer credentials returning `401`.
7. Invalid credentials returning `401` rather than anonymous access.
8. Insufficient role returning `403`.
9. Public signup being unable to select an administrator role.
10. Cross-department reads and writes being denied.
11. Doctors being unable to change another doctor's availability.
12. Deleted or mismatched department and doctor links being rejected.
13. `/auth/me` returning the correct safe profile.
14. CORS allowing trusted origins and rejecting untrusted origins.
15. Login throttling or lockout behavior, once implemented.

### Problems in current tests

- The signup route test targets `/api/v1/users`, but the implementation exposes
  `/api/v1/users/signup`.
- The signup test submits and expects `role="admin"`, which encodes the
  privilege-escalation flaw as expected behavior.
- Repository tests only assert that `hashed_password` is a string; they do not
  prove it is a bcrypt hash or that plaintext verification works.
- Existing feature route tests generally omit authorization headers and
  expect success, reinforcing default-open API behavior.
- The only protected department-list route test remains unfinished.

## Validation status

All auth, user, and main modules passed Python syntax compilation.

Targeted pytest collection could not start because the existing virtual
environment did not contain the declared `PyJWT` and `email-validator`
packages. The dependencies are present in `pyproject.toml`, so synchronize the
approved local environment and run tests with the repository's documented
`uv run --no-sync pytest ...` workflow afterward.

## Recommended implementation order

1. Close public administrator signup and establish a secure bootstrap path.
2. Make all API routes authenticated by default.
3. Implement department and object-level authorization.
4. Remove the fallback signing key and strengthen JWT validation.
5. Add authentication and cross-tenant authorization tests.
6. Fix and mount `/auth/me`.
7. Restrict CORS.
8. Add login throttling, MFA or SSO, token lifecycle controls, and audit
   events.
