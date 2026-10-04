# Backend Backlog

Deferred issues that should be addressed outside the current implementation step.

## Open

### BL-002: Automated email dispatch for invitations (SMTP / SendGrid)

**Area:** Authentication / Notifications  
**Priority:** Medium

#### Problem

In the initial implementation of Phase 7 invitations, invitation links/tokens are returned directly to the admin (via the API / frontend dashboard) to be copied and shared manually out-of-band. While secure and practical for testing and desktop use, enterprise deployment will benefit from automated email delivery directly to the invited doctor or department administrator.

#### Required work

1. Integrate an email service provider client (e.g. SMTP, SendGrid, Amazon SES, or Mailgun) with credentials managed via environment variables.
2. Design responsive email templates for:
   - Department Administrator invitation.
   - Doctor invitation.
   - Viewer invitation.
3. Include deep-links containing the one-time invitation token pointing to the web/mobile signup page.
4. Implement background task execution (FastAPI `BackgroundTasks` or Celery/Redis) so email dispatch does not block the HTTP request.
5. Provide delivery retry and failure logging without leaking the raw token into log files.

#### Completion criteria

- Creating an invitation can trigger an asynchronous email send.
- Email delivery failure does not roll back an already-created invitation; admins can still view/copy the link manually as a fallback.
- Test suite provides mock email transport verification.

### BL-003: Multi-hospital / healthcare network hierarchy

**Area:** Multi-tenancy & Data modeling  
**Priority:** Low

#### Problem

Currently, the system is multi-tenant at the `Department` level. Multiple departments can live side-by-side in the same database with full tenant isolation. In large enterprise deployments, healthcare networks may operate multiple physical hospitals, clinics, or campuses (e.g. "St. Jude General Hospital" and "Westside Clinic"), each containing multiple clinical departments.

#### Required work

1. Introduce a parent `Hospital` (or `Organization`) model:
   - Fields: `id`, `name`, `code`, `address`, etc.
2. Add an optional or foreign key `hospital_id: UUID` to the `Department` table.
3. Update department provisioning so super-admins can group departments by hospital.
4. Optionally introduce a `HOSPITAL_ADMIN` role that can manage departments and view high-level schedules across an entire hospital facility without having root `SUPER_ADMIN` system access.

#### Completion criteria

- Departments can be associated with a parent `Hospital`.
- Department-level shift solving, rosters, and schedules remain fully isolated and unaffected.

### BL-004: Self-service department request & approval queue

**Area:** Account onboarding  
**Priority:** Low

#### Problem

Phase 7 implements direct Super Admin department provisioning (Approach A). If the platform transitions to a self-serve SaaS model, prospective department heads may want to submit a department registration request publicly, which remains pending until reviewed and approved by a Super Admin.

#### Required work

1. Create a public endpoint `POST /api/v1/department-requests` accepting department details, head doctor contact info, and registration justification.
2. Store requests in a `department_request` table with status `PENDING`, `APPROVED`, or `REJECTED`.
3. Provide Super Admin endpoints to list pending requests and approve/reject them.
4. Upon approval, automatically trigger the department provisioning and admin invitation flow.

#### Completion criteria

- Prospective admins can submit requests without instant activation.
- Super Admins can audit, approve, or reject requests from an admin dashboard.

### BL-005: Department Admin delegation and role transfer

**Area:** User & Role Management / Administration  
**Priority:** Medium

#### Problem

Currently, `DEPARTMENT_ADMIN` accounts are provisioned exclusively via Super-Admin department provisioning or admin invitation tokens. A Department Admin cannot directly appoint an existing department member (e.g. a Doctor or Viewer in their department) to become a Department Admin.

In hospital environments, chief physicians and department heads frequently rotate, step down, or share administrative scheduling duties with a deputy or head nurse. Department Admins need the ability to appoint a colleague as a Department Admin (or hand over leadership).

#### Required work

1. Create a service and endpoint (e.g. `POST /api/v1/departments/me/admins/appoint` or `PATCH /api/v1/users/{user_id}/role`):
   - Only callable by an active `DEPARTMENT_ADMIN` (or `SUPER_ADMIN`).
   - Strictly scoped: the target user must belong to the caller's `department_id`.
2. Support two delegation modes or policies:
   - **Co-Admin Appointment**: Elevate a target Doctor or Viewer to `DEPARTMENT_ADMIN` while keeping the current admin active.
   - **Role Handover (Succession)**: Promote the target member to `DEPARTMENT_ADMIN` and atomically demote the current admin to `DOCTOR` or `VIEWER`.
3. Check and adjust role-shape invariants if needed:
   - If a Doctor is promoted to `DEPARTMENT_ADMIN`, verify whether their `doctor_id` link is retained (allowing them to remain on clinical duty rosters while having admin access).
4. Add audit logging and email notification upon role change.

#### Completion criteria

- Department Admins can appoint another member in their department to `DEPARTMENT_ADMIN`.
- Attempts to appoint members from other departments or promote across tenant boundaries are rejected with HTTP 403 / 404.
- Invariants on `User` and `Department` are verified and tested.

### BL-006: Self-service staff registration with Department Admin approval queue

**Area:** User Onboarding / Authentication  
**Priority:** Medium

#### Problem

Currently, staff members (Doctors and Viewers) must be invited explicitly by a Department Admin via an out-of-band invitation token. In larger hospital environments or open-onboarding scenarios, new staff members may prefer to register directly on the public site by selecting their department and role without waiting for a pre-issued invitation token.

However, open public registration cannot activate accounts immediately without compromising tenant isolation and roster integrity. A self-registration request must enter a pending approval state awaiting explicit review and acceptance by a Department Admin of that department.

#### Required work

1. Create a public endpoint `POST /api/v1/auth/registration-requests`:
   - Accepts applicant details: name, email, password, target department, requested role (`DOCTOR` or `VIEWER`), and optional clinical notes/employee ID.
   - Stores the request with status `PENDING` (e.g. in a dedicated `user_registration_request` table or staging status).
2. Provide Department Admin management endpoints:
   - `GET /api/v1/departments/me/registration-requests`: List pending applicant requests for their department.
   - `POST /api/v1/departments/me/registration-requests/{id}/approve`: Approves the request, creates/activates the `UserModel`, optionally links to an existing doctor or auto-provisions a `DoctorModel`, and notifies the applicant.
   - `POST /api/v1/departments/me/registration-requests/{id}/reject`: Rejects the request with an optional reason.
3. Integrate with email notifications (see BL-002) to alert admins of new registration requests and notify applicants of approval/rejection.

#### Completion criteria

- Staff members can submit self-registration requests publicly.
- Unapproved requests cannot authenticate or access protected routes.
- Department Admins can review, link to doctors, and approve/reject requests exclusively within their own department.

### BL-007: Multi-Factor Authentication (MFA / TOTP) and OAuth2 SSO (Google Login)

**Area:** Security & Identity  
**Priority:** Medium

#### Problem

Hospital systems handle sensitive clinical scheduling and physician on-duty rosters. Relying solely on username/password authentication poses security risks (credential stuffing, weak passwords).

Furthermore, hospital staff and physicians frequently rely on institutional identity providers (e.g., Google Workspace, Microsoft Entra ID) for Single Sign-On (SSO). Adding Google OAuth2 login and Multi-Factor Authentication (MFA) will enhance enterprise compliance and streamline physician access.

#### Required work

1. **OAuth2 / Google Sign-In Integration**:
   - Integrate Google OAuth2 / OpenID Connect (OIDC) authorization code flow with PKCE.
   - Add routes `GET /api/v1/auth/oauth/google/authorize` and `GET /api/v1/auth/oauth/google/callback`.
   - Match verified Google email addresses to active accounts or route them through the invitation/approval workflow.
2. **Time-Based One-Time Password (TOTP) MFA**:
   - Support TOTP (RFC 6238) using standard authenticator apps (Google Authenticator, Authy).
   - Endpoints for MFA enrollment: QR code generation and verification of the initial TOTP code.
   - Store encrypted TOTP secret keys at rest on `UserModel`.
   - Update authentication flow: if a user has MFA enabled, validating email/password returns a short-lived challenge token (`HTTP 202 Accepted` / `mfa_required`) requiring a secondary code submission to `POST /api/v1/auth/mfa/verify` before issuing access tokens.
   - Provide secure one-time backup recovery codes.

#### Completion criteria

- Users can authenticate seamlessly using "Sign in with Google".
- Users and administrators can enable TOTP MFA on their accounts.
- Accounts with MFA enabled cannot obtain access tokens without a valid second factor.

### BL-008: Distributed Rate Limiting for Multi-Instance / Kubernetes Deployments (Redis)

**Area:** Security, Infrastructure & Scalability  
**Priority:** Medium

#### Problem

The initial rate limiting implementation uses an in-memory sliding window (`InMemoryRateLimiter`). While optimal for local desktop use and single-server deployments (zero infrastructure overhead, microsecond memory lookups), in-memory rate limiting does not share state across horizontally scaled environments (such as multiple Kubernetes pods or multi-instance server clusters behind a load balancer). 

In a cluster with $N$ pods, an attacker's requests can be distributed across pods, effectively multiplying their allowed request volume by $N$ before being throttled on any single node.

#### Required work

1. Implement a `RedisRateLimiter` adapter following the same interface as `InMemoryRateLimiter`:
   - `is_allowed(key, max_requests, window_seconds) -> tuple[bool, int]`
2. Use Redis atomic sorted sets (`ZADD`, `ZREMRANGEBYSCORE`, `ZCARD`) or a Redis Lua script to evaluate sliding windows atomically across all pods.
3. Configure Redis connection parameters via environment variables (`REDIS_URL`, `REDIS_PASSWORD`).
4. Provide a factory or dependency fallback:
   - If `REDIS_URL` is set, use `RedisRateLimiter`.
   - If `REDIS_URL` is unset, gracefully fall back to `InMemoryRateLimiter` (maintaining standalone and offline compatibility).
5. Ensure rate limit response headers (`Retry-After`) and HTTP 429 status behavior remain identical across both backends.

#### Completion criteria

- Multiple API worker processes or Kubernetes pods share rate limit state via Redis.
- Standalone local deployments continue to run with `InMemoryRateLimiter` without requiring Redis.
- Unit and integration tests verify the distributed sliding window behavior against a mock/test Redis instance.

### BL-009: Optimize and clean up Excel and ICS calendar utility routines

**Area:** Utilities / Export  
**Priority:** Low

#### Problem

1. `src/utils/excel.py` iterates over the same doctor/date matrix twice and prints raw doctor names to stdout rather than using structured logging.
2. `excel_to_ics.py` writes the calendar file to disk, reads the entire file back into memory, and writes it a second time solely to strip blank lines.

#### Required work

1. In `src/utils/excel.py`:
   - Consolidate redundant matrix iteration into a single pass.
   - Replace bare `print()` statements with structured log statements or remove them.
2. In `excel_to_ics.py`:
   - Serialize the `.ics` content cleanly in memory without intermediate disk read/write churn.
3. Add automated unit tests covering both the Excel export and ICS file generation.

#### Completion criteria

- Excel export executes in a single matrix pass without stdout console noise.
- ICS generation writes once to the target path.
- Unit tests verify valid `.xlsx` and `.ics` outputs.

### BL-010: Consolidate or retire legacy schedule_maker.py implementation

**Area:** Solver Engine / Architecture  
**Priority:** Medium

#### Problem

The codebase maintains two divergent scheduling solver implementations:
1. `src/scheduler.py` (`ShiftScheduler`): The modern, active implementation integrated with domain models, configurable soft-constraint weights, and API workflows.
2. `src/schedule_maker.py`: An older standalone script with hardcoded constraints, a different API shape, and separate argument parsers.

Maintaining both files introduces behavioral drift and confusion regarding which constraint formulation is canonical.

#### Required work

1. Audit CLI and desktop entry points to verify if anything still imports `src/schedule_maker.py`.
2. Either:
   - Retire `src/schedule_maker.py` entirely, migrating any unique CLI features to a modern CLI command importing `src.scheduler.ShiftScheduler`.
   - Turn `src/schedule_maker.py` into a thin backward-compatible adapter wrapping `ShiftScheduler`.
3. Remove redundant tests or update them to target `ShiftScheduler`.

#### Completion criteria

- Only one CP-SAT model building implementation exists in the codebase.
- Any CLI workflows run against `ShiftScheduler`.

### BL-011: Configurable Department Schedule Constraint Weights & Solver Settings

**Area:** Solver Engine / Department Preferences / UI  
**Priority:** Medium

#### Problem

Currently, the solver uses default penalty weights and limits from `ScheduleConfig` (`max_duties_per_month: 8`, `solver_time_limit: 120s`, `w_every_other_penalty: 4`, `w_balance_full_wkends_off: 20`, etc.). Different clinical departments have different operational priorities (e.g., intensive care may strictly penalize short gaps between shifts, while pediatrics may prioritize balanced weekends off). Department administrators need the ability to customize these weights.

#### Required work

1. **Spreadsheet Support**: Support an optional 3rd sheet (`Settings` or `Config`) in the `.xlsx` workbook parsing key-value pairs into `ScheduleConfig`.
2. **Database Persistence**: Add a `schedule_config: Optional[dict]` JSON column on `DepartmentModel` to store persistent department scheduling preferences.
3. **API Overrides**: Support an optional `config_override` payload on `POST /api/v1/schedule/generate-from-excel`.
4. **Frontend UI**: Provide an "Advanced Settings" modal in Flutter with weight sliders allowing administrators to tune penalty weights before triggering generation.

#### Completion criteria

- Admins can customize constraint weights via Excel, API, or UI.
- The solver optimizes according to the custom weights.
- When omitted, all flows gracefully fall back to the default `ScheduleConfig()`.

### BL-012: Comprehensive Codebase Docstring Audit & Standardization (PEP 257 / Google Style)

**Area:** Code Quality & Documentation  
**Priority:** Low

#### Problem

Docstring coverage and formatting vary across different development phases in the codebase. Older modules have brief single-line summaries without parameter specifications, while newer modules follow comprehensive Google-style / PEP 257 format with explicit `Args`, `Returns`, and `Raises` sections. Standardizing all docstrings will improve developer onboarding, IDE auto-complete/hover documentation, and future automated API reference generation.

#### Required work

1. Audit all backend packages: `src/auth`, `src/department`, `src/doctor`, `src/schedule`, `src/shift`, `src/team`, `src/user`, `src/utils`, and `src/scheduler.py`.
2. Standardize all public functions, classes, and methods to Google-style docstrings.
3. Ensure all raised exceptions (e.g. `ValueError`, `HTTPException`) are explicitly documented in a `Raises:` section.
4. Optionally enable docstring linting rules (`ruff` rule group `D` or `pydocstyle`) in development tooling.

#### Completion criteria

- 100% of public functions, methods, and classes possess standardized Google-style docstrings.
- Docstring linter reports 0 errors across the codebase.

### BL-013: Codebase Complexity & Nesting Audit (DRY & Max 3 Levels of Indentation)

**Area:** Code Architecture & Maintainability  
**Priority:** Medium

#### Problem

Deeply nested code (such as nested loops with inner `if/else` ladders and `try/except` blocks) increases cognitive complexity, makes unit testing difficult, and increases the likelihood of scoping bugs. Additionally, repeated logic across controllers, services, and utilities violates the DRY (*Don't Repeat Yourself*) principle.

#### Required work

1. **Enforce Maximum 3 Levels of Indentation**:
   - Audit functions across `src/` (particularly loops in `scheduler.py`, `excel_parser.py`, auth services, and controllers).
   - Refactor functions exceeding 3 levels of nesting using guard clauses (early returns / `continue`), generator expressions, and focused helper function extractions.
2. **DRY Review**:
   - Identify repeated patterns across repositories, controllers, and services (e.g., entity existence checks, tenant isolation filters, and JSON serialization patterns).
   - Extract duplicated logic into reusable utility functions.
3. **Automated Enforcement**:
   - Configure cyclomatic complexity and nesting depth checks in our linter (e.g. `ruff` rule `C901` for cyclomatic complexity and `PLR0912` for branch depth).

#### Completion criteria

- No function in the codebase exceeds 3 levels of nesting/indentation.
- Duplicated patterns across features are unified into shared helpers.
- Full test suite passes with zero regressions.

### BL-014: Date-Range Filtering for Department Shift Assignments Query

**Area:** Shift Feature / Performance & Query Optimization  
**Priority:** Medium

#### Problem

`get_active_shift_assignments_for_department` currently queries all active assignments across all time. As historical data accumulates, fetching an entire department's history to render a single month's calendar will degrade query performance and saturate network bandwidth.

#### Required work

1. Add optional `start_date` and `end_date` parameters to `get_active_shift_assignments_for_department` in `src/shift/repository.py`.
2. Add corresponding SQL filtering (`ShiftAssignmentModel.date >= start_date`, `ShiftAssignmentModel.date <= end_date`).
3. Expose optional `start_date` and `end_date` query parameters on `GET /api/v1/shifts/assignments` in `src/shift/routes.py` and `src/shift/controllers.py`.
4. Add unit and integration tests verifying query filtering within date bounds.

#### Completion criteria

- API efficiently filters assignments within the requested date range (e.g. `2026-11-01` to `2026-11-30`).
- Query parameters remain optional and backward-compatible.

### BL-015: Idempotent Schedule Publishing with Soft-Delete Cleanup

**Area:** Schedule Feature / Publishing & Data Integrity  
**Priority:** Medium

#### Problem

When publishing a generated schedule into official `ShiftAssignment` records, regenerating or re-publishing a month would conflict with existing records or violate the unique constraint `uq_shift_assignment_doctor_date` (`doctor_id`, `date`).

#### Required work

1. In the schedule publishing service, atomically soft-delete (`is_deleted = True`) any existing active `ShiftAssignment` records for that department within the target month.
2. Bulk-insert the newly published assignments within the same database transaction.
3. Update `ScheduleDraft.status` to `"published"`.
4. Add automated test coverage verifying idempotent re-publishing.

#### Completion criteria

- An admin can safely re-publish a month without duplicate entries, constraint violations, or orphaned assignments.
- Previous assignments for that month are cleanly soft-deleted within the same transaction.

### BL-016: PostgreSQL Service & Driver Integration for Docker Compose

**Area:** Database / Infrastructure & Docker Deployment  
**Priority:** Medium

#### Problem

The initial Docker Compose setup uses SQLite with a mounted volume for simplicity and zero external dependencies. While suitable for desktop, local development, and single-instance deployments, enterprise and cloud deployments require concurrent multi-user write handling and robust clustering, which PostgreSQL provides.

Additionally, `backend/pyproject.toml` does not currently include PostgreSQL drivers (`psycopg2-binary` or `asyncpg`). Attempting to point `DATABASE_URL` to a PostgreSQL instance will result in runtime errors (`ModuleNotFoundError`).

#### Required work

1. Add a PostgreSQL database driver to backend dependencies via `uv add psycopg2-binary`.
2. Add a `db` service running `postgres:16-alpine` in `docker-compose.yml` with health checks (`pg_isready`).
3. Configure environment variables (`POSTGRES_USER`, `POSTGRES_PASSWORD`, `POSTGRES_DB`, `DATABASE_URL=postgresql://user:password@db:5432/schedule_maker`).
4. Ensure `backend` service in `docker-compose.yml` depends on `db` with `condition: service_healthy`.
5. Maintain SQLite compatibility as default fallback when `DATABASE_URL` is omitted or points to SQLite.
6. Add automated integration test verifying connection and table creation against a PostgreSQL instance.

#### Completion criteria

- Backend seamlessly connects to PostgreSQL in Docker without driver errors.
- Docker Compose can start backend and postgres container with guaranteed ordering and health checks.
- Standalone local runs still default gracefully to SQLite.

## Resolved / Closed


### BL-001: Enable SQLite foreign-key enforcement

**Area:** Database integrity  
**Resolution:** Enabled `PRAGMA foreign_keys=ON` via SQLAlchemy `@event.listens_for(Engine, "connect")` in `src/db/connection.py` for all `sqlite3.Connection` instances. Added regression tests in `src/tests/test_models.py` verifying PRAGMA activation, rejection of nonexistent foreign keys with `IntegrityError`, and acceptance of nullable foreign keys.
