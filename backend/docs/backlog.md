# Backend Backlog

Deferred issues that should be addressed outside the current implementation step.

## Open

### BL-001: Enable SQLite foreign-key enforcement

**Area:** Database integrity  
**Priority:** High

#### Problem

The SQLModel models declare foreign keys, but SQLite does not enforce them
unless `PRAGMA foreign_keys=ON` is enabled for every database connection.
The current database engine does not enable this setting.

This can allow orphaned records, such as a user whose `department_id` does not
reference an existing department. It can also make development and tests behave
differently from production databases that enforce foreign keys by default.

#### Required work

1. Enable `PRAGMA foreign_keys=ON` for every SQLAlchemy SQLite connection.
2. Keep the configuration conditional so it is not executed for other database
   engines.
3. Add a regression test proving that a nonexistent foreign-key value raises an
   integrity error.
4. Add a test proving that nullable foreign keys still accept `NULL`, including
   the tenantless super-admin case.

#### Completion criteria

- Every application and test SQLite connection reports foreign-key enforcement
  as enabled.
- SQLite rejects records that reference nonexistent parent records.
- Nullable foreign-key columns continue to accept `NULL`.
- The existing test suite remains green.

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

