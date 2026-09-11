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
