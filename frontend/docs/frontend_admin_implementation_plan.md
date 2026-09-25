# Implementation plan: role-specific admin

This is **step 4** of the [frontend roadmap](frontend_implementation_roadmap.md) for existing administration and schedule publication, and the admin-side companion to [requests](frontend_requests_implementation_plan.md) in steps 5-6. The first release targets Web, Windows, and Android/iOS. This is a plan, not a claim that an admin screen already exists.

## 1. Two different admin roles

| Role | Supported actions now | Additional planned actions and gate |
| --- | --- | --- |
| `super_admin` | List departments; provision a department with an initial admin invitation; invite a department admin for an existing department; list/revoke invitations | No department schedule, roster, or routine request review. A super admin has **no** `department_id`. |
| `department_admin` | Create/read doctors, teams, positions, shifts; read scoped users; invite doctor/viewer staff, list/revoke scoped invitations; manually create assignments and effective unavailability via existing routes | Privately prepare/publish a schedule, and approve/reject pending unavailability and swap requests **only after** the backend supports those workflows. |

Sources: `backend\src\auth\routes.py`, `auth\schemas.py`, `auth\dependencies.py`, `department\routes.py`, `doctor\routes.py`, `team\routes.py`, `position\routes.py`, `shift\routes.py`, `user\routes.py`. Existing APIs mostly create and read; DTOs for update do **not** mean PATCH/DELETE routes are available. Keep unsupported edit/delete controls out of the live UI.

Invitation contracts matter for forms: super-admin provisioning takes **department name and code**, not an admin email; a super-admin re-invitation for an existing department takes its ID. A department-admin staff invitation takes role (`doctor` or `viewer`) and optional `doctor_id`; it does **not** take the invitee's email. The created invitation returns a **raw token once**; later invitation lists do not return it. The initial plan should show a one-time copy/share action with explicit expiry/handling and never log or persist the raw token.

## 2. MVVM and widget ownership

| Layer | Candidate responsibilities | Reuse |
| --- | --- | --- |
| Domain and data | Department/resource/invitation entities, DTOs, and small repositories for the supported API calls; later publishing and decision repositories | Reuse roster/shift/position mapping with the schedule feature when it actually needs the same data. Reuse request domain status/decision mapping with the requests feature rather than defining a second request model. |
| ViewModels | `ProvisionDepartmentViewModel`, `DepartmentResourcesViewModel` (split by resource if it grows), `InvitationsViewModel`; later `PublishScheduleViewModel` and `RequestReviewViewModel` | Separate workflow state/loading/errors; do not build one giant `AdminController` with unrelated mutable fields. Derive role/department from auth, not admin form input. |
| Views | `AdminScreen`/role-aware sections, `ProvisionDepartmentForm`, `DepartmentResourcesList`, `StaffInviteForm`, `InvitationList`, later `PublishScheduleAction`, `PendingRequestList`, and decision confirmation | Use shared app shell, form fields/action feedback, async list states, and request-status chip. Keep role-specific forms owned by admin; only extract shared patterns after a second real use. |

Suggested folder structure: `lib\features\admin\domain\`, `data\`, `presentation\` and `test\features\admin\`. The auth feature owns identity/role, the schedule feature owns assignment rendering, and the requests feature owns the request state vocabulary. Keep `View -> ViewModel -> repository -> API`; the backend, not navigation, enforces role and tenant boundaries.

## 3. Build in this order

1. **Role-aware admin entry.** Route a `super_admin` to provisioning/invitations and a `department_admin` to their department tools. Doctors/viewers cannot enter the admin destination; show explicit 403 handling if the server rejects a request despite the UI guard. Why: the two admin roles have different API scopes.
2. **Supported super-admin actions.** Start with department list, provision form (name/code), existing-department admin invitation, invitation list/revoke, and the one-time token display/copy. Test loading, empty, conflict, expired/revoked invitation, and failed writes. Why: provisioning is the supported route for creating a department and its first admin; the client must not invent a public admin-signup endpoint.
3. **Supported department-admin actions.** Add scoped doctor/team/position/shift lists and simple create forms, then doctor/viewer invitations and current-user list if useful. Use the exact request DTO fields and server-derived department scope; do not offer unsupported updates/deletes or ask for staff email in an invitation payload the backend does not accept. Why: these APIs can support an initial working admin area with seeded schedule data.
4. **Publishing (backend gate).** Plan an admin-only draft overview and explicit publish confirmation for a department/month. **Do not wire a Publish button** to `POST /shifts/{id}/assignments` or treat a raw assignment as published. The backend must supply scoped draft access, a publish operation, and published-only reads for members. Why: publishing determines when doctors/viewers may see a roster and must be enforced server-side.
5. **Request review (backend gate).** Introduce a department-scoped pending list and approve/reject action alongside [unavailability](frontend_requests_implementation_plan.md); extend the same list/status UI for swaps. The backend must own request statuses, transitions, authorization, and effective changes. Why: admin approval is a workflow, not a direct write to an unavailability or shift assignment endpoint.
6. **Cross-platform and role checks.** Test compact mobile forms, wide desktop/Web lists, keyboard/focus, and success/error feedback with provider overrides. Integrate against actual role-protected endpoints only when each gate exists.

## 4. Backend prerequisites and completion

- **Available now:** super-admin department listing/provisioning and admin invitations, department-admin limited resource creation/reads and doctor/viewer invitations. A department admin can add a shift assignment, but that is not a publishing API.
- **Not available now:** draft/published schedule states and publication; pending unavailability/swap lists and approve/reject endpoints; general edit/delete management; automatic invitation email. Show these as blocked dependencies, **not** enabled actions or fake success messages.
- **Acceptance for the supported slice:** super admins can only provision/manage their authorized invitations; department admins can only manage their own roster/invites; doctors/viewers see no admin navigation. On 403 or foreign-resource 404 the UI reports an error instead of retrying with a different department ID. Invitation tokens appear only when created and are not recoverable from the list.
- **Acceptance after backend gates:** one department's draft/publish or request decision cannot affect another department; a member never reads a draft; an approved request changes the effective schedule/unavailability only once. Validate these against backend authorization tests as well as Flutter provider/widget tests.
