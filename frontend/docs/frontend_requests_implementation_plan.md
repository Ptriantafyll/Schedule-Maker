# Implementation plan: approved requests

This covers **steps 5-6** of the [frontend roadmap](frontend_implementation_roadmap.md): first unavailability requests, then shift-swap requests. Both are **pending until a department admin approves or rejects them**. The [admin plan](frontend_admin_implementation_plan.md) owns the review screen; the [schedule plan](frontend_schedule_implementation_plan.md) supplies published assignments for swap selection. Target Web, Windows, and Android/iOS from the start.

## 1. Business flow and current API gap

1. A doctor submits a request; the doctor can see its pending status but the submission has **no immediate effect** on the schedule or effective unavailability.
2. A department admin reviews requests **in their own department** and approves or rejects one. The server validates the decision and applies the effective change only on approval. The requester then sees the result.
3. A viewer cannot submit or review requests. A super admin has no department scope; do not expose a department request queue to that role.

The backend currently has `GET`/`POST /api/v1/doctors/{doctor_id}/unavailability` for an authenticated doctor (own record) or department admin. The `POST` creates an **effective** unavailability immediately and returns a `DoctorUnavailabilityRead`; it is **not** a pending request. There is no request status/review API or shift-swap API in registered routes. Sources: `backend\src\doctor\routes.py`, `doctor\schemas.py`, `doctor\controllers.py`, and `backend\src\main.py`. Do not connect the doctor's proposed request Submit button to that direct-write endpoint.

**Backend prerequisites, not invented endpoint specifications:** submission, own/list statuses, department-scoped admin pending list, approve/reject transition, and authorization for each request type. Approval of unavailability must create/activate the effective record; rejection must not. A swap must validate the involved assignments/doctors and apply any reassignment atomically. Establish rules for past dates, deadlines, overlap with assigned shifts, duplicate requests, recipient consent/eligibility, and cancellation **before** implementing the backend contract; this plan does not silently choose those policies.

## 2. MVVM structure and widget reuse

> [!IMPORTANT]
> **Presentation Layer Rule**: For the presentation layer (UI screens, dialogs, forms, layout widgets), **always ask the user for the design first and ask clarifying questions** before writing any code or proposing UI designs.

| Layer | Unavailability (build first) | Swap (add second) | Shared boundary |
| --- | --- | --- | --- |
| Domain/data | Pending request with date(s), identity/department and status; DTOs and request repository | Pending request referencing the requester's published assignment and proposed swap details; DTOs/repository methods when backend contract exists | One small request ID/status vocabulary (`pending`, `approved`, `rejected`) and list mapping if the server shares it. Avoid one DTO full of unrelated optional fields. |
| ViewModel | `UnavailabilityRequestViewModel` for date validation/submission; `MyRequestsViewModel` for status/history | `SwapRequestViewModel` for assignment selection/submission, reusing history/state patterns | `RequestReviewViewModel` belongs in admin; share repository/domain types, not a single ViewModel with all doctor/admin actions. Riverpod providers handle loading, error, empty, and success explicitly. |
| View | `UnavailabilityRequestForm` using Material date selection; `RequestList`, `RequestCard` | `SwapAssignmentPicker`, `SwapRequestForm` using the published schedule | Share a status chip and common feedback/submit-button pattern with the admin queue after real reuse. Do **not** reuse schedule's read-only day cell as a request date picker or build a generic form with many unused fields. |

Suggested folders: `lib\features\requests\domain\`, `data\`, `presentation\` and `test\features\requests\`. The [auth plan](frontend_auth_implementation_plan.md) owns identity and `doctor_id`; do not ask the user to choose an arbitrary doctor/department ID. The schedule feature provides a **published** assignment selection model for swaps; the request repository owns submission/status data. `View -> ViewModel -> repository -> API`, with provider overrides for test fixtures.

## 3. Step-by-step work

1. **Agree the request lifecycle with the backend team.** Record status meanings, who can submit/review, when an approved request becomes effective, and how duplicates/conflicts are resolved. Mark the production integration **blocked** until APIs implement it. Why: the current direct-write endpoint has different business behavior.
2. **Build shared request status/history UI with explicit fixtures.** Add immutable request models, test-only pending/approved/rejected fixtures, repository interface, `MyRequestsViewModel`, list/cards, loading/empty/error/retry, and status labels. Fixture mode is for tests/prototypes only; no success-shaped fallback when a real request endpoint is missing. Why: both request types and admin review need the same status language.
3. **Build unavailability submission with separate state.** Use the current user's doctor identity, a date picker, duplicate/past-date validation once policy is agreed, submit/progress/confirmation, and a pending history entry **only if the server actually accepted a pending request**. Add a matching admin pending list and approve/reject workflow when its backend capability lands. Why: a request must not quietly become an effective unavailability at submit time.
4. **Build swap submission using published schedule data.** Select an assignment belonging to the requester; present any proposed partner/shift options only according to the eventual backend eligibility/consent contract. Reuse the request list/status/error patterns, not unavailability's date-only form. Add admin decision UX after the backend can atomically validate and apply swaps. Why: changing assignment rows directly would bypass approval and shift constraints.
5. **Connect each flow to the live API only after its gate passes.** Keep the new backend's doctor/department scope, decision conflict behavior, and response statuses visible in UI errors. Refresh request history and the published schedule after an approved decision; do not optimistically claim approval on a failed or stale request. Why: the server is the authority for actual assignment/unavailability changes.
6. **Test every target platform.** Cover short/long names, dates, focus and keyboard navigation, narrow phone versus wide desktop/Web layouts, signed-out/session-expired transitions, and 403/404/409/422/network errors as the contracts become available.

## 4. Acceptance checkpoints and non-goals

- **Unavailability phase:** doctor submission becomes `pending`, appears in own history and own department's admin queue; rejection leaves effective unavailability unchanged; approval creates/activates it **once**. A direct-write 201 is not counted as this phase working.
- **Swap phase:** only an eligible request tied to a published assignment is accepted; approval/rejection is recorded once, and only approval updates assignments. Conflicting or stale assignments produce an explicit error. Rules about recipient consent and eligibility remain a backend design decision until agreed.
- **Permission boundary:** doctors cannot review others' requests; viewers cannot submit; super admins do not receive department queues; a department admin cannot approve another department's requests. Backend authorization tests, not hidden buttons, prove these boundaries.
- **Frontend tests:** DTO mapping, `AsyncNotifier` transitions, form validation, pending/approved/rejected/empty/error widgets, status updates in doctor and admin views, and responsive layout with fixture repositories. Live integration tests require real request endpoints; leave that checkpoint blocked meanwhile.
- Cancellation, automatic notifications, scheduling cutoffs, and user-to-user negotiation are intentionally unspecified rather than implied by a screen label.
