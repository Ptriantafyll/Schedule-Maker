# Implementation plan: published schedule

This is **step 3** of the [frontend roadmap](frontend_implementation_roadmap.md), after [authentication](frontend_auth_implementation_plan.md) and the shared app shell. Start with a **read-only** schedule UI for Web, Windows, and Android/iOS using clearly identified seeded **published** fixtures. No real doctor/viewer schedule may go live until the backend protects drafts.

## 1. Purpose, users, and backend boundary

- **Doctor:** view the published department schedule and filter to their own assignments using `doctor_id` from `/auth/me`.
- **Viewer:** read the published department schedule; no editing or requests.
- **Department admin:** view published schedules here; draft preparation/publication belong to the [admin plan](frontend_admin_implementation_plan.md).
- **Super admin:** no department scope and no departmental schedule destination.

The current `GET /api/v1/shifts/assignments` returns **all active assignments** in a department, without a month filter or publication status. It returns assignment IDs, `doctor_id`, `shift_id`, and a date; `GET /api/v1/doctors/roster`, `/shifts/`, and `/positions/` can supply display labels. There is no registered `/api/v1/schedule` reader, draft/published marker, generate, or export route. Evidence: `backend\src\shift\routes.py`, `backend\src\shift\schemas.py`, `backend\src\shift\repository.py`, `backend\src\main.py`. Older backend planning docs describe **proposed**, not mounted, endpoints.

**Blocking backend prerequisite:** Publish a department/month schedule explicitly, keep drafts inaccessible to doctors/viewers **in the server query/authorization layer**, and provide a bounded published-schedule read for a selected month. Admins need separately authorized draft access and a publish action. Decide how revisions of an already published month become visible before wiring live UI. A client-side filter or a nonempty assignment list does not establish publication.

## 2. MVVM ownership and suggested widgets

> [!IMPORTANT]
> **Presentation Layer Rule**: For the presentation layer (UI screens, dialogs, forms, layout widgets), **always ask the user for the design first and ask clarifying questions** before writing any code or proposing UI designs.

| Layer | What to plan | Why |
| --- | --- | --- |
| Domain model | Immutable `ScheduleMonth`, `PublishedSchedule` (month, department, publication status, assignments), `ShiftAssignment` (IDs and date), display-ready doctor/shift/position reference models | Dates and IDs should be predictable across layouts; the published state must come from a trusted server response in production. Never infer it from local assignments. |
| Data | Schedule DTOs/mapper; fixture repository for tests/development; later an API-backed `ScheduleRepository` that reads **only** published months for members | Keep raw API responses out of widgets. If roster/shift lookups are also needed in admin, share one resource repository/mapper rather than duplicate JSON parsing. |
| ViewModel | `ScheduleViewModel`/Riverpod `AsyncNotifier` for selected month, load/retry, and optional own-shifts filter | A single source of truth for current month and loading/empty/error/data state. The doctor filter is for UX, not the backend's publication or department access check. |
| View | `ScheduleScreen`, `MonthNavigator`, `ScheduleMonthGrid`, `ScheduleDayCell`, `ShiftAssignmentCard`, compact `ScheduleList`, optional `MyShiftsFilter` | A calendar fits wide screens; an agenda/list fits narrow screens. Keep assignment rendering schedule-owned, not the same widget as a request date picker. |

Suggested feature structure: `lib\features\schedule\domain\`, `data\`, and `presentation\` with mirrored `test\features\schedule\`. Import role/identity from the auth provider and reuse existing theme, shared async feedback, and any month-navigation pattern shared later by requests. Prefer Material widgets and `intl` for labels; avoid inventing a generic calendar framework before both use cases are known.

## 3. Implement in this order

1. **Define the read contract and fixtures.** Specify what a published month and assignment mean in frontend domain terms, and what extra backend capability is missing. Create seeded **test-only published** examples for an empty month, multiple shifts, and a doctor's own assignments. Why: the screen can be designed without claiming the current raw assignment endpoint is safe.
2. **Model/DTO/repository seam.** Parse dates as date-only values (`YYYY-MM-DD`), keep IDs stable, and assemble display labels from roster/shift/position data where needed. Return an explicit empty published month versus a not-yet-published/blocked month; do not silently turn errors into an empty schedule. Add repository/mapper tests with role and missing-reference cases. Why: date/time conversions or missing labels should not corrupt the calendar.
3. **Month ViewModel.** Load a selected month, move previous/next, retry failures, and optionally show only the signed-in doctor's assignments. Watch auth/session state so a logout or department change invalidates old data. Model loading, published-empty, published-with-data, not-published, and error deliberately; avoid fetching every department assignment on every month switch. Why: screens need one predictable state machine.
4. **Responsive views.** Render a read-only month grid on wide Web/Windows/tablets and a scannable date-grouped list on narrow phones. Add an accessible month header, assignment card, clear empty/unpublished/error messages, and role-aware labels. Use a local month header shared with requests only after it is truly reused. Why: the same data must remain usable without duplicating business rules per layout.
5. **Test the UI with fixtures, then stop at the API gate.** Cover month boundaries, multiple assignments per date, no published schedule, an empty published schedule, retry, doctor "my shifts", viewer read-only mode, super-admin exclusion, and narrow/wide layouts. Wire production repository **only after** the backend enforces publication and bounded department/month reads. Why: a mock UI milestone is not a safe production release.

## 4. Acceptance criteria and explicit deferrals

- With published fixture data, doctor/viewer/admin members see the correct month; doctors can filter their own shifts; no view offers editing. The super admin is routed to admin provisioning instead.
- No ordinary user sees draft assignments, including through direct API access. This requires backend tests and a real published-only API, **not** just a hidden draft label in Flutter.
- Empty published, unpublished, loading, network error, forbidden, and missing/foreign records have distinct UI outcomes. A response with no assignments is not automatically "published and empty."
- Tests exercise the repository mapper, Riverpod month changes, widget states, and cross-platform responsive layouts. Live integration verifies role/tenant boundaries once the backend gate is delivered.
- Solver generation, Excel export, draft editing, schedule changes, and real-time notifications are **not** part of this read-only feature. Publishing belongs to admin; swap requests depend on a published assignment and are planned [separately](frontend_requests_implementation_plan.md).
