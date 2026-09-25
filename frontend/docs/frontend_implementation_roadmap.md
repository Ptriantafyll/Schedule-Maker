# Flutter frontend implementation roadmap

This is an implementation **plan**, not a description of finished screens. Build the app feature by feature with MVVM and Riverpod, starting with authentication. The first usable release should support Flutter Web, Windows desktop, and Android/iOS. Schedule examples can use **seeded test fixtures** while the backend catches up; fixtures must never silently replace real production data.

Feature plans: [Authentication](frontend_auth_implementation_plan.md) | [Schedule](frontend_schedule_implementation_plan.md) | [Requests](frontend_requests_implementation_plan.md) | [Admin](frontend_admin_implementation_plan.md). The existing [auth progress tracker](auth_progress.md) records implementation status separately from this roadmap.

## 1. What exists today

- `frontend\lib\main.dart` and `frontend\test\widget_test.dart` still contain the starter counter. There are no feature screens or ViewModels. `lib\theme\app_theme.dart` supplies Material 3 light/dark themes; Riverpod, Dio, secure storage, and intl are already declared in `pubspec.yaml`.
- `lib\core\config\app_config.dart`, `core\network\api_client.dart`, `core\network\api_exception.dart`, and `core\storage\token_storage.dart` are starting points, **not** a complete auth/session implementation.
- Backend auth endpoints exist: login, invite-only signup, profile, refresh, logout, and admin invitations. Signup returns a user profile (201), **not** tokens; success takes the invitee to login. See `backend\src\auth\routes.py` and `backend\src\auth\schemas.py`.
- `GET /api/v1/shifts/assignments` currently returns all active department assignments, without a month filter or draft/published status. There is no registered schedule publication endpoint; planned routes in `backend\docs\api_frontend_implementation_plan.md` are not implemented routes. See `backend\src\shift\routes.py`, `backend\src\shift\models.py`, and `backend\src\main.py`.
- `POST /api/v1/doctors/{doctor_id}/unavailability` writes an effective unavailability immediately. The backend has no pending-request approval or shift-swap API. See `backend\src\doctor\routes.py`. Do **not** use that POST as a substitute for an approval request.

### Backend readiness gates

| Capability | Available now | Needed before a real user-facing flow |
| --- | --- | --- |
| Login, invite signup, `/auth/me`, refresh, logout | Yes, with different native/browser session transports | Finish Flutter integration and correct existing auth docs. |
| Role-scoped doctor roster, shifts, positions, assignments | Yes, but assignments are unfiltered and not published-only | Add server-enforced draft/published lifecycle and bounded month reads before showing any assignments to doctors/viewers. |
| Super-admin department provisioning; department-admin doctor/team/position/shift creation and invitations | Selected create/read/invite actions exist | Limit initial forms to supported actions; do not imply edit/delete works. |
| Schedule publishing | No | Add department-admin publication and published-only member reads on the backend. A client-side filter does not prevent drafts leaking. |
| Pending unavailability + admin decision | No; direct unavailability write exists | Add request submission, own/status listing, scoped admin review/decision, and an effective unavailability record only on approval. |
| Shift swap + admin decision | No | Add a server-side request/decision workflow with assignment checks and atomic application on approval. |

**Meaning of "blocked":** The corresponding UI/VM can be tested with explicit fixtures, but do not connect it to an unrelated endpoint, claim the flow works end-to-end, or expose it as a live action until the backend enforces its contract.

## 2. Implementation order (first, second, third...)

1. **Authentication and necessary core networking** - Build models/DTOs, session transport, repository, auth ViewModel, login, invitation signup, restore/refresh/logout, and an auth gate. Why first: every other feature needs a real identity, department scope, and reliable session. Completion: native and web users can sign in, accept an invitation then sign in, restore or expire a session, and sign out without leaking tokens. Details: [auth plan](frontend_auth_implementation_plan.md).
2. **Thin application shell and shared UI foundation** - Replace the counter with guarded routing and role-aware destinations; use responsive Material 3 navigation (`NavigationBar` on narrow layouts, `NavigationRail` or equivalent when wide). Set up explicit loading/error/empty/retry behavior and invitation-link handling on each platform. Why second: later screens reuse one navigation and feedback convention instead of building four incompatible ones.
3. **Read-only schedule (fixture-backed first)** - Add schedule models, repository interface, month ViewModel, calendar/list layout, shift cards, and "my shifts" for doctors. Test with **seeded published fixtures** and an empty month. Why third: it is the main user-facing destination and supplies assignment selection for later swap requests. **Live integration gate:** server-enforced publication and month-scoped reads; never expose the existing all-assignment endpoint as a published schedule. Details: [schedule plan](frontend_schedule_implementation_plan.md).
4. **Admin foundations, then publishing** - Build separate super-admin provisioning/invites and department-admin roster/setup/invites using existing routes. Add draft/publish actions **only when** backend publication exists. Why fourth: staff/data setup and the admin shell will also host request review. Details: [admin plan](frontend_admin_implementation_plan.md).
5. **Unavailability requests with admin review** - Add a doctor's date submission and request history, plus a department-admin pending list and decision UI. Why fifth: it establishes the shared approval/status pattern before swaps. **Live integration gate:** new pending/approval APIs; the existing direct-write endpoint cannot represent a pending request. Details: [requests plan](frontend_requests_implementation_plan.md).
6. **Shift-swap requests with admin review** - Reuse request status/list/feedback, but create assignment-specific selection and validation. Why sixth: swaps require published assignments and the approval pattern from step 5. **Live integration gate:** a backend swap workflow that safely checks and applies assignments; no client-side reassignment.
7. **End-to-end hardening on every target platform** - Verify role isolation, publication, request state changes, responsive layout, accessibility, and browser/native session behavior. A feature with an open backend gate is not "done"; keep its fixture-only status visible in planning.

Each step includes small unit, Riverpod provider, and widget tests as it is built; cross-platform checks are **not** postponed until step 7. Backend work is listed only as a prerequisite, not as frontend implementation.

## 3. MVVM and folder ownership

Within `lib\features\auth`, `schedule`, `requests`, and `admin`, keep:

- **Model:** Pure Dart domain entities and rules in `domain\`; API request/response DTOs, JSON parsing, remote data sources, and repository implementations in `data\`. Do not put Dio or Flutter in domain entities.
- **ViewModel:** Small Riverpod `Notifier`/`AsyncNotifier` providers in `presentation\` for one cohesive workflow (e.g. session, selected schedule month, request submission, request review). Expose loading, success, empty, and errors deliberately.
- **View:** Screens and widgets in `presentation\` that observe state and send user actions to ViewModels. No HTTP, token storage, or ad-hoc permission parsing in `build()`.

Flow: **View -> ViewModel -> repository -> API**; state flows back through Riverpod. Use providers for dependency injection and override them in tests. Keep `lib\core\` for configuration/network/session infrastructure, `lib\theme\` for styles, and `lib\common_widgets\` for genuinely cross-feature UI. A widget is "shared" only when at least two real consumers need the same behavior; start with built-in Material widgets rather than speculative abstractions.

### Widget reuse map

| Owner | Candidate widgets/patterns | Reuse boundary |
| --- | --- | --- |
| Shared shell | `AdaptiveAppShell`, role-aware navigation, route guard | All authenticated feature destinations; permissions still enforced on the backend. |
| Shared feedback/form | Loading action button, consistent form-field decoration, async loading/error/empty/retry, confirmation dialog | Extract after login/signup, schedule/requests, or multiple admin actions demonstrate the same need. Prefer standard `TextFormField`, `Future`/Riverpod states, and Material dialogs until then. |
| Shared date/status | Month navigator/date formatting; request-status chip | Reuse a month header for schedule and requests if both need it; reuse the chip in doctor and admin request lists. Do not force their day cells or forms to be identical. |
| Auth | `AuthGate`, login/invite forms, password visibility field | Credentials and invitation tokens remain auth-owned. |
| Schedule | Month grid/list, schedule day cell, shift-assignment card | A read-only assignment is not an availability date picker. |
| Requests | Unavailability date picker/form, swap assignment selector/form, request list | Share status/feedback, not a form full of unrelated optional fields. |
| Admin | Provision/invite/roster forms, scoped review list, publish action | Separate super-admin and department-admin flows. |

### Roles and navigation

| Role | Destination(s) in the plan | Important boundary |
| --- | --- | --- |
| Super admin | Admin: department provisioning and admin invitations | No department scope; do not send to a department schedule or roster API. |
| Department admin | Published schedule, department admin, pending request review | May prepare/publish only their own department schedule once supported. |
| Doctor | Published department schedule, own assignments, own requests | Submit their own requests; no admin controls. |
| Viewer | Published department schedule | Read-only; no request or admin actions. |

Derive the role and identifiers from `/api/v1/auth/me`; hiding navigation is UX, not access control. Handle 401 (session), 403 (permission), 404 (missing or out-of-scope), form validation/conflict, and rate limiting distinctly.

### Web/Windows/mobile from the start

- **Native (Windows, Android/iOS):** Persist rotating refresh credentials in supported secure storage, attach short-lived access tokens to protected requests, and use body-token refresh/logout. Check the storage plugin on each target platform.
- **Web:** Use browser-managed `HttpOnly` refresh cookie (never store a JSON refresh token in browser storage), in-memory access token, credentials-enabled requests, and `X-CSRF-Token` for cookie refresh. Configure frontend origin/CORS, HTTPS and secure cookies for deployment; validate browser restart/redirect behavior.
- **All:** Validate `API_BASE_URL`, do not log tokens or embed them in normal URLs, keep layouts usable at narrow/wide widths, and include platform-appropriate invitation-link routing. Decide routing/deep-link implementation during shell setup; avoid a dependency before its need is clear.

## 4. Acceptance checkpoints and deferrals

For each feature, test repository mapping, ViewModel transitions, and view states (loading, data, empty, error); then confirm the actual backend response and role boundary before marking integration complete. Schedule tests must show drafts **never** reach ordinary users; request tests must keep pending submissions distinct from effective changes and prove an admin decision is scoped to the correct department. Run `flutter test` and `flutter analyze` when implementing code, plus targeted browser/native integration checks.

Out of scope for this first plan: open public signup, offline sync, solver generation/export UI, unsupported admin edit/delete actions, notifications, or cancellation/consent/deadline rules not yet agreed for requests. Revisit these explicitly rather than inventing UI behavior.
