# Schedule Publishing & Month Navigation Progress

This document tracks the step-by-step progress and verification status for the **Schedule Publishing & Dynamic Target Month Navigation** feature across the backend and Flutter frontend, implementing the specifications from [schedule_publishing_and_month_navigation_plan.md](schedule_publishing_and_month_navigation_plan.md).

---

## Progress Overview

| Step | Phase | Scope / Endpoint | Status | Verified Tests |
| :--- | :--- | :--- | :--- | :--- |
| **Step 1** | Backend | Schedule History & Listing (`GET /api/v1/schedules`) | `[x] Completed` | 6 unit & integration tests |
| **Step 2** | Backend | Target Month Resolution (`GET /api/v1/schedules/target-month`) | `[x] Completed` | 10 unit & integration tests |
| **Step 3** | Backend | Schedule Publishing (`POST /api/v1/schedules/{id}/publish`) | `[x] Completed` | 15 unit & integration tests |
| **Step 4** | Frontend | Remote Data Sources & Repositories | `[x] Completed` | 27 unit tests (Models, DTOs, DataSource, Repo) |
| **Step 5** | Frontend | State Management & Controller Updates | `[x] Completed` | 23 unit tests (State & Controller) |
| **Step 6** | Frontend | UI Navigation & Publish Action | `[x] Completed` | 35 unit & widget tests across widgets & screen |

---

## Step 1: Backend — Schedule History & Listing Query `[x] Completed`

### 1. Business & Architectural Goals
- Enabled administrators and department members to retrieve historical drafts and published schedules.
- Replaced monolithic draft payload with lightweight summaries to keep dropdown navigation fast and bandwidth-efficient.

### 2. Implementation Deliverables
- **Schema** ([`src/schedule/schemas.py`](file:///C:/Users/ptria/source/repos/Schedule-Maker/backend/src/schedule/schemas.py)):
  - Implemented `ScheduleSummaryRead` containing metadata (`id`, `department_id`, `target_month`, `status`, `total_duties`, `solver_status`, `source_filename`, `created_at`, `updated_at`) while omitting heavy `assignments` and `unavailabilities` JSON structures.
- **Repository** ([`src/schedule/repository.py`](file:///C:/Users/ptria/source/repos/Schedule-Maker/backend/src/schedule/repository.py)):
  - Implemented `list_schedules_for_department(session, department_id)`:
    - Filters by `department_id` and `not_(ScheduleDraft.is_deleted)`.
    - Orders by `ScheduleDraft.target_month DESC, ScheduleDraft.created_at DESC`.
- **Controller** ([`src/schedule/controllers.py`](file:///C:/Users/ptria/source/repos/Schedule-Maker/backend/src/schedule/controllers.py)):
  - Implemented `list_schedules_controller` with multi-tenant verification via `_resolve_and_verify_department`.
- **Route** ([`src/schedule/routes.py`](file:///C:/Users/ptria/source/repos/Schedule-Maker/backend/src/schedule/routes.py)):
  - Mounted `GET /api/v1/schedules` protected by `require_department_member`.

### 3. Tests Verified (Green Phase)
- `test_schedule_summary_read_schema_serialization` in [`test_models.py`](file:///C:/Users/ptria/source/repos/Schedule-Maker/backend/src/schedule/tests/test_models.py)
- `test_list_schedules_for_department_repository_sorting_and_filtering` in [`test_controllers.py`](file:///C:/Users/ptria/source/repos/Schedule-Maker/backend/src/schedule/tests/test_controllers.py)
- `test_list_schedules_controller_rejects_cross_department_access` in [`test_controllers.py`](file:///C:/Users/ptria/source/repos/Schedule-Maker/backend/src/schedule/tests/test_controllers.py)
- `test_list_schedules_endpoint_success` in [`test_routes.py`](file:///C:/Users/ptria/source/repos/Schedule-Maker/backend/src/schedule/tests/test_routes.py)
- `test_list_schedules_endpoint_excludes_soft_deleted` in [`test_routes.py`](file:///C:/Users/ptria/source/repos/Schedule-Maker/backend/src/schedule/tests/test_routes.py)
- `test_list_schedules_endpoint_rejects_cross_department_access` in [`test_routes.py`](file:///C:/Users/ptria/source/repos/Schedule-Maker/backend/src/schedule/tests/test_routes.py)

---

## Step 2: Backend — Dynamic Target Month Resolution `[x] Completed`

### 1. Business & Architectural Goals
- Replaced hardcoded `'2026-11'` strings with dynamic target month calculation based on actual operational history.
- Established **`ShiftAssignment` as the operational source of truth**:
  - The next month to schedule is the month immediately following the department's latest active shift assignment.
  - If no assignments exist yet, falls back to next calendar month from today (`datetime.date.today() + 1 month`).

### 2. Implementation Deliverables
- **Repository** ([`src/shift/repository.py`](file:///C:/Users/ptria/source/repos/Schedule-Maker/backend/src/shift/repository.py)):
  - Implemented `get_latest_shift_assignment_date_for_department(session, department_id)`:
    - Joins `ShiftAssignment` $\to$ `Shift` $\to$ `Position`.
    - Filters by `Position.department_id == department_id`.
    - Excludes soft-deleted entities across all three tables.
    - Queries `func.max(ShiftAssignmentModel.date)`.
- **Service** ([`src/schedule/service.py`](file:///C:/Users/ptria/source/repos/Schedule-Maker/backend/src/schedule/service.py)):
  - Implemented `_compute_next_month_str(year, month) -> str` handling December $\to$ January year rollovers with zero-padding.
  - Implemented `resolve_next_target_month(session, department_id)` returning `{"next_target_month": str, "last_published_month": str | None}`.
- **Schema** ([`src/schedule/schemas.py`](file:///C:/Users/ptria/source/repos/Schedule-Maker/backend/src/schedule/schemas.py)):
  - Added `TargetMonthResponse` with `next_target_month: str` and `last_published_month: Optional[str] = None`.
- **Controller & Route** ([`src/schedule/controllers.py`](file:///C:/Users/ptria/source/repos/Schedule-Maker/backend/src/schedule/controllers.py), [`src/schedule/routes.py`](file:///C:/Users/ptria/source/repos/Schedule-Maker/backend/src/schedule/routes.py)):
  - Implemented `get_target_month_controller`.
  - Mounted `GET /api/v1/schedules/target-month` protected by `require_department_member`.

### 3. Tests Verified (Green Phase)
- Repository tests in [`test_shift.py`](file:///C:/Users/ptria/source/repos/Schedule-Maker/backend/src/shift/tests/test_shift.py):
  - `test_get_latest_shift_assignment_date_for_department_returns_max_date`
  - `test_get_latest_shift_assignment_date_excludes_deleted_and_foreign_records`
  - `test_get_latest_shift_assignment_date_empty_department_returns_none`
- Service tests in [`test_service.py`](file:///C:/Users/ptria/source/repos/Schedule-Maker/backend/src/schedule/tests/test_service.py):
  - `test_resolve_next_target_month_advances_following_month`
  - `test_resolve_next_target_month_handles_december_year_rollover`
  - `test_resolve_next_target_month_empty_department_falls_back_to_next_calendar_month`
  - `test_resolve_next_target_month_ignores_soft_deleted_assignments`
- Controller tests in [`test_controllers.py`](file:///C:/Users/ptria/source/repos/Schedule-Maker/backend/src/schedule/tests/test_controllers.py):
  - `test_get_target_month_controller_rejects_cross_department_access`
- Schema tests in [`test_models.py`](file:///C:/Users/ptria/source/repos/Schedule-Maker/backend/src/schedule/tests/test_models.py):
  - `test_target_month_response_schema_serialization`
- Route integration tests in [`test_routes.py`](file:///C:/Users/ptria/source/repos/Schedule-Maker/backend/src/schedule/tests/test_routes.py):
  - `test_get_target_month_endpoint_empty_department`
  - `test_get_target_month_endpoint_with_published_assignments`
  - `test_get_target_month_endpoint_rejects_cross_department_access`

---

## Step 3: Backend — Schedule Publishing `[x] Completed`

### 1. Business & Architectural Goals
- Materialize draft schedules into permanent [`ShiftAssignment`](file:///C:/Users/ptria/source/repos/Schedule-Maker/backend/src/shift/models.py#L35) records.
- Ensure strict **ACID compliance**: all intermediate operations (soft-deletes, auto-provisioning staging, bulk assignment writes) use uncommitted flushes (`commit=False` / `session.flush()`), finalized by a single atomic commit at the end.
- Enforce idempotency ([BL-015](backend/docs/backlog.md)): soft-delete existing assignments for that month/department before inserting new ones.
- Auto-provision missing doctors, positions, and shifts for the department ([BL-017](backend/docs/backlog.md)).
- Transition `ScheduleDraft.status = "published"` and update `updated_at`.
- Verify that publishing immediately causes `GET /api/v1/schedules/target-month` to advance to the subsequent month.

### 2. Implementation Deliverables
- **Repositories**:
  - [`src/shift/repository.py`](file:///C:/Users/ptria/source/repos/Schedule-Maker/backend/src/shift/repository.py):
    - `soft_delete_shift_assignments_for_department_month(..., commit=True)`
    - `bulk_create_shift_assignments(..., commit=True)`
    - `stage_shift(session, shift_data)` (in-session flush without commit)
  - [`src/doctor/repository.py`](file:///C:/Users/ptria/source/repos/Schedule-Maker/backend/src/doctor/repository.py):
    - `stage_doctor(session, name, department_id, team_id)` (in-session flush without commit)
  - [`src/position/repository.py`](file:///C:/Users/ptria/source/repos/Schedule-Maker/backend/src/position/repository.py):
    - `stage_position(session, position_name, duty_days, department_id)` (in-session flush without commit)
  - [`src/schedule/repository.py`](file:///C:/Users/ptria/source/repos/Schedule-Maker/backend/src/schedule/repository.py):
    - `update_schedule_draft_status(session, draft, status, commit=True)` (updates timestamp, status, and single atomic commit)
- **Service** ([`src/schedule/service.py`](file:///C:/Users/ptria/source/repos/Schedule-Maker/backend/src/schedule/service.py)):
  - Implemented `publish_schedule_draft(session, draft_id, department_id)`:
    - Verifies draft belongs to department and is not already published.
    - Soft-deletes prior assignments for the target month (`commit=False`).
    - Pre-caches active doctors, positions, and shifts to eliminate N+1 queries.
    - Auto-provisions missing entities using repository staging functions.
    - Performs bulk insert of new assignments (`commit=False`).
    - Calls `update_schedule_draft_status(..., commit=True)` for single atomic commit.
- **Schemas** ([`src/schedule/schemas.py`](file:///C:/Users/ptria/source/repos/Schedule-Maker/backend/src/schedule/schemas.py)):
  - Updated `ScheduleAssignmentItem` with `day_name: Optional[str] = None` and `doctor_email: Optional[str] = None` for robust serialization.
- **Controller** ([`src/schedule/controllers.py`](file:///C:/Users/ptria/source/repos/Schedule-Maker/backend/src/schedule/controllers.py)):
  - Implemented `publish_schedule_draft_controller`:
    - Validates draft existence (`404 Not Found`).
    - Enforces department tenant isolation (`403 Forbidden`).
    - Translates domain `ValueError` to `400 Bad Request`.
- **Route** ([`src/schedule/routes.py`](file:///C:/Users/ptria/source/repos/Schedule-Maker/backend/src/schedule/routes.py)):
  - Mounted `POST /api/v1/schedules/{draft_id}/publish` protected by `require_department_admin`.

### 3. Tests Verified (Green Phase)
- ACID Rollback & Staging Tests:
  - `test_stage_doctor_flushes_without_committing` in [`test_doctor.py`](file:///C:/Users/ptria/source/repos/Schedule-Maker/backend/src/doctor/tests/test_doctor.py)
  - `test_stage_position_flushes_without_committing` in [`test_position.py`](file:///C:/Users/ptria/source/repos/Schedule-Maker/backend/src/position/tests/test_position.py)
  - `test_stage_shift_flushes_without_committing` in [`test_shift.py`](file:///C:/Users/ptria/source/repos/Schedule-Maker/backend/src/shift/tests/test_shift.py)
  - `test_soft_delete_shift_assignments_commit_false_can_be_rolled_back` in [`test_shift.py`](file:///C:/Users/ptria/source/repos/Schedule-Maker/backend/src/shift/tests/test_shift.py)
  - `test_bulk_create_shift_assignments_commit_false_can_be_rolled_back` in [`test_shift.py`](file:///C:/Users/ptria/source/repos/Schedule-Maker/backend/src/shift/tests/test_shift.py)
  - `test_update_schedule_draft_status_commit` in [`test_models.py`](file:///C:/Users/ptria/source/repos/Schedule-Maker/backend/src/schedule/tests/test_models.py)
  - `test_update_schedule_draft_status_rollback` in [`test_models.py`](file:///C:/Users/ptria/source/repos/Schedule-Maker/backend/src/schedule/tests/test_models.py)
- Service Tests in [`test_service.py`](file:///C:/Users/ptria/source/repos/Schedule-Maker/backend/src/schedule/tests/test_service.py):
  - `test_publish_schedule_draft_success`
  - `test_publish_schedule_draft_already_published_raises_error`
  - `test_publish_schedule_draft_idempotency_cleans_prior_assignments`
  - `test_publish_schedule_draft_not_found_raises_error`
- Route Tests in [`test_routes.py`](file:///C:/Users/ptria/source/repos/Schedule-Maker/backend/src/schedule/tests/test_routes.py):
  - `test_publish_schedule_draft_endpoint_success` (verifies month advance!)
  - `test_publish_schedule_draft_endpoint_viewer_forbidden`
  - `test_publish_schedule_draft_endpoint_cross_department_forbidden`
  - `test_publish_schedule_draft_endpoint_already_published_returns_400`
  - `test_publish_schedule_draft_endpoint_nonexistent_returns_404`

---

## Step 4: Frontend — Remote Data Sources & Repositories `[x] Completed`

### 1. Business & Architectural Goals
- Connected Flutter client directly to backend Step 1, Step 2, and Step 3 endpoints.
- Provided strongly typed domain models and DTO parsing for `ScheduleSummary` and `TargetMonthInfo`.
- Exposed clean repository contracts for listing history, resolving next target months, and publishing drafts.

### 2. Implementation Deliverables
- **Domain Models**:
  - [`ScheduleSummary`](file:///C:/Users/ptria/source/repos/Schedule-Maker/frontend/lib/features/admin/domain/models/schedule_summary.dart): Immutable entity containing metadata (`id`, `departmentId`, `targetMonth`, `status`, `totalDuties`, `solverStatus`, `sourceFilename`, `createdAt`, `updatedAt`) and `isPublished` getter.
  - [`TargetMonthInfo`](file:///C:/Users/ptria/source/repos/Schedule-Maker/frontend/lib/features/admin/domain/models/target_month_info.dart): Immutable entity containing `nextTargetMonth`, `lastPublishedMonth`, and `hasPublishedSchedules` getter.
- **DTOs**:
  - [`ScheduleSummaryDto`](file:///C:/Users/ptria/source/repos/Schedule-Maker/frontend/lib/features/admin/data/dtos/schedule_summary_dto.dart) with `fromJson` and `toDomain()`.
  - [`TargetMonthInfoDto`](file:///C:/Users/ptria/source/repos/Schedule-Maker/frontend/lib/features/admin/data/dtos/target_month_info_dto.dart) with `fromJson` and `toDomain()`.
- **Remote Data Source** ([`schedule_draft_remote_data_source.dart`](file:///C:/Users/ptria/source/repos/Schedule-Maker/frontend/lib/features/admin/data/datasources/schedule_draft_remote_data_source.dart)):
  - `fetchScheduleHistory({String? departmentId})` calling `GET /api/v1/schedules`.
  - `fetchTargetMonthInfo({String? departmentId})` calling `GET /api/v1/schedules/target-month`.
  - `publishScheduleDraft({required String draftId})` calling `POST /api/v1/schedules/$draftId/publish`.
- **Repository** ([`schedule_draft_repository.dart`](file:///C:/Users/ptria/source/repos/Schedule-Maker/frontend/lib/features/admin/data/repositories/schedule_draft_repository.dart)):
  - `fetchScheduleHistory`, `fetchTargetMonthInfo`, and `publishScheduleDraft` exposed and implemented in `ScheduleDraftRepositoryImpl`.

### 3. Tests Verified (Green Phase)
- Model & DTO tests in [`schedule_summary_dto_test.dart`](file:///C:/Users/ptria/source/repos/Schedule-Maker/frontend/test/features/admin/data/dtos/schedule_summary_dto_test.dart) & [`target_month_info_dto_test.dart`](file:///C:/Users/ptria/source/repos/Schedule-Maker/frontend/test/features/admin/data/dtos/target_month_info_dto_test.dart) (15 tests).
- Data Source tests in [`schedule_draft_remote_data_source_test.dart`](file:///C:/Users/ptria/source/repos/Schedule-Maker/frontend/test/features/admin/data/datasources/schedule_draft_remote_data_source_test.dart) (13 tests).
- Repository tests in [`schedule_draft_repository_test.dart`](file:///C:/Users/ptria/source/repos/Schedule-Maker/frontend/test/features/admin/data/repositories/schedule_draft_repository_test.dart) (12 tests).

---

## Step 5: Frontend — State Management & Controller Updates `[x] Completed`

### 1. Business & Architectural Goals
- Eliminated hardcoded `'2026-11'` references from state and controller logic.
- Expanded `ScheduleGenerationState` to hold selected month, target month metadata, schedule history, and publishing loading state.
- Orchestrated dashboard startup (`initializeDashboard`), dynamic month selection (`selectMonth`), and schedule publication with automatic metadata refreshing (`publishCurrentDraft`).

### 2. Implementation Deliverables
- **State Model** ([`schedule_generation_state.dart`](file:///C:/Users/ptria/source/repos/Schedule-Maker/frontend/lib/features/admin/presentation/state/schedule_generation_state.dart)):
  - Added `selectedMonth: String`, `targetMonthInfo: TargetMonthInfo?`, `scheduleHistory: List<ScheduleSummary>`, and `isPublishing: bool`.
  - Added computed getter `isPublished => draft?.isPublished ?? false`.
  - Integrated `listEquals` and `Object.hashAll` for deep value equality.
- **Controller** ([`schedule_generation_controller.dart`](file:///C:/Users/ptria/source/repos/Schedule-Maker/frontend/lib/features/admin/presentation/controllers/schedule_generation_controller.dart)):
  - Implemented `initializeDashboard({String? departmentId})`.
  - Implemented `selectMonth(String targetMonth, {String? departmentId})`.
  - Implemented `publishCurrentDraft({String? departmentId})` with `isPublishing` loading lifecycle and metadata refresh.
  - Dynamic month defaulting in `generateFromExcel`.

### 3. Tests Verified (Green Phase)
- State tests in [`schedule_generation_state_test.dart`](file:///C:/Users/ptria/source/repos/Schedule-Maker/frontend/test/features/admin/presentation/state/schedule_generation_state_test.dart) (6 tests).
- Controller tests in [`schedule_generation_controller_test.dart`](file:///C:/Users/ptria/source/repos/Schedule-Maker/frontend/test/features/admin/presentation/controllers/schedule_generation_controller_test.dart) (17 tests).

---

## Step 6: Frontend — UI Navigation & Publish Action `[x] Completed`

### 1. Business & Architectural Goals
- Enable hospital administrators to switch smoothly between scheduling periods and review past drafts or published schedules directly on the dashboard.
- Display clear, live status indicators distinguishing between in-progress draft schedules and locked published schedules.
- Provide a safe, one-click schedule publication workflow that commits drafts to the central database with modal confirmation.

### 2. Implementation Deliverables
- **Month Selector & Status Badge** ([`AdminHeroCard`](file:///C:/Users/ptria/source/repos/Schedule-Maker/frontend/lib/features/admin/presentation/widgets/admin_hero_card.dart)):
  - Integrated `PopupMenuButton<String>` with an arrow dropdown indicator (`Icons.arrow_drop_down`) next to the month title header.
  - Dynamically renders active and historical months from `availableMonths` with checkmarks on the selected month.
  - Implemented dynamic status badge: `'Published'` with green theme (`Colors.green.shade50`, `Colors.green.shade800`) vs `'Draft'` with amber theme (`Colors.amber.shade50`, `Colors.amber.shade900`).
- **Publish Action** ([`DraftPreviewCanvas`](file:///C:/Users/ptria/source/repos/Schedule-Maker/frontend/lib/features/admin/presentation/widgets/draft_preview_canvas.dart)):
  - Added primary `FilledButton.icon` (`Icons.cloud_upload_outlined`) labeled `'Publish Schedule'` placed next to `'Export to Excel'`.
  - Implemented in-flight loading state with `CircularProgressIndicator` during asynchronous publication.
  - Automatically disables the button (`onPressed: null`) when `isPublishing == true` or when the active draft is already published (`draft.isPublished == true`).
  - Added responsive compact mode (`IconButton` with tooltip) for mobile viewports (`maxWidth < 500`).
- **Screen Integration & Confirmation Dialog** ([`AdminScreen`](file:///C:/Users/ptria/source/repos/Schedule-Maker/frontend/lib/features/admin/presentation/screens/admin_screen.dart)):
  - Replaced hardcoded startup with `initializeDashboard(departmentId: user?.departmentId)` in `initState()`.
  - Implemented `_handlePublishPressed()` displaying an `AlertDialog` confirmation with `'Cancel'` and `'Publish'` actions.
  - Connected `selectMonth()` callback to `AdminHeroCard.onMonthSelected`.
  - Provided feedback via SnackBar upon successful publication.
  - Replaced all legacy hardcoded `'2026-11'` strings with `genState.selectedMonth`.

### 3. Tests Verified (Green Phase)
- **`AdminHeroCard` Widget Tests** ([`admin_hero_card_test.dart`](file:///C:/Users/ptria/source/repos/Schedule-Maker/frontend/test/features/admin/presentation/widgets/admin_hero_card_test.dart)): 7 passed
  - `renders month title, status pill, description, and generate button`
  - `tapping Generate Schedule button invokes onGeneratePressed callback`
  - `shows loading indicator and disables button when isGenerating is true`
  - `renders Published badge when isPublished is true`
  - `renders Draft badge when isPublished is false`
  - `renders popup menu and invokes onMonthSelected when an item is selected`
  - `does not render popup button when availableMonths is empty`
- **`DraftPreviewCanvas` Widget Tests** ([`draft_preview_canvas_test.dart`](file:///C:/Users/ptria/source/repos/Schedule-Maker/frontend/test/features/admin/presentation/widgets/draft_preview_canvas_test.dart)): 16 passed
  - `renders Publish Schedule button when draft is present`
  - `tapping Publish Schedule invokes onPublishPressed callback`
  - `shows loading indicator and disables Publish button when isPublishing is true`
  - `disables Publish Schedule button when draft is already published`
  - (+ 12 existing table, calendar, and export tests)
- **`AdminScreen` Integration Tests** ([`admin_screen_test.dart`](file:///C:/Users/ptria/source/repos/Schedule-Maker/frontend/test/features/admin/presentation/screens/admin_screen_test.dart)): 12 passed
  - `selecting month from popup menu loads draft for selected month`
  - `tapping Publish Schedule shows confirmation dialog and publishes on confirm`
  - (+ 10 existing shell, navigation, and generation tests)
- **Full Test Baselines**:
  - Frontend: **255 / 255 passed**
  - Backend: **772 / 772 passed**
