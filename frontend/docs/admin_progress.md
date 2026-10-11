# Admin Feature Progress & Step-by-Step Implementation

This document tracks the end-to-end implementation of the **Admin Feature & Schedule Generation Dashboard** (`features/admin`) in the Flutter frontend, based on [`admin.png`](file:///C:/Users/ptria/Downloads/admin.png) and following **MVVM** and **Riverpod 2.x**.

See the full [Admin Implementation Plan](frontend_admin_implementation_plan.md) for architectural specifications.

---

## Progress Overview

| Phase | Description | Status |
| :--- | :--- | :--- |
| **Phase 1** | Shell, Navigation & Profile Drawer (`ProfileDrawer`, `BottomNavBar`, `AdminScreen`) | `[x] Completed` |
| **Phase 2** | Dashboard UI Components (`AdminHeroCard`, `AdminMetricCards`, `DraftPreviewCanvas`) | `[x] Completed` |
| **Phase 3** | Schedule Generation Modalities (Manual Entry & Excel `.xlsx` Upload) | `[x] Completed` |
| **Phase 4** | Backend Verification & Live Full-Stack Integration | `[x] Completed` |
---

## Phase 1: Shell, Navigation & Profile Drawer

### Step 1.1: `ProfileDrawer` Widget (Shared)
- **Status:** `[x] Completed`
- **Business Reason:** Any logged-in user (doctor, nurse, or administrator) needs rapid access to their personal identity profile, active hospital role, and secure sign-out without cluttering the main operational canvas.
- **Technical Implementation Needed:**
  - Create `lib/shared/widgets/profile_drawer.dart`.
  - Reads `currentUserProvider` to display the user's `fullName`, `email`, and formatted `role.displayName`.
  - Includes user avatar header, account info list tiles, and a prominent Logout action invoking `ref.read(authControllerProvider.notifier).logout()`.
  - **TDD Test:** Create `test/shared/widgets/profile_drawer_test.dart` verifying full name rendering, role display, fallback state, and logout invocation.

---

### Step 1.2: `BottomNavBar` Widget (Shared)
- **Status:** `[x] Completed`
- **Business Reason:** A persistent bottom navigation bar adhering to hospital tablet/mobile navigation standards, giving all users 1-tap switching between Dashboard, Schedule, Requests, and (for administrators) Admin control planes.
- **Technical Implementation Needed:**
  - Create `lib/shared/widgets/bottom_nav_bar.dart`.
  - Material 3 `NavigationBar` with destinations:
    1. Dashboard (`Icons.grid_view`)
    2. Schedule (`Icons.calendar_month`)
    3. Requests (`Icons.swap_horiz`)
    4. Admin (`Icons.admin_panel_settings`, rendered only when `showAdminTab: true`)
  - Takes `int currentIndex`, `bool showAdminTab`, and `ValueChanged<int> onTabSelected`.
  - **TDD Test:** Create `test/shared/widgets/bottom_nav_bar_test.dart` verifying destination labels, icon rendering, conditional admin tab display, and tap callbacks.

---

### Step 1.3: `AdminScreen` Scaffold Integration
- **Status:** `[x] Completed`
- **Business Reason:** The root admin screen combining the top header bar, avatar button (opening the drawer), responsive body, and persistent bottom navigation bar.
- **Technical Implementation Needed:**
  - Create `lib/features/admin/presentation/screens/admin_screen.dart`.
  - AppBar with "MedShift Admin" title, circular profile avatar leading action, and notification bell action.
  - Attached `Drawer: const ProfileDrawer()`.
  - Tapping avatar triggers `Scaffold.of(context).openDrawer()`.
  - Attached `bottomNavigationBar: BottomNavBar(currentIndex: 3, showAdminTab: true, ...)`.
  - **TDD Test:** Create `test/features/admin/presentation/screens/admin_screen_test.dart`.

---

## Phase 2: Dashboard UI Components (`admin.png` Mockup Fidelity)

### Step 2.1: `AdminHeroCard` (Month Header & Generate Trigger)
- **Status:** `[x] Completed`
- **Business Reason:** The central call-to-action widget that shows the scheduling period, solver readiness status, and the primary "Generate Schedule" button.
- **Technical Implementation Needed:**
  - Create `lib/features/admin/presentation/widgets/admin_hero_card.dart`.
  - **TDD Test:** Create `test/features/admin/presentation/widgets/admin_hero_card_test.dart`.

---

### Step 2.2: `AdminMetricCards` (4-Card Pre-Flight System)
- **Status:** `[x] Completed`
- **Business Reason:** Gives the hospital administrator a 360-degree operational picture before triggering generation: timeline urgency (due date), approval blockers (pending leaves), workforce capacity (staff pool), and schedule risk areas (coverage bottlenecks).
- **Technical Implementation Needed:**
  - Create `lib/features/admin/presentation/widgets/admin_metric_cards.dart`.
  - 4 Cards:
    1. Schedule Due Date (`dueDate`, `daysLeft`, `submissionStatus`)
    2. Pending Approvals (`pendingApprovalsCount`, `pendingApprovalsStatus`)
    3. Staff Pool & Capacity (`availableStaffCount`, `totalShiftsCount`, `staffCapacitySubtitle`)
    4. Coverage Bottlenecks (`bottleneckDaysCount`, `bottleneckStatus`, `bottleneckDates`)
  - **TDD Test:** Create `test/features/admin/presentation/widgets/admin_metric_cards_test.dart` verifying all 4 cards, stats, badges, and responsive layout.

---

### Step 2.3: `DraftPreviewCanvas` (Awaiting Generation & Schedule View)
- **Status:** `[x] Completed`
- **Business Reason:** Immediate visual confirmation of generated duty assignments allowing the admin to inspect coverage, view conflicts, and finalize the schedule.
- **Technical Implementation Needed:**
  - Create `lib/features/admin/presentation/widgets/draft_preview_canvas.dart`.
  - **TDD Test:** Create `test/features/admin/presentation/widgets/draft_preview_canvas_test.dart`.

---

## Phase 3: Schedule Generation Modalities

### Step 3.1: Generation Mode Selection Dialog
- **Status:** `[x] Completed`
- **Technical Implementation:**
  - Created `lib/features/admin/presentation/widgets/generation_mode_dialog.dart`.
  - Added `GenerationMode` enum (`currentRoster`, `excelUpload`).
  - Implemented Material 3 `AlertDialog` with options for current department roster or Excel file upload.
  - Connected `onGeneratePressed` in `AdminScreen` to launch `GenerationModeDialog`.
  - **TDD Test:** Created `test/features/admin/presentation/widgets/generation_mode_dialog_test.dart` and integrated in `test/features/admin/presentation/screens/admin_screen_test.dart` (all passing).

---

### Step 3.2: `ExcelUploadDialog` (Spreadsheet Import Flow)
- **Status:** `[x] Completed`
- **Dependency Added:** `file_picker: ^13.1.0` (using `FilePicker.pickFile` directly).
- **Technical Implementation:**
  - Created `lib/features/admin/presentation/widgets/excel_upload_dialog.dart`.
  - Provided `onPickFile` dependency-injection hook for headless, fast unit testing without native OS blocking.
  - Displays drag/browse outline box, file details (name, formatted size), clear/remove action, and conditional generate schedule button.
  - Connected in `AdminScreen` under `GenerationMode.excelUpload` case to toggle draft preview and display confirmation snackbar.
  - **TDD Test:** Created `test/features/admin/presentation/widgets/excel_upload_dialog_test.dart` (all 5 tests passing) and integrated in `admin_screen_test.dart`.

---

### Step 3.3: `ScheduleGenerationController` (Riverpod ViewModel)
- **Status:** `[x] Completed`
- **Technical Implementation:**
  - Created `lib/features/admin/presentation/controllers/schedule_generation_controller.dart`.
  - Defined `GenerationStatus` (`idle`, `solving`, `success`, `error`) and immutable `ScheduleGenerationState`.
  - Implemented `ScheduleRepository` contract with `DefaultScheduleRepository` (enforcing Backend Readiness Gate).
  - Wired into `AdminScreen`: drives `AdminHeroCard(isGenerating: genState.isSolving)` and `DraftPreviewCanvas(isGenerated: genState.isGenerated)`.
  - **TDD Test:** Created `test/features/admin/presentation/controllers/schedule_generation_controller_test.dart` (all 5 tests passing).

---

### Step 3.4: Manual Roster & Operational Setup Flow
- **Status:** `[ ] Deferred to Backlog (BL-003)`
- **Notes:** Full manual creation requires positions, shifts, unavailabilities, and pre-assignments. Deferred to focus on Phase 4 Excel-based generation.

---

## Phase 4: Excel-Based Solver Route & Live Full-Stack Integration

### Step 4.1: Backend Excel Solver Endpoint & Exporter (`backend/src/`)
- **Status:** `[x] Completed`
- **Technical Implementation Completed:**
  - Created `POST /api/v1/schedules/generate-from-excel` in FastAPI accepting `.xlsx` multipart upload.
  - Implemented workbook parser validating columns, range unavailabilities, and pre-assignments.
  - Implemented CP-SAT solver service optimizing duty fairness and rest periods.
  - Stored output in `ScheduleDraft` table (SQLite/PostgreSQL) with JSON assignments and unavailabilities.
  - Created `GET /api/v1/schedules/draft` to query active drafts by department and month.
  - Created `GET /api/v1/schedules/export-excel` downloading formatted `.xlsx` with blacked-out unavailabilities and weekend highlights.
  - **TDD Test Suite:** 50/50 schedule module unit/controller/route tests passed; full backend test suite: 727/727 tests passed.

---

### Step 4.2: Domain Models & Serialization (`features/admin/domain/models/` & `data/dtos/`)
- **Status:** `[x] Completed`
- **Technical Implementation Completed:**
  - Created pure domain models `ScheduleAssignment` and `ScheduleDraft` in `lib/features/admin/domain/models/schedule_draft.dart` (zero Flutter/UI dependencies).
  - Created data transfer objects `ScheduleAssignmentDto` and `ScheduleDraftDto` in `lib/features/admin/data/dtos/schedule_draft_dto.dart` with `fromJson`, `toJson`, and `toDomain()` mappings matching the FastAPI backend schema.
  - **TDD Test Suite:** Created `test/features/admin/domain/models/schedule_draft_test.dart` (5/5 passed) and `test/features/admin/data/dtos/schedule_draft_dto_test.dart` (7/7 passed).

---

### Step 4.3: Remote Data Source & Repository (`features/admin/data/`)
- **Status:** `[x] Completed`
- **Technical Implementation Completed:**
  - Created `ScheduleDraftRemoteDataSource` in `lib/features/admin/data/datasources/schedule_draft_remote_data_source.dart` using Dio and `file.readAsBytes()` (from `file_picker: ^13.1.0`) with methods:
    - `generateFromExcel(file, targetMonth, departmentId)` -> `Future<ScheduleDraft>`
    - `getActiveDraft(targetMonth, departmentId)` -> `Future<ScheduleDraft?>` (handles 404 gracefully)
    - `exportExcel(targetMonth, departmentId)` -> `Future<List<int>>`
  - Created `ScheduleDraftRepository` interface and `ScheduleDraftRepositoryImpl` in `lib/features/admin/data/repositories/schedule_draft_repository.dart`, correctly forwarding `departmentId` for Super Admin and Department Admin workflows.
  - **TDD Test Suite:** Created `schedule_draft_remote_data_source_test.dart` (6/6 passed) and `schedule_draft_repository_test.dart` (6/6 passed).

---

### Step 4.4: Controller Live Integration (`features/admin/presentation/`)
- **Status:** `[x] Completed`
- **Technical Implementation Completed:**
  - Extracted presentation state to `lib/features/admin/presentation/state/schedule_generation_state.dart` with `GenerationStatus`, immutable `ScheduleGenerationState`, `draft`, `isExporting`, and safe `clearDraft` / `clearError` copyWith semantics.
  - Removed legacy in-file mock repository and wired `ScheduleGenerationController` to `scheduleDraftRepositoryProvider`.
  - Implemented `loadActiveDraft(...)`, `generateFromExcel(...)`, and `exportCurrentDraft(...)` with guaranteed `finally` cleanup of `isExporting`.
  - Enforced strict $\le 3$ indentation depth (`BL-013`).
  - **TDD Test Suite:** Created `schedule_generation_state_test.dart` (5/5 passed) and updated `schedule_generation_controller_test.dart` (10/10 passed).

---

### Step 4.5: `DraftPreviewCanvas` Table & Calendar Views & Export Action
- **Status:** `[x] Completed`
- **Technical Implementation Completed:**
  - **Step 4.5a (Table View & Export Button)**:
    - View mode toggle (`CanvasViewMode.table`, `CanvasViewMode.calendar`).
    - Added "Export to Excel" action button in canvas header with loading spinner when `isExporting` is active, enabled when draft is present.
    - Compact icon layout support on mobile widths with tooltips.
    - Render solver metadata summary bar (status chip e.g. `OPTIMAL`, total duties count, source filename).
    - Render scrollable `DataTable` displaying columns: Date, Day, Doctor, Position, Shift.
  - **Step 4.5b (Calendar View with Green Shift Dots & Day Details)**:
    - Monthly calendar grid for target month with `mainAxisExtent: 48` preventing mobile vertical overflow.
    - Green indicator dot per assigned shift on each date.
    - Day selection interaction displaying doctor & shift details for that date.
  - **Step 4.5c (AdminScreen Startup & Export Integration)**:
    - `AdminScreen.initState()` automatically invokes `controller.loadActiveDraft(...)` for the current admin's department.
    - Export button triggers `controller.exportCurrentDraft(...)` with cross-platform file saving (`FilePicker.platform.saveFile` on desktop/web, share sheet / save on mobile).
    - Fixed query param (`draft_id`) and binary response type (`ResponseType.bytes`) in `ApiClient.get`.
    - **TDD Test Suite:** `draft_preview_canvas_test.dart` (12/12 passed) and `admin_screen_test.dart` (all passed).

---

### Step 4.6: End-to-End Test Suite Verification & Flutter Analysis
- **Status:** `[x] Completed`
- **Verification Results:**
  - `flutter analyze`: **0 issues found** across all frontend directories.
  - `flutter test`: **All 208/208 tests passed** with 100% green status.
  - **Live Full-Stack Verification**: Successfully generated a complete 10-doctor schedule from `sample_hospital_schedule_10_doctors.xlsx` in Emergency Medicine department, rendered both Table and Calendar views responsively, and exported `schedule_2026-11.xlsx` via live FastAPI backend.
