# Frontend Schedule Excel Integration Plan

## 1. Overview & Objectives

This document specifies the frontend architecture and implementation steps to connect the Flutter admin dashboard to the live backend Excel solver endpoint (`POST /api/v1/schedules/generate-from-excel`) and display generated schedules on the **`DraftPreviewCanvas`**.

### Goal
1. Allow administrators to pick an `.xlsx` file via `ExcelUploadDialog`.
2. Upload the file to the FastAPI backend, trigger the CP-SAT solver, and receive the solved `ScheduleDraft` JSON.
3. Automatically load and render any existing draft from the database on `AdminScreen` startup so the draft survives page reloads.
4. Render the solved schedule inside `DraftPreviewCanvas` in a clean, clinical table/calendar view with an **"Export to Excel"** download action.

---

## 2. Architecture & Data Contracts

### 2.1 Domain Models
Location: `frontend/lib/features/admin/domain/models/schedule_draft.dart`

```dart
class ScheduleAssignment {
  const ScheduleAssignment({
    required this.date,
    required this.dayName,
    required this.doctorName,
    required this.doctorEmail,
    required this.position,
    required this.shift,
  });

  final String date;        // Format: "YYYY-MM-DD"
  final String dayName;     // e.g. "Monday"
  final String doctorName;  // e.g. "Dr. Gregory House"
  final String doctorEmail; // e.g. "house@hospital.org"
  final String position;    // e.g. "ER"
  final String shift;       // e.g. "Night"

  factory ScheduleAssignment.fromJson(Map<String, dynamic> json);
  Map<String, dynamic> toJson();
}

class ScheduleDraft {
  const ScheduleDraft({
    required this.id,
    required this.departmentId,
    required this.targetMonth,
    required this.sourceFilename,
    required this.totalDuties,
    required this.status,
    required this.assignments,
    required this.createdAt,
  });

  final String id;
  final String departmentId;
  final String targetMonth;
  final String sourceFilename;
  final int totalDuties;
  final String status;
  final List<ScheduleAssignment> assignments;
  final DateTime createdAt;

  factory ScheduleDraft.fromJson(Map<String, dynamic> json);
  Map<String, dynamic> toJson();
}
```

---

## 3. Layered Implementation Breakdown

### 3.1 Data Layer (`features/admin/data/`)
1. **`ScheduleRemoteDataSource`** (`sources/schedule_remote_data_source.dart`):
   - Uses `Dio` HTTP client with authentication interceptors.
   - `Future<ScheduleDraft> generateFromExcel({required PlatformFile file, required String departmentId, required String month})`:
     - Creates `FormData.fromMap({'file': MultipartFile.fromBytes(...), 'department_id': ..., 'target_month': ...})`.
     - Calls `POST /api/v1/schedules/generate-from-excel`.
   - `Future<ScheduleDraft?> getActiveDraft({required String departmentId, required String month})`:
     - Calls `GET /api/v1/schedules/draft?department_id=...&target_month=...`.
     - Returns `null` if HTTP 404 (no draft generated yet).
   - `Future<List<int>> downloadExcel({required String draftId})`:
     - Calls `GET /api/v1/schedules/export-excel?draft_id=...` returning file bytes.
2. **`ScheduleRepositoryImpl`** (`repositories/schedule_repository_impl.dart`):
   - Implements `ScheduleRepository` and maps remote data source responses.

---

### 3.2 State Management (`features/admin/presentation/controllers/`)
Update `schedule_generation_controller.dart`:
- Add `ScheduleDraft? draft` to `ScheduleGenerationState`:
  ```dart
  class ScheduleGenerationState {
    const ScheduleGenerationState({
      this.status = GenerationStatus.idle,
      this.draft,
      this.errorMessage,
      this.generatedAt,
    });

    final GenerationStatus status;
    final ScheduleDraft? draft;
    final String? errorMessage;
    final DateTime? generatedAt;

    bool get isGenerated => draft != null;
    bool get isSolving => status == GenerationStatus.solving;
    String get source => draft?.sourceFilename ?? '';
  }
  ```
- Add `loadActiveDraft(String departmentId, String month)` to hydrate existing draft on startup.
- Update `generateFromExcel(PlatformFile file)` to pass the file through the repository to the live API.

---

### 3.3 Presentation Layer (`features/admin/presentation/widgets/`)
1. **`DraftPreviewCanvas` Enhancements**:
   - Header: Add **"Export Excel"** `FilledButton.tonalIcon` or `IconButton` in the action row.
   - Generated State: Instead of static placeholder text, render a responsive Material 3 `DataTable` or scrollable list displaying:
     - Date & Day (e.g. `Nov 1 (Sun)`)
     - Assigned Doctor Name (`Dr. Gregory House`)
     - Position Badge (`ER`)
     - Shift (`Night`)
2. **`AdminScreen` Startup Hydration**:
   - In `initState()` or post-frame callback, read the current user's `departmentId` and call `controller.loadActiveDraft(...)` so previous drafts appear immediately.

---

## 4. Step-by-Step TDD Plan

### Step 1: Models & Serialization (TDD) (COMPLETED)
- **Test**: `test/features/admin/domain/models/schedule_draft_test.dart` and `test/features/admin/data/dtos/schedule_draft_dto_test.dart` (12/12 passed).
- Verified JSON parsing and serialization for `ScheduleAssignment` and `ScheduleDraft`.

### Step 2: Remote Data Source & Repository (TDD) (COMPLETED)
- **Test**: `test/features/admin/data/datasources/schedule_draft_remote_data_source_test.dart` and `test/features/admin/data/repositories/schedule_draft_repository_test.dart` (12/12 passed).
- Multipart upload to `POST /api/v1/schedules/generate-from-excel` using `await file.readAsBytes()`.
- Fetching active draft via `GET /api/v1/schedules/draft`.
- Downloading Excel via `GET /api/v1/schedules/export-excel`.

### Step 3: Controller Live Integration (TDD) (COMPLETED)
- **Test**: `test/features/admin/presentation/state/schedule_generation_state_test.dart` (5/5 passed) and `test/features/admin/presentation/controllers/schedule_generation_controller_test.dart` (10/10 passed).
- Extracted `ScheduleGenerationState` to `presentation/state/`.
- Tested `loadActiveDraft`, `generateFromExcel`, and `exportCurrentDraft` updating state with returned `ScheduleDraft`.

### Step 4: `DraftPreviewCanvas` Table & Calendar Views & Export (TDD) (COMPLETED)
- **Test**: `test/features/admin/presentation/widgets/draft_preview_canvas_test.dart` (12/12 passed) and `admin_screen_test.dart`.
- Verified table and calendar rendering with multiple doctor assignments, column headers, status badges, mobile responsiveness, and export button tap callback.

### Step 5: Full Integration Verification (COMPLETED)
- Ran `flutter test`: 208/208 tests passed.
- Ran `flutter analyze`: 0 issues found.
- Live end-to-end integration verified: solved draft loaded, rendered in Table and Calendar modes, and exported to `.xlsx` cleanly.
