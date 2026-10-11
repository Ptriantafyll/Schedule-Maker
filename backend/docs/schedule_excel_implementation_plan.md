# Backend Schedule Excel Generation Implementation Plan

## 1. Overview & Objectives

This document specifies the backend implementation for **Excel-based Schedule Generation** in the Schedule-Maker FastAPI backend.

### Goal
Allow hospital administrators to upload an Excel (`.xlsx`) workbook containing doctor rosters, positions, unavailabilities, and pre-assignments, run the Google OR-Tools CP-SAT constraint solver, store the resulting draft schedule in the database (`ScheduleDraft`), and return the solved schedule (with an option to export/download the completed Excel workbook).

### Key Architectural Decisions
1. **Lightweight Database Persistence (No External Object Storage)**:
   - Stores the generated schedule directly in a single `ScheduleDraft` SQLModel table inheriting `SyncBase`.
   - The solved assignments are stored in an indexed `assignments` JSON column (`sa_column=Column(JSON)`).
   - This keeps the system zero-installation and single-file SQLite on desktop, while fully synchronizing with PostgreSQL in the cloud.
2. **Decoupled from Relational Doctor/Shift Tables**:
   - Avoids creating unverified or duplicate doctor/position/shift records in the database.
   - The Excel workbook is self-contained: the solver runs using domain models parsed from the spreadsheet.
3. **Test-Driven Development (TDD)**:
   - Pytest unit tests for the Excel parser and solver service are created before writing implementation code.
   - Mock test fixtures with valid and invalid `.xlsx` workbooks.

---

## 2. Excel Workbook Data Contract

The uploaded `.xlsx` file must conform to the following schema:

### Sheet 1: `Doctors` (Primary Required Sheet)
- **Rows**: Each row represents an individual doctor.
- **Columns**:
  1. `Name` (String, Required): Full name of the doctor (e.g., `Dr. Gregory House`).
  2. `Email` (String, Required): Unique identifier and contact email (e.g., `house@hospital.org`).
  3. `Position` (String, Required): Clinical position or specialty role (e.g., `ER`, `ICU`, `Attending`, `Resident`).
  4. `Team` (String, Optional): Team identifier (e.g., `Team Alpha`) to balance duties across teams.
  5. `Unavailability` (String, Optional): Comma-separated day numbers (e.g., `3, 14, 22`) or ISO dates (`2026-11-03, 2026-11-14`).
  6. `Pre-Assignments` (String, Optional): Fixed locked-in assignments, format `Day:Shift` or `Date:Shift` (e.g., `5:Morning, 12:Night`).
  7. `Max Duties` (Integer, Optional): Custom duty cap override (defaults to 7 or 8 duties per month).

### Sheet 2: `Positions & Shifts` (Optional Configuration Sheet)
If omitted, the solver defaults to 1 doctor per calendar day covering the positions listed in Sheet 1.
- **Columns**:
  1. `Position` (String): Position name matching Sheet 1 (e.g., `ER`).
  2. `Shift Name` (String): Name of shift (e.g., `Morning`, `Night`, `24h Duty`).
  3. `Doctors Per Shift` (Integer): Required coverage count (default `1`).
  4. `Duty Days` (String): Operational days of week (e.g., `0,1,2,3,4,5,6` or `Mon-Sun`).
  5. `Grants Day Off` (Boolean/String): Enforces post-shift rest day (`Yes`/`No` or `TRUE`/`FALSE`).

---

## 3. Database Schema: `ScheduleDraft`

Location: `backend/src/schedule/models.py`

```python
import uuid
from typing import Optional
from sqlmodel import Field, Column, JSON
from src.db.schemas import SyncBase

class ScheduleDraft(SyncBase, table=True):
    """Stores generated schedule drafts directly in SQLite / PostgreSQL."""
    __tablename__ = "schedule_draft"

    department_id: uuid.UUID = Field(foreign_key="department.id", index=True)
    target_month: str = Field(index=True)       # Format: "YYYY-MM" (e.g. "2026-11")
    source_filename: str                        # e.g. "november_roster.xlsx"
    total_duties: int = 0
    solver_status: str = "OPTIMAL"              # "OPTIMAL" | "FEASIBLE" | "INFEASIBLE"
    status: str = Field(default="draft")        # "draft" | "published" | "archived"
    
    # Solved schedule assignments stored as JSON
    assignments: list[dict] = Field(default_factory=list, sa_column=Column(JSON))
```

### Assignment JSON Item Structure
```json
{
  "date": "2026-11-01",
  "day_name": "Sunday",
  "doctor_name": "Dr. Gregory House",
  "doctor_email": "house@hospital.org",
  "position": "ER",
  "shift": "Night"
}
```

---

## 4. API Endpoints Specification

Location: `backend/src/schedule/routes.py` (mounted at `/api/v1/schedules`)

### 1. `POST /api/v1/schedules/generate-from-excel`
- **Content-Type**: `multipart/form-data`
- **Parameters**:
  - `file`: `UploadFile` (binary `.xlsx` workbook)
  - `department_id`: `UUID` (Target hospital department)
  - `target_month`: `str` (Format: `YYYY-MM`)
- **Process**:
  1. Validate file extension (`.xlsx`).
  2. Parse spreadsheet sheets and validate required columns.
  3. Instantiate `ShiftScheduler` or `schedule_maker` solver with the parsed domain objects.
  4. Run CP-SAT solver with 60s timeout.
  5. If infeasible, return HTTP 422 with diagnostic error message.
  6. Upsert the generated schedule into `schedule_draft` table for the department and month.
  7. Return `ScheduleDraftRead` JSON response.

### 2. `GET /api/v1/schedules/draft`
- **Parameters**:
  - `department_id`: `UUID`
  - `target_month`: `str` (`YYYY-MM`)
- **Returns**: Active `ScheduleDraft` for the department or HTTP 404 if none exists.

### 3. `GET /api/v1/schedules/export-excel`
- **Parameters**:
  - `draft_id`: `UUID`
- **Returns**: Downloadable `.xlsx` file with color-coded weekend shifts, doctor rows, and date columns.

---

## 5. Step-by-Step Implementation Plan

### Step 1: Feature Package & Model Setup (COMPLETED)
- Create directory `backend/src/schedule/`.
- Create `models.py` with `ScheduleDraft` and Pydantic schemas (`ScheduleDraftRead`, `ScheduleAssignmentItem`).
- Register `ScheduleDraft` in `backend/src/db/connection.py` so database tables are initialized.
- Verified with unit tests in `test_models.py` (4/4 passed).

### Step 2: Excel Parser Utility (`excel_parser.py`) (COMPLETED)
- **TDD Test**: `backend/src/schedule/tests/test_excel_parser.py` (10/10 passed).
- Implemented `parse_excel_schedule_workbook(file_bytes, department_name, target_month)`:
  - Validates sheet `Doctors` and column names.
  - Supports 2-sheet workbook (`Doctors` and `Shifts`) with flexible weekday range parsing.
  - Graceful fallback when `Shifts` sheet is omitted (1 doctor per day with day-off grant).
  - Strict vs flexible pre-assignments parsing.
  - Converts into `models.Department`, `models.Doctor`, `models.Position`, and `models.Shift`.

### Step 3: Solver Integration Service (`service.py`) (COMPLETED)
- **TDD Test**: `backend/src/schedule/tests/test_service.py` (6/6 passed).
- Implemented `generate_schedule_from_excel(session, file_bytes, department_id, target_month, source_filename)`:
  - Uses Google OR-Tools CP-SAT solver.
  - Enforces hard constraints: 1 shift per day per doctor, weekend duty limits, full weekend off.
  - Enforces soft preferences: block balance, every-other day penalty.
  - Transforms solver solution to list of assignment dicts.
  - Persists directly using `schedule_repository.create_schedule_draft`.

### Step 4: Controller & FastAPI Router Integration (`controllers.py` & `routes.py`) (COMPLETED)
- **TDD Tests**:
  - `backend/src/schedule/tests/test_controllers.py` (14/14 passed).
  - `backend/src/schedule/tests/test_routes.py` (12/12 passed).
- **Controllers Implementation (`controllers.py`)**:
  - Centralized all exception handling and input validation in `controllers.py`.
  - Enforced tenant scope boundaries with `_resolve_and_verify_department`: non-super-admins cannot target other departments (HTTP 403), super-admins must provide a target department (HTTP 400).
  - File extension validation: rejects non-`.xlsx` filenames (HTTP 400).
  - `generate_schedule_from_excel_controller`: Validates inputs, coordinates service generation, and translates domain `ValueError` to `HTTPException(400)`.
  - `get_schedule_draft_controller`: Queries draft by ID, translates missing draft to `HTTPException(404)`.
  - `get_schedule_draft_for_department_month_controller`: Queries active draft for department and month, translates missing draft to `HTTPException(404)`.
  - `export_schedule_draft_controller`: Queries draft by ID, enforces tenant scope against `draft.department_id`, and generates Excel download bytes.
- **Routes Implementation (`routes.py`)**:
  - Pure declarative HTTP wiring (100% free of `HTTPException`).
  - `POST /api/v1/schedules/generate-from-excel` (201 Created).
  - `GET /api/v1/schedules/draft` (200 OK).
  - `GET /api/v1/schedules/export-excel` (200 OK, attachment download).
  - Mounted `schedule_routes.router` in `backend/src/main.py` under `/api/v1`.
- **Excel Exporter Service (`service.py`)**:
  - `export_schedule_draft_to_excel`: Formats `ScheduleDraft` assignments into a structured `.xlsx` calendar grid.
  - Highlights doctor unavailability with solid black fill (`#000000`).
  - Highlights weekends with light gray fill (`#F2F2F2`).
  - Computes monthly total duties per doctor in a summary column.
  - Generates binary `.xlsx` streams in-memory via `io.BytesIO()`.
  - Added `unavailabilities` dictionary to `ScheduleDraft` model and `ScheduleDraftRead` schema.
  - `backend/src/schedule/tests/test_service.py` (10/10 passed).

### Step 5: Verification & Full Suite (COMPLETED)
- Run `pytest backend/src/schedule/tests/`: **50/50 tests passed**.
- Run full backend test suite: **727/727 tests passed** with zero regressions across all modules.


