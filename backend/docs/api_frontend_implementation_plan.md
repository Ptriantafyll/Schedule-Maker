# API and Frontend Implementation Plan (Single Department with Hybrid DB Sync)

This plan integrates the offline-first database architecture described in [database_plan.md](file:///C:/Users/ptria/source/repos/Schedule-Maker/backend/database_plan.md). The application will run a local FastAPI instance communicating with a SQLite database (`hospital_schedule.db`) on the administrator's desktop, and sync data upstream/downstream to a central PostgreSQL instance.

---

## 1. Dual-Database & Sync Architecture

```
┌──────────────────────────────────────┐          ┌──────────────────────────────────────┐
│        Local Desktop (Admin)         │          │             Cloud Server             │
│                                      │          │                                      │
│  ┌───────────────┐  ┌─────────────┐  │  Sync    │  ┌───────────────┐  ┌─────────────┐  │
│  │ Flutter Win   │──│ Local API   │──┼──────────┼─►│ Central API   │──│ Remote DB   │  │
│  │ Desktop App   │  │ (FastAPI)   │  │ (REST)   │  │ (FastAPI)     │  │ (PostgreSQL)│  │
│  └───────────────┘  └─────────────┘  │          │  └───────────────┘  └─────────────┘  │
│          ▲                 │         │          │         ▲                 │          │
│          │ Reads/Writes    ▼         │          │         │ Reads/Writes    ▼          │
│        ┌─────────────────────┐       │          │       ┌─────────────────────┐        │
│        │ Local DB (SQLite)   │       │          │       │ PostgreSQL Database │        │
│        └─────────────────────┘       │          │       └─────────────────────┘        │
└──────────────────────────────────────┘          └──────────────────────────────────────┘
```

### Core Sync Rules (to avoid conflicts):
1.  **UUIDv4 Primary Keys:** Every table uses UUID strings generated on the client-side rather than auto-incrementing integers.
2.  **Sync Audit Metadata:** Every table inherits from `SyncBase` tracking modification times (`updated_at`), soft delete flags (`is_deleted`), and sync states (`sync_status`).
3.  **Soft Deletes:** Deletes on SQLite mark the record `is_deleted = True` and set `sync_status = False` so the change can be synced to PostgreSQL.

---

## 2. Database Schema (SQLModel Framework)

All models inherit from a common `SyncBase` abstract class:

```python
import uuid
from datetime import datetime
from typing import Optional, List
from sqlmodel import Field, SQLModel, Relationship

class SyncBase(SQLModel):
    """Abstract base class tracking synchronization states."""
    id: uuid.UUID = Field(default_factory=uuid.uuid4, primary_key=True, index=True)
    created_at: datetime = Field(default_factory=datetime.utcnow)
    updated_at: datetime = Field(default_factory=datetime.utcnow)
    is_deleted: bool = Field(default=False)
    sync_status: bool = Field(default=False)
```

### Entities

```python
class TeamDoctorLink(SQLModel, table=True):
    team_id: uuid.UUID = Field(foreign_key="team.id", primary_key=True)
    doctor_id: uuid.UUID = Field(foreign_key="doctor.id", primary_key=True)

class PositionEligibleDoctorLink(SQLModel, table=True):
    position_id: uuid.UUID = Field(foreign_key="position.id", primary_key=True)
    doctor_id: uuid.UUID = Field(foreign_key="doctor.id", primary_key=True)

class ScheduleConfig(SyncBase, table=True):
    w_every_other_penalty: int = 4
    w_gap_penalty: int = 2
    w_block_dev_penalty: int = 2
    w_full_wkend_off_bonus: int = 5
    w_balance_full_wkends_off: int = 20
    w_diff_wkend_duty_day: int = 2
    solver_time_limit: int = 120
    max_duties_per_month: int = 8
    month_blocks: int = 3

class Department(SyncBase, table=True):
    name: str = Field(index=True, unique=True)
    code: str

class Team(SyncBase, table=True):
    name: str
    department_id: uuid.UUID = Field(foreign_key="department.id")

class Doctor(SyncBase, table=True):
    name: str
    email: str
    max_consecutive_shifts: int = 5
    department_id: uuid.UUID = Field(foreign_key="department.id")
    
    # Relationships
    unavailabilities: List["Unavailability"] = Relationship(back_populates="doctor")
    pre_assignments: List["PreAssignment"] = Relationship(back_populates="doctor")

class Unavailability(SyncBase, table=True):
    doctor_id: uuid.UUID = Field(foreign_key="doctor.id")
    unavailable_date: datetime  # Tracks specific day numbers as actual dates
    
    doctor: Doctor = Relationship(back_populates="unavailabilities")

class Position(SyncBase, table=True):
    name: str
    duty_days: str = Field(default="[0,1,2,3,4,5,6]")  # JSON list of weekdays
    department_id: uuid.UUID = Field(foreign_key="department.id")

class Shift(SyncBase, table=True):
    name: str
    doctors_per_shift: int = 1
    grants_day_off: bool = False
    position_id: uuid.UUID = Field(foreign_key="position.id")

class PreAssignment(SyncBase, table=True):
    doctor_id: uuid.UUID = Field(foreign_key="doctor.id")
    shift_id: uuid.UUID = Field(foreign_key="shift.id")
    shift_date: datetime
    
    doctor: Doctor = Relationship(back_populates="pre_assignments")

class ScheduleAssignment(SyncBase, table=True):
    """Stores the output assignments from the solver."""
    date: datetime
    doctor_id: uuid.UUID = Field(foreign_key="doctor.id")
    position_id: uuid.UUID = Field(foreign_key="position.id")
    shift_id: uuid.UUID = Field(foreign_key="shift.id")
    department_id: uuid.UUID = Field(foreign_key="department.id")
```

---

## 3. Backend REST API Endpoints

### Authentication & Config
*   `POST /api/v1/auth/login` — Basic token/credentials login for the Department Head.
*   `GET /api/v1/config` / `PUT /api/v1/config` — View and edit the local `ScheduleConfig`.

### Doctors & Unavailability
*   `GET /api/v1/doctors` — List all active doctors (`is_deleted == False`).
*   `POST /api/v1/doctors` — Create or update a doctor (client-generated UUID).
*   `DELETE /api/v1/doctors/{id}` — Soft delete a doctor (`is_deleted = True`).
*   `GET /api/v1/doctors/{id}/unavailability` — Get unavailability dates.
*   `POST /api/v1/doctors/{id}/unavailability` — Set/replace unavailability dates.
*   `GET /api/v1/doctors/{id}/pre-assignments` / `POST /api/v1/doctors/{id}/pre-assignments` — Manage pre-assigned shifts.

### Positions & Shifts
*   `GET /api/v1/positions` — List positions and their associated shifts.
*   `POST /api/v1/positions` — Create/edit a position.
*   `POST /api/v1/positions/{id}/shifts` — Create/edit a shift for a position.

### Teams
*   `GET /api/v1/teams` — List teams and member links.
*   `POST /api/v1/teams` — Create/edit a team.

### Schedule Generation & Exports
*   `POST /api/v1/schedule/generate` — Load parameters from SQLite, build and run CP-SAT model, and save output assignments.
*   `GET /api/v1/schedule` — Retrieve active schedule assignments for a month/year.
*   `GET /api/v1/schedule/export` — Export schedule as a styled Excel sheet.

### Synchronization
*   `GET /api/v1/sync?since=TIMESTAMP` (Downstream Sync)
    *   Retrieves all records from PostgreSQL modified or created after `TIMESTAMP`.
    *   Client downloads and updates local SQLite database.
*   **POST `/api/v1/sync` (Upstream Sync)**
    *   Receives local records where `sync_status == False`.
    *   Saves them to PostgreSQL, resolving conflicts via the latest `updated_at` timestamps.
    *   Returns successful IDs so the client marks them `sync_status = True`.

---

## 4. Frontend Architecture (Flutter Windows App)

The Flutter app located in the [frontend](file:///C:/Users/ptria/source/repos/Schedule-Maker/frontend) directory will be updated to target the desktop environment using a modern, clean navigation structure.

### App Pages / Views

1.  **Login View**
    *   Simple password entry for the department administrator.
2.  **Dashboard / Main View**
    *   **Sidebar Navigation** to switch between sections.
    *   Overview statistics (number of doctors, upcoming month generation status).
3.  **Schedule Viewer**
    *   Interactive grid showing the current month's schedule.
    *   Highlight weekends in black/white as in the Excel exporter.
    *   Action button to "Generate New Schedule" (opens a dialog to select month/year).
    *   Action button to "Export to Excel".
4.  **Doctors & Unavailability Manager**
    *   Table listing all doctors.
    *   Inline form to add a new doctor.
    *   Clicking a doctor opens a calendar view to select/deselect their unavailability dates for the month.
    *   Configure Pre-assignments (assigning specific doctors to specific shifts/days).
5.  **Department Configurations**
    *   List of Positions & Shifts (editable).
    *   List of Teams (editable).
    *   Tweakable weights (`ScheduleConfig` parameters like gap penalties, weekend bonuses).

---

## 5. Phased Implementation Roadmap

### Phase 1: Local Backend & DB Setup (Current Focus)
1.  Add dependencies: `sqlmodel`, `fastapi`, `uvicorn`.
2.  Implement SQLModel data models matching the schema.
3.  Rewrite [scheduler.py](file:///C:/Users/ptria/source/repos/Schedule-Maker/backend/src/scheduler.py) solver entry points to load/save using database sessions.
4.  Develop local FastAPI endpoints for the admin UI to read/write local SQLite configurations.

### Phase 2: Local UI Integration (Windows Desktop)
1.  Configure the Flutter windows application to connect to the local API on `http://localhost:8000`.
2.  Implement Flutter screens: Dashboard, Doctor Setup, Unavailability Calendar, Position & Shift Manager, and Schedule Grid.

### Phase 3: Central Database Deployment
1.  Deploy a cloud-managed PostgreSQL database.
2.  Deploy the FastAPI app in a production environment (server/container) connected to Postgres.

### Phase 4: Sync Engine Construction
1.  Implement the upstream and downstream endpoints `/api/v1/sync` on both local/cloud environments.
2.  Implement a background sync service in Flutter/Python that automatically triggers updates when internet connection is detected.
