"""
SQLModel database table definitions for the Schedule-Maker application.
These models are used by the FastAPI layer to interact with the local SQLite
database and the remote PostgreSQL database.

Note: The scheduler (scheduler.py) still uses the plain dataclass models in
models.py. These db_models will eventually replace them once the API layer is
fully wired up.
"""

import uuid
import datetime
from sqlmodel import Column, SQLModel, Field, JSON

class SyncBase(SQLModel):
    """Abstract base class providing synchronization metadata for all tables.

    All database tables inherit from this class to get:
    - A client-generated UUID primary key (avoids ID collisions during sync)
    - Timestamps for conflict resolution during sync
    - Soft delete support (is_deleted flag instead of physical row deletion)
    - Sync status tracking (sync_status flag to know what needs to be pushed)
    """
    id: uuid.UUID = Field(default_factory=uuid.uuid4, primary_key=True)
    created_at: datetime.datetime = Field(
        default_factory=lambda: datetime.datetime.now(datetime.timezone.utc)
    )
    updated_at: datetime.datetime = Field(
        default_factory=lambda: datetime.datetime.now(datetime.timezone.utc)
    )
    is_deleted: bool = Field(default=False)
    sync_status: bool = Field(default=False)


class ScheduleConfig(SyncBase, table=True):
    """Represents the scheduling configuration stored in the database."""

    department_id: uuid.UUID = Field(
        foreign_key="department.id",
        unique=True,
        ondelete="CASCADE"
    )

    w_every_other_penalty: int = 4
    w_gap_penalty: int = 2
    w_block_dev_penalty: int = 2
    w_full_wkend_off_bonus: int = 5
    w_balance_full_wkends_off: int = 20
    w_diff_wkend_duty_day: int = 2
    solver_time_limit: int = 120
    max_duties_per_month: int = 8
    month_blocks: int = 3


class Shift(SyncBase, table=True):
    """Represents a work shift stored in the database."""
    name: str
    doctors_per_shift: int = 1
    grants_day_off: bool = False
    position_id: uuid.UUID = Field(foreign_key="position.id")



class DoctorUnavailability(SyncBase, table=True):
    """Tracks specific dates a doctor cannot work for a given month."""
    doctor_id: uuid.UUID = Field(foreign_key="doctor.id")
    date: str  # ISO Format: YYYY-MM-DD


class DoctorPreAssignment(SyncBase, table=True):
    """Hard constraints: locked-in (date, shift) assignments before solver runs."""
    doctor_id: uuid.UUID = Field(foreign_key="doctor.id")
    shift_id: uuid.UUID = Field(foreign_key="shift.id")
    date: str  # ISO Format: YYYY-MM-DD


class Position(SyncBase, table=True):
    """Represents a position that needs to be staffed, such as "ER" or "ICU", stored in the database."""
    name: str
    department_id: uuid.UUID = Field(foreign_key="department.id")
    duty_days: list[int] = Field(default_factory=list, sa_column=Column(JSON))


class DoctorPosition(SyncBase, table=True):
    """Association table for the many-to-many relationship between doctors and positions."""
    doctor_id: uuid.UUID = Field(foreign_key="doctor.id")
    position_id: uuid.UUID = Field(foreign_key="position.id")


class ShiftAssignment(SyncBase, table=True):
    """Represents the final schedule assignments after the solver runs."""
    doctor_id: uuid.UUID = Field(foreign_key="doctor.id")
    shift_id: uuid.UUID = Field(foreign_key="shift.id")
    date: str  # ISO Format: YYYY-MM-DD
