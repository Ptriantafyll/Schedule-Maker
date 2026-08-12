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
from sqlmodel import SQLModel, Field 


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
