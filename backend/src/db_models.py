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


class Department(SyncBase, table=True):
    """Represents a hospital department stored in the database."""
    name: str = Field(index=True, unique=True)
    code: str


class Doctor(SyncBase, table=True):
    """Represents a doctor stored in the database."""
    name: str
    email: str
    department_id: uuid.UUID = Field(foreign_key="department.id")
