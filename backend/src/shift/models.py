"""Feature-local ORM models for the `shift` feature.

This file defines a self-contained `Shift` SQLModel table used by the
shift feature. It's intentionally implemented here so the feature can
own its model shape independently (useful during refactors or feature
extraction). Keep migrations in sync manually if you change this schema.
"""
from __future__ import annotations

import uuid
import datetime

from sqlmodel import Field
from src.db.schemas import SyncBase


class Shift(SyncBase, table=True):
    """Represents a work shift stored in the database."""
    name: str = Field(index=True, unique=True)
    doctors_per_shift: int = 1
    grants_day_off: bool = False
    position_id: uuid.UUID = Field(foreign_key="position.id")


class ShiftAssignment(SyncBase, table=True):
    """Represents the final schedule assignments after the solver runs."""
    doctor_id: uuid.UUID = Field(foreign_key="doctor.id")
    shift_id: uuid.UUID = Field(foreign_key="shift.id")
    date: datetime.date  # ISO Format: YYYY-MM-DD
