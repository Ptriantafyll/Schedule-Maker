"""
Module doctor.models.py

This file defines a self-contained `Doctor` SQLModel table used by the
doctor feature. It's intentionally implemented here so the feature can 
own its model shape independently (useful during refactors or featur
extraction). Keep migrations in sync manually if you change this schema.
"""

from __future__ import annotations

import uuid
import datetime
from sqlmodel import Field

from src.db.schemas import SyncBase


class Doctor(SyncBase, table=True):
    """
    Represents a hospital doctor

    Fields mirror the DB-level needs used by the application:
    - `name` 
    - `email` is indexed and unique
    - `department_id` references a Department row
    - `team_id` references a Team row
    """

    name: str = Field(index=True)
    email: str = Field(index=True, unique=True)
    department_id: uuid.UUID = Field(foreign_key="department.id")
    team_id: uuid.UUID = Field(foreign_key="team.id")


class DoctorUnavailability(SyncBase, table=True):
    """Tracks specific dates a doctor cannot work for a given month."""
    doctor_id: uuid.UUID = Field(foreign_key="doctor.id")
    date: datetime.date


class DoctorPreAssignment(SyncBase, table=True):
    """Hard constraints: locked-in (date, shift) assignments before solver runs."""
    doctor_id: uuid.UUID = Field(foreign_key="doctor.id")
    shift_id: uuid.UUID = Field(foreign_key="shift.id")
    date: datetime.date


class DoctorPosition(SyncBase, table=True):
    """Association table for the many-to-many relationship between doctors and positions."""
    doctor_id: uuid.UUID = Field(foreign_key="doctor.id")
    position_id: uuid.UUID = Field(foreign_key="position.id")
