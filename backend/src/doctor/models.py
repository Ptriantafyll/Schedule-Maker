"""
Module doctor.models.py

This file defines a self-contained `Doctor` SQLModel table used by the
doctor feature. It's intentionally implemented here so the feature can 
own its model shape independently (useful during refactors or featur
extraction). Keep migrations in sync manually if you change this schema.
"""

from __future__ import annotations

import uuid

from sqlmodel import Field
from src.db.schemas import SyncBase


class Doctor(SyncBase, table=True):
    """
    Represents a hospital doctor

    Fileds mirror the DB-level needs used by the application:
    - `name` 
    - `email`
    - `department_id`
    - `team_id`
    """

    name: str = Field(index=True)
    email: str = Field(index=True, unique=True)
    department_id: uuid.UUID = Field(foreign_key="department.id")
    team_id: uuid.UUID = Field(foreign_key="team.id")
