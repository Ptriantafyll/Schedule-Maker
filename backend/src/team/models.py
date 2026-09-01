"""Feature-local ORM models for the `team` feature.

This file defines a self-contained `Team` SQLModel table used by the
team feature. It's intentionally implemented here so the feature can
own its model shape independently (useful during refactors or feature
extraction). Keep migrations in sync manually if you change this schema.
"""
from __future__ import annotations

import uuid

from sqlmodel import Field
from sqlalchemy import UniqueConstraint
from src.db.schemas import SyncBase


class Team(SyncBase, table=True):
    """Represents a hospital team.

    Fields mirror the DB-level needs used by the application:
    - `name` is indexed and unique
    - `department_id` references a Department row
    """

    __table_args__ = (
        UniqueConstraint(
            "department_id",
            "name",
            name="uq_team_department_name",
        ),
    )

    name: str = Field(index=True)

    department_id: uuid.UUID = Field(foreign_key="department.id")
