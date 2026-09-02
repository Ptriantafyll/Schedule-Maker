"""Feature-local ORM models for the `position` feature.

This file defines a self-contained `Position` SQLModel table used by the
position feature. It's intentionally implemented here so the feature can
own its model shape independently (useful during refactors or feature
extraction). Keep migrations in sync manually if you change this schema.
"""

from __future__ import annotations
import uuid

from sqlmodel import Field, Column, JSON
from sqlalchemy import UniqueConstraint
from src.db.schemas import SyncBase


class Position(SyncBase, table=True):
    """Represents a position that needs to be staffed, such as "ER" or "ICU", stored in the database."""
    __table_args__ = (
        UniqueConstraint(
            "department_id",
            "name",
            name="uq_position_department_name",
        ),
    )

    name: str = Field(index=True)
    department_id: uuid.UUID = Field(foreign_key="department.id")
    duty_days: list[int] = Field(default_factory=list, sa_column=Column(JSON))
