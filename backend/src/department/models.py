"""Feature-local ORM models for the `department` feature.

This file defines a self-contained `Department` SQLModel table used by the
department feature. It's intentionally implemented here so the feature can
own its model shape independently (useful during refactors or feature
extraction). Keep migrations in sync manually if you change this schema.
"""
from __future__ import annotations

from typing import Optional
import uuid

from sqlmodel import Field
from src.db.schemas import SyncBase


class Department(SyncBase, table=True):
    """Represents a hospital department.

    Fields mirror the DB-level needs used by the application:
    - `name` is indexed and unique
    - `code` is an arbitrary short identifier
    - `backup_department_id` can reference another Department row
    """
    name: str = Field(index=True, unique=True)
    code: str

    backup_department_id: Optional[uuid.UUID] = Field(
        default=None, foreign_key="department.id"
    )


__all__ = ["Department"]
