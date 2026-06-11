"""Feature-local ORM models for the `department` feature.

This file defines a self-contained `Department` SQLModel table used by the
department feature. It's intentionally implemented here so the feature can
own its model shape independently (useful during refactors or feature
extraction). Keep migrations in sync manually if you change this schema.
"""
from __future__ import annotations

from typing import Optional
import uuid
import datetime

from sqlmodel import SQLModel, Field


class SyncBase(SQLModel):
    """Base fields used for sync/metadata across tables."""
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


__all__ = ["Department", "SyncBase"]
