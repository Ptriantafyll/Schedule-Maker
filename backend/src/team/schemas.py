"""Pydantic request/response schemas for the `team` feature.

These DTOs are intentionally separate from the DB `SQLModel` types to
allow evolution of API contracts independently of the persistence schema.
They are configured with `from_attributes = True` so SQLModel/ORM instances can
be returned directly from FastAPI endpoints when convenient.
"""
from __future__ import annotations

from typing import Optional
import uuid
import datetime

from pydantic import BaseModel, ConfigDict


class TeamBase(BaseModel):
    """Shared team fields used by create/update/read DTOs.

    Fields:
    - `name`: human-readable team name (unique).
    - `department_id`: UUID referencing the parent department.
    """

    name: str
    department_id: uuid.UUID


class TeamCreate(TeamBase):
    """Schema for team creation requests.

    Inherits all required fields from `TeamBase`. Use this DTO as the
    request body for POST /teams.
    """


class TeamUpdate(BaseModel):
    """Schema for partial team updates.

    All fields are optional so the client can PATCH a subset of attributes.
    """

    name: Optional[str] = None
    department_id: Optional[uuid.UUID] = None


class TeamRead(TeamBase):
    """Schema returned to clients for team resources.

    Extends `TeamBase` with read-only metadata populated by the
    persistence layer (IDs, timestamps, and sync flags).
    `from_attributes = True` allows creating this model from ORM/SQLModel objects
    via `model_validate` or by returning ORM instances directly from FastAPI
    endpoints when `response_model` is set to this class.
    """

    id: uuid.UUID
    created_at: datetime.datetime
    updated_at: datetime.datetime
    is_deleted: bool = False
    sync_status: bool = False

    model_config = ConfigDict(from_attributes=True)


__all__ = [
    "TeamBase",
    "TeamCreate",
    "TeamUpdate",
    "TeamRead"
]
