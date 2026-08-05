"""Pydantic request/response schemas for the `department` feature.

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


class DepartmentBase(BaseModel):
    """Shared department fields used by create/update/read DTOs.

    Fields:
    - `name`: human-readable department name (unique).
    - `code`: short department code/identifier.
    - `backup_department_id`: optional UUID referencing a backup department.
    """

    name: str
    code: str
    backup_department_id: Optional[uuid.UUID] = None


class DepartmentCreate(DepartmentBase):
    """Schema for department creation requests.

    Inherits all required fields from `DepartmentBase`. Use this DTO as the
    request body for POST /departments.
    """


class DepartmentUpdate(BaseModel):
    """Schema for partial department updates.

    All fields are optional so the client can PATCH a subset of attributes.
    """

    name: Optional[str] = None
    code: Optional[str] = None
    backup_department_id: Optional[uuid.UUID] = None


class DepartmentRead(DepartmentBase):
    """Schema returned to clients for department resources.

    Extends `DepartmentBase` with read-only metadata populated by the
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
    "DepartmentBase",
    "DepartmentCreate",
    "DepartmentUpdate",
    "DepartmentRead",
]
