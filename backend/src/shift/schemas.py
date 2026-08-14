"""Pydantic request/response schemas for the `shift` feature.

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


class ShiftBase(BaseModel):
    """Shared shift fields used by create/update/read DTOs.

    Fields:
    - `name`: human-readable shift name (unique).
    - `department_id`: UUID referencing the parent department.
    """

    name: str
    doctors_per_shift: int = 1
    grants_day_off: bool = False
    position_id: uuid.UUID


class ShiftCreate(ShiftBase):
    """Schema for shift creation requests.

    Inherits all required fields from `ShiftBase`. Use this DTO as the
    request body for POST /shifts.
    """


class ShiftUpdate(BaseModel):
    """Schema for partial shift updates.

    All fields are optional so the client can PATCH a subset of attributes.
    """

    name: Optional[str] = None
    doctors_per_shift: Optional[str] = None
    grants_day_odd: Optional[bool] = None
    position_id: Optional[uuid.UUID] = None


class ShiftRead(ShiftBase):
    """Schema returned to clients for shift resources.

    Extends `ShiftBase` with read-only metadata populated by the
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

#############################################
# ShiftAssigment
#############################################


class ShiftAssignmentBase(BaseModel):
    """
    Shared shift assignment fields used by create/update/read DTOs.
    """

    doctor_id: uuid.UUID
    shift_id: uuid.UUID
    date: datetime.date


class ShiftAssignmentCreate(BaseModel):
    """POST shifts/{id}/assignments

    shift_if comes from the path, not the body.
    """
    doctor_id: uuid.UUID
    date: datetime.date


class ShiftAssignmentUpdate(BaseModel):
    """Schema for partial shift updates.

    All fields are optional so the client can PATCH a subset of attributes.
    """

    doctor_id: uuid.UUID
    shift_id: uuid.UUID
    date: datetime.date


class ShiftAssignmentRead(ShiftAssignmentBase):
    """Schema returned to clients for shift assignment resources.

    Extends `ShiftAssignmentBase` with read-only metadata populated by the
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
    "ShiftBase",
    "ShiftCreate",
    "ShiftUpdate",
    "ShiftRead",
    "ShiftAssignmentBase",
    "ShiftAssignmentCreate",
    "ShiftAssignmentUpdate",
    "ShiftAssignmentRead"
]
