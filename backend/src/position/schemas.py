"""Pydantic request/response schemas for the `position` feature.

These DTOs are intentionally separate from the DB `SQLModel` types to
allow evolution of API contracts independently of the persistence schema.
They are configured with `from_attributes = True` so SQLModel/ORM instances can
be returned directly from FastAPI endpoints when convenient.
"""
from __future__ import annotations

from typing import Optional
import uuid
import datetime

from pydantic import BaseModel, ConfigDict, field_validator


def validate_duty_days_list(duty_days: Optional[list[int]]) -> Optional[list[int]]:
    """Validates that duty days are integers within 0..6 (Mon-Sun), non-empty, and returns them sorted."""
    if duty_days is None:
        return None

    if len(duty_days) == 0:
        raise ValueError("Duty days cannot be empty.")

    for day in duty_days:
        if not isinstance(day, int) or day < 0 or day > 6:
            raise ValueError(
                f"Invalid duty day: {day}. Must be an integer between 0 (Monday) and 6 (Sunday)."
            )

    return sorted(list(set(duty_days)))


class PositionBase(BaseModel):
    """Shared position fields used by create/update/read DTOs.

    Fields:
    - `name`: human-readable position name (unique).
    - `department_id`: UUID referencing the parent department.
    """

    name: str
    department_id: uuid.UUID
    duty_days: list[int]


class PositionCreate(BaseModel):
    """Schema for position creation requests.

    Inherits all required fields from `PositionBase`. Use this DTO as the
    request body for POST /positions.
    """

    name: str
    duty_days: list[int]

    model_config = ConfigDict(extra="forbid")

    @field_validator("duty_days")
    @classmethod
    def validate_duty_days(cls, v: Optional[list[int]]) -> Optional[list[int]]:
        return validate_duty_days_list(v)


class PositionUpdate(BaseModel):
    """Schema for partial position updates.

    All fields are optional so the client can PATCH a subset of attributes.
    """

    name: Optional[str] = None
    department_id: Optional[uuid.UUID] = None
    duty_days: Optional[list[int]] = None

    @field_validator("duty_days")
    @classmethod
    def validate_duty_days(cls, v: Optional[list[int]]) -> Optional[list[int]]:
        return validate_duty_days_list(v)


class PositionRead(PositionBase):
    """Schema returned to clients for position resources.

    Extends `PositionBase` with read-only metadata populated by the
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
    "PositionBase",
    "PositionCreate",
    "PositionUpdate",
    "PositionRead"
]
