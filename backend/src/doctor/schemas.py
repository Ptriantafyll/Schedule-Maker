"""Pydantic request/response schemas for the `department` feature.

These DTOs are intentionally separate from the DB `SQLModel` types to
allow evolution of API contracts independently of the persistence schema.
They are configured with `orm_mode = True` so SQLModel/ORM instances can
be returned directly from FastAPI endpoints when convenient.
"""

from __future__ import annotations

from typing import Optional
import uuid
import datetime
from pydantic import BaseModel

# ─── Doctor ───────────────────────────────────────────────


class DoctorBase(BaseModel):
    """
    Shared doctor fields used by create/update/read DTOs

    Fields:
    - `name`: doctor's name
    - `email`: doctor's email (unique)
    - `department_id`: uuid of the department that the doctor belongs to
    - `team_id`: uuid of the team that the doctor belongs to

    """

    name: str
    email: str
    department_id: uuid.UUID
    team_id: uuid.UUID


class DoctorCreate(DoctorBase):
    """
    Schema for doctor creation requests

    Inherits all required fields from `DoctorBasw`. Use this DTO as the request body for POST /departments.
    """


class DoctorUpdate(DoctorBase):
    """
    Schema for partial doctur updates.

    All fields are optional so the clent can PATCH a subset of attributes.
    """

    name: Optional[str] = None
    email: Optional[str] = None
    department_id: Optional[uuid.UUID] = None
    team_id: Optional[uuid.UUID] = None


class DoctorRead(DoctorBase):
    """Schema returned to clients for doctor resources.

    Extends `DoctorBase` with read-only metadata populated by the
    persistence layer (IDs, timestamps, and sync flags).
    `orm_mode = True` allows creating this model from ORM/SQLModel objects
    via `from_orm` or by returning ORM instances directly from FastAPI
    endpoints when `response_model` is set to this class.
    """

    id: uuid.UUID
    name: str
    email: str
    created_at: datetime.datetime
    updated_at: datetime.datetime
    is_deleted: bool = False
    sync_status: bool = False


# ─── DoctorUnavailability ────────────────────────────────

class DoctorUnavailabilityBase(BaseModel):
    """
    Shared doctor unavailability fields used by create/update/read DTOs

    Fields:
    - `doctor_id`: doctor's id
    - `date`: date that the doctor is unavailable

    """

    doctor_id: uuid.UUID
    date: str                        # ISO YYYY-MM-DD


class DoctorUnavailabilityCreate(BaseModel):
    """
    Schema for the creation of doctor unavailabilities

    Inherits all required fields from DoctorUnavailabilityBase.

    Use this DTO as the request body for POST /doctors/{id}/unavailabilities

    doctor_id comes from the path, not the body.
    """
    date: str

    id: uuid.UUID


class DoctorUnavailabilityRead(DoctorUnavailabilityBase):
    """
    Schema returned to clients for doctor unavailability resources.
    """

    created_at: datetime.datetime
    updated_at: datetime.datetime

    class Config:
        orm_mode = True

# ─── DoctorPreAssignment ─────────────────────────────────


class DoctorPreAssignmentBase(BaseModel):
    """
    Shared doctor pre assignment fields used by create/update/read DTOs

    Fields:
    - `doctor_id`: doctor's id
    - `shift_id`: shift's id
    - `date`: date that the doctor is unavailable

    """
    doctor_id: uuid.UUID
    shift_id: uuid.UUID
    date: str


class DoctorPreAssignmentCreate(BaseModel):
    """POST /doctors/{id}/preassignments

    doctor_id comes from the path, not the body.
    """
    shift_id: uuid.UUID
    date: str


class DoctorPreAssignmentRead(DoctorPreAssignmentBase):
    """
    Schema returned to clients for doctor pre assignment resources.
    """
    id: uuid.UUID
    created_at: datetime.datetime
    updated_at: datetime.datetime

    class Config:
        orm_mode = True

# ─── DoctorPosition ──────────────────────────────────────


class DoctorPositionBase(BaseModel):
    """
    Shared doctor pre assignment fields used by create/update/read DTOs

    Fields:
    - `doctor_id`: doctor's id
    - `position_id`: position's id

    """
    doctor_id: uuid.UUID
    position_id: uuid.UUID


class DoctorPositionCreate(BaseModel):
    """POST /doctors/{id}/positions

    doctor_id comes from the path, not the body.
    """
    position_id: uuid.UUID


class DoctorPositionRead(DoctorPositionBase):
    """
    Schema returned to clients for doctor position resources.
    """
    id: uuid.UUID
    created_at: datetime.datetime
    updated_at: datetime.datetime

    class Config:
        orm_mode = True
