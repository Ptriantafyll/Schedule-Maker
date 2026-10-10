
"""
Pydantic schemas for draft schedules.
"""

from __future__ import annotations
from typing import Optional
import uuid
import datetime
from pydantic import BaseModel, ConfigDict


class ScheduleAssignmentItem(BaseModel):
    """Schema for individual shift assignment within a schedule draft."""

    date: str  # "YYYY-MM-DD"
    day_name: Optional[str] = None  # "Monday", "Sunday"
    doctor_name: str
    doctor_email: Optional[str] = None
    position: str  # "ER", "ICU"
    shift: str  # "Night", "Morning"

    model_config = ConfigDict(from_attributes=True)


class ScheduleDraftRead(BaseModel):
    """Schema returned to API clients for a schedule draft."""

    id: uuid.UUID
    department_id: uuid.UUID
    target_month: str
    source_filename: str
    total_duties: int
    solver_status: str
    status: str
    assignments: list[ScheduleAssignmentItem] = []
    unavailabilities: dict[str, list[int]] = {}
    created_at: datetime.datetime
    updated_at: datetime.datetime

    model_config = ConfigDict(from_attributes=True)


class ScheduleSummaryRead(BaseModel):
    """Lightweight schema for listing schedule summaries without heavy assignments."""

    id: uuid.UUID
    department_id: uuid.UUID
    target_month: str
    source_filename: str
    total_duties: int
    solver_status: str
    status: str
    created_at: datetime.datetime
    updated_at: datetime.datetime

    model_config = ConfigDict(from_attributes=True)


class TargetMonthResponse(BaseModel):
    """Response model for target month resolution."""
    next_target_month: str
    last_published_month: Optional[str] = None
