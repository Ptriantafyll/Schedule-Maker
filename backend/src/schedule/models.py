"""
Models for the schedule feature
"""

from __future__ import annotations
import uuid
from sqlmodel import Field, Column, JSON
from src.db.schemas import SyncBase


class ScheduleDraft(SyncBase, table=True):
    """Stores generated schedule drafts into SQLite/PostgreSQL"""

    __tablename__ = "schedule_draft"

    department_id: uuid.UUID = Field(foreign_key="department.id", index=True)
    target_month: str = Field(index=True)  # Format YYYY-MM
    source_filename: str
    total_duties: int = Field(default=0)
    solver_status: str = Field(default="OPTIMAL")
    status: str = Field(default="draft")
    # The array of solved shift assignments stored directly as a JSON column:
    assignments: list[dict] = Field(
        default_factory=list, sa_column=Column(JSON)
    )
    # Doctor unavailabilities mapping {doctor_name: [day1, day2, ...]}:
    unavailabilities: dict[str, list[int]] = Field(
        default_factory=dict, sa_column=Column(JSON)
    )

