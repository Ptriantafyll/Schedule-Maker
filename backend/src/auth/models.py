"""
Module auth.models.py

Defines the Invitation SQLModel table representing scoped account invitations.
"""

from __future__ import annotations

from typing import Optional
import uuid
import datetime
from sqlmodel import Field

from src.db.schemas import SyncBase
from src.user.models import UserRole


class Invitation(SyncBase, table=True):
    """
    Represents an invitation to a user
    """
    token_hash: str = Field(index=True, unique=True)
    role: UserRole = Field()
    department_id: uuid.UUID = Field(foreign_key="department.id")
    doctor_id: Optional[uuid.UUID] = Field(
        default=None, foreign_key="doctor.id", nullable=True)
    created_by_user_id: uuid.UUID = Field(foreign_key="user.id")
    expires_at: datetime.datetime = Field()
    used_at: Optional[datetime.datetime] = Field(default=None, nullable=True)
    revoked_at: Optional[datetime.datetime] = Field(
        default=None, nullable=True)

    @property
    def is_expired(self) -> bool:
        """Checks whether an invitation has expired"""
        now = datetime.datetime.now(datetime.timezone.utc)
        if self.expires_at.tzinfo is None:
            now = datetime.datetime.utcnow()
        return now > self.expires_at

    @property
    def is_active(self) -> bool:
        """Checks whether an invitation is active on all fronts"""
        return (
            not self.is_deleted
            and self.used_at is None
            and self.revoked_at is None
            and not self.is_expired
        )
