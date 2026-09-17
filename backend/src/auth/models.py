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
        if getattr(self.expires_at, "tzinfo", None) is None:
            now = now.replace(tzinfo=None)
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


class RefreshSession(SyncBase, table=True):
    """
    Represents the refresh sessions in the db.
    """
    user_id: uuid.UUID = Field(foreign_key="user.id", index=True)
    refresh_token_hash: str = Field(index=True, unique=True)
    session_family: uuid.UUID = Field(index=True)
    csrf_token_hash: Optional[str] = Field(default=None, nullable=True)
    expires_at: datetime.datetime = Field()
    last_used_at: Optional[datetime.datetime] = Field(
        default=None, nullable=True)
    revoked_at: Optional[datetime.datetime] = Field(
        default=None, nullable=True)
    revoked_reason: Optional[str] = Field(default=None, nullable=True)
    replaced_by_session_id: Optional[uuid.UUID] = Field(
        default=None, foreign_key="refreshsession.id", nullable=True)

    @property
    def is_expired(self) -> bool:
        """Checks whether an invitation has expired"""
        now = datetime.datetime.now(datetime.timezone.utc)
        if getattr(self.expires_at, "tzinfo", None) is None:
            now = now.replace(tzinfo=None)
        return now > self.expires_at

    @property
    def is_active(self) -> bool:
        """Checks whether an invitation is active on all fronts"""
        return (
            not self.is_deleted
            and self.replaced_by_session_id is None
            and self.revoked_at is None
            and not self.is_expired
        )
