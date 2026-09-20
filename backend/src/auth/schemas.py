"""
Pydantic schemas for authentication, token handling, and account invitations.
"""

from __future__ import annotations

import datetime
import uuid
from typing import Optional
from pydantic import BaseModel, ConfigDict, EmailStr, field_validator
from src.user.models import UserRole
from src.department.schemas import DepartmentRead
from src.auth.security import validate_password_strength

# ----------------------------------------
# Authentication Tokens
# ----------------------------------------


class Token(BaseModel):
    """JWT bearer access token response schema."""
    access_token: str
    token_type: str = "bearer"
    refresh_token: Optional[str] = None
    csrf_token: Optional[str] = None


class TokenPayload(BaseModel):
    """Decoded JWT payload containing verified standard claims."""
    sub: str  # user.id as string
    exp: datetime.datetime
    iat: datetime.datetime
    jti: str
    iss: str
    aud: str
    token_type: str = "access"


# ----------------------------------------
# Invitations
# ----------------------------------------

class StaffInvitationCreate(BaseModel):
    """
    Request schema used by Department Admins to generate staff invitations.

    Department admins may only issue invitations for DOCTOR or VIEWER roles.
    Department ID is derived securely from the authenticated admin's session.
    """
    role: UserRole = UserRole.DOCTOR
    doctor_id: Optional[uuid.UUID] = None

    model_config = ConfigDict(extra="forbid")

    @field_validator("role")
    @classmethod
    def validate_staff_role(cls, v: UserRole) -> UserRole:
        if v not in (UserRole.DOCTOR, UserRole.VIEWER):
            raise ValueError(
                "Department admins may only invite doctors or viewers.")
        return v


class DepartmentAdminProvisioningCreate(BaseModel):
    """
    Request schema used by Super-Admins to provision a new department
    and generate its initial Department Admin invitation token atomically.
    """
    department_name: str
    department_code: str

    model_config = ConfigDict(extra="forbid")

    @field_validator("department_name", "department_code")
    @classmethod
    def strip_and_validate_non_empty(cls, v: str) -> str:
        cleaned = v.strip()
        if not cleaned:
            raise ValueError("Field cannot be empty or whitespace only.")
        return cleaned


class DepartmentAdminInvitationCreate(BaseModel):
    """
    Request schema used by Super-Admins to generate an admin invitation
    for an existing department.
    """
    department_id: uuid.UUID

    model_config = ConfigDict(extra="forbid")


class InvitationSignupRequest(BaseModel):
    """
    Public registration request schema submitted by an invitee.

    The invitee provides their bearer token and personal profile details.
    Role and department assignments are derived immutably from the invitation record.
    """
    invitation_token: str
    first_name: str
    last_name: str
    email: EmailStr
    password: str

    model_config = ConfigDict(extra="forbid")

    @field_validator("invitation_token", "first_name", "last_name")
    @classmethod
    def strip_and_validate_non_empty(cls, v: str) -> str:
        cleaned = v.strip()
        if not cleaned:
            raise ValueError("Field cannot be empty or whitespace only.")
        return cleaned

    @field_validator("password")
    @classmethod
    def validate_password(cls, v: str) -> str:
        if not v:
            raise ValueError("Password cannot be empty.")
        validate_password_strength(v)
        return v


class InvitationCreatedResponse(BaseModel):
    """
    Response schema returned upon invitation generation.

    This is the ONLY schema that exposes the raw plaintext invitation token.
    Once returned to the creator, the raw token cannot be retrieved again.
    """
    id: uuid.UUID
    role: UserRole
    department_id: uuid.UUID
    doctor_id: Optional[uuid.UUID] = None
    expires_at: datetime.datetime
    raw_token: str


class InvitationRead(BaseModel):
    """
    Administrative read schema for viewing invitation records and audit metadata.

    Security Guarantee: Never exposes token_hash or raw_token.
    """
    id: uuid.UUID
    role: UserRole
    department_id: uuid.UUID
    doctor_id: Optional[uuid.UUID] = None
    created_by_user_id: uuid.UUID
    expires_at: datetime.datetime
    used_at: Optional[datetime.datetime] = None
    revoked_at: Optional[datetime.datetime] = None
    created_at: datetime.datetime
    updated_at: datetime.datetime
    is_deleted: bool = False
    sync_status: bool = False

    model_config = ConfigDict(from_attributes=True)


class DepartmentProvisioningResponse(BaseModel):
    """Response returned upon provisioning a department with its first admin invitation."""
    department: DepartmentRead
    invitation: InvitationCreatedResponse


class RefreshTokenRequest(BaseModel):
    """Refresh token request for non-cookie users."""
    refresh_token: Optional[str] = None
    csrf_token: Optional[str] = None

    model_config = ConfigDict(extra="forbid")

    @field_validator("refresh_token", "csrf_token")
    @classmethod
    def strip_and_validate_not_empty(cls, v: Optional[str]) -> Optional[str]:
        if v is not None:
            cleaned = v.strip()
            if not cleaned:
                raise ValueError("Field cannot be empty or whitespace only.")
            return cleaned
        return v


class RefreshSessionRead(BaseModel):
    """Used for administrative or user dashboard visibility."""

    id: uuid.UUID
    user_id: uuid.UUID
    session_family: uuid.UUID
    expires_at: datetime.datetime
    last_used_at: Optional[datetime.datetime] = None
    revoked_at: Optional[datetime.datetime] = None
    revoked_reason: Optional[str] = None
    replaced_by_session_id: Optional[uuid.UUID] = None
    created_at: datetime.datetime
    updated_at: datetime.datetime
    is_deleted: bool = False
    is_active: bool
    is_expired: bool

    model_config = ConfigDict(from_attributes=True)
