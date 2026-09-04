"""
Pydantic schemas for users
"""

import uuid
import datetime
from typing import Optional
from pydantic import BaseModel, EmailStr, ConfigDict, field_validator
from src.user.models import UserRole


class UserBase(BaseModel):
    """Base user schema"""
    email: EmailStr
    full_name: str
    role: UserRole = UserRole.DOCTOR
    department_id: Optional[uuid.UUID] = None
    doctor_id: Optional[uuid.UUID] = None


class UserCreate(UserBase):
    """
    Schema for creating a new user
    POST /users/signup
    """
    password: str


class UserLogin(BaseModel):
    """
    Schema for a user logging in
    POST /users/login
    """

    email: EmailStr
    password: str


class UserRead(UserBase):
    """
    GET /users/<id>
    """

    id: uuid.UUID
    is_deleted: bool
    sync_status: bool
    created_at: datetime.datetime
    updated_at: datetime.datetime

    model_config = ConfigDict(from_attributes=True)


class UserAccountBase(BaseModel):
    """
    Base schema for user accounts
    """

    email: EmailStr
    full_name: str
    role: UserRole
    department_id: Optional[uuid.UUID] = None
    doctor_id: Optional[uuid.UUID] = None

    model_config = ConfigDict(extra="forbid")


class UserAccountCreate(UserAccountBase):
    """
    Schema for creating a new user account
    POST /users/signup
    """

    password: str


class UserPersistenceCreate(UserAccountBase):
    """
    Schema for creating a new user account in the persistence layer
    This schema is used internally and is not exposed to the API.
    """

    hashed_password: str

    @field_validator("hashed_password")
    @classmethod
    def validate_hashed_password(cls, value: str) -> str:
        """Validates that the hashed_password is a valid bcrypt hash"""
        if not (value.startswith(("$2b$", "$2a$", "$2y$"))):
            raise ValueError("hashed_password must be a valid bcrypt hash")
        return value
