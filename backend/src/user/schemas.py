"""
Pydantic schemas for users
"""

import uuid
import datetime
from typing import Optional
from pydantic import BaseModel, EmailStr, ConfigDict
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
