"""
ORM models for user
"""

import enum
import uuid
from typing import Optional
from sqlmodel import Field
from src.db.schemas import SyncBase


class UserRole(str, enum.Enum):
    """Role class for the users"""
    SUPER_ADMIN = "super_admin"
    DEPARTMENT_ADMIN = "department_admin"
    DOCTOR = "doctor"
    VIEWER = "viewer"


class User(SyncBase, table=True):
    """Represents an app user/account stored in the database"""
    email: str = Field(index=True, unique=True)
    hashed_password: str
    full_name: str
    role: UserRole = Field(default=UserRole.DOCTOR)

    department_id: Optional[uuid.UUID] = Field(
        default=None, foreign_key="department.id")
    doctor_id: Optional[uuid.UUID] = Field(
        default=None, foreign_key="doctor.id")
