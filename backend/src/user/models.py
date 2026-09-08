"""
ORM models for user
"""

import enum
import uuid
from typing import Optional
from sqlmodel import Field
from sqlalchemy import CheckConstraint
from src.db.schemas import SyncBase


class UserRole(str, enum.Enum):
    """Role class for the users"""
    SUPER_ADMIN = "super_admin"
    DEPARTMENT_ADMIN = "department_admin"
    DOCTOR = "doctor"
    VIEWER = "viewer"


class User(SyncBase, table=True):
    """Represents an app user/account stored in the database"""

    __table_args__ = (
        CheckConstraint(
            "("
            f"(role = '{UserRole.SUPER_ADMIN.name}' "
            "AND department_id IS NULL "
            "AND doctor_id IS NULL) OR "
            f"(role = '{UserRole.DEPARTMENT_ADMIN.name}' "
            "AND department_id IS NOT NULL) OR "
            f"(role = '{UserRole.DOCTOR.name}' "
            "AND department_id IS NOT NULL "
            "AND doctor_id IS NOT NULL) OR "
            f"(role = '{UserRole.VIEWER.name}' "
            "AND department_id IS NOT NULL "
            "AND doctor_id IS NULL)"
            ")",
            name="ck_user_role_shape",
        ),
    )

    email: str = Field(index=True, unique=True)
    hashed_password: str
    full_name: str
    role: UserRole = Field()

    department_id: Optional[uuid.UUID] = Field(
        default=None, foreign_key="department.id")
    doctor_id: Optional[uuid.UUID] = Field(
        default=None, foreign_key="doctor.id")
