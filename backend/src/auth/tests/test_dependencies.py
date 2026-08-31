"""
Tests for the auth dependencies
"""

import uuid
import pytest

from fastapi import HTTPException

from src.auth import dependencies
from src.user.models import UserRole
from src.user.models import User as UserModel

############################
# Helpers
############################


def _make_user(role: UserRole) -> UserModel:
    """Create an unpersisted user for role-guard tests"""
    return UserModel(
        email=f"{role.value}@test.com",
        full_name="Test User",
        hashed_password="not-used-by-role-guard-tests",
        role=role,
        department_id=(
            None
            if role == UserRole.SUPER_ADMIN
            else uuid.uuid4()
        ),
        doctor_id=(
            uuid.uuid4()
            if role == UserRole.DOCTOR
            else None
        ),
    )


@pytest.mark.parametrize(
    "role",
    [
        UserRole.DEPARTMENT_ADMIN,
        UserRole.DOCTOR,
        UserRole.VIEWER,
    ]
)
def test_department_member_guard_allows_department_roles(role):
    """Tests requiring department member for an operation"""
    user = _make_user(role)

    returned_user = dependencies.require_department_member(
        current_user=user
    )

    assert returned_user is user


def test_department_member_guard_rejects_super_admin():
    """Tests that the department member guard rejects super admin"""
    user = _make_user(UserRole.SUPER_ADMIN)

    with pytest.raises(HTTPException) as exc_info:
        dependencies.require_department_member(
            current_user=user
        )

    assert exc_info.value.status_code == 403
    assert exc_info.value.detail == "Insufficient permissions for this operation"
    assert exc_info.value.headers is None


def test_department_admin_guard_allows_department_admin():
    """Tests that the department admin guard allows department admin"""
    user = _make_user(UserRole.DEPARTMENT_ADMIN)

    returned_user = dependencies.require_department_admin(
        current_user=user
    )

    assert returned_user is user


@pytest.mark.parametrize(
    "role",
    [
        UserRole.SUPER_ADMIN,
        UserRole.DOCTOR,
        UserRole.VIEWER,
    ]
)
def test_department_admin_guard_rejects_other_roles(role):
    """Tests that the department admin guard rejects other roles"""
    user = _make_user(role)

    with pytest.raises(HTTPException) as exc_info:
        dependencies.require_department_admin(
            current_user=user
        )

    assert exc_info.value.status_code == 403
    assert exc_info.value.detail == "Insufficient permissions for this operation"
    assert exc_info.value.headers is None


@pytest.mark.parametrize(
    "role",
    [
        UserRole.DEPARTMENT_ADMIN,
        UserRole.DOCTOR,
    ]
)
def test_doctor_or_department_guard_allows_department_admin(role):
    """Tests that the doctor or department admin guard allows department admin and doctor roles"""
    user = _make_user(role)

    returned_user = dependencies.require_doctor_or_department_admin(
        current_user=user
    )

    assert returned_user is user


@pytest.mark.parametrize(
    "role",
    [
        UserRole.SUPER_ADMIN,
        UserRole.VIEWER,
    ]
)
def test_doctor_or_department_admin_guard_rejects_other_roles(role):
    """Tests that the doctor or department admin guard rejects other roles"""
    user = _make_user(role)

    with pytest.raises(HTTPException) as exc_info:
        dependencies.require_doctor_or_department_admin(
            current_user=user
        )

    assert exc_info.value.status_code == 403
    assert exc_info.value.detail == "Insufficient permissions for this operation"
    assert exc_info.value.headers is None
