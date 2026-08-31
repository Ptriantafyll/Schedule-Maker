"""
Tests for the auth dependencies
"""

import uuid
import pytest

from fastapi import HTTPException

from src.auth import dependencies
from src.user.models import UserRole
from src.user.models import User as UserModel

EXPECTED_USER_ROLES = {
    UserRole.SUPER_ADMIN,
    UserRole.DEPARTMENT_ADMIN,
    UserRole.DOCTOR,
    UserRole.VIEWER,
}

ROLE_GUARD_RULES = (
    (
        "department-member",
        dependencies.require_department_member,
        (
            UserRole.DEPARTMENT_ADMIN,
            UserRole.DOCTOR,
            UserRole.VIEWER,
        ),
    ),
    (
        "department-admin",
        dependencies.require_department_admin,
        (
            UserRole.DEPARTMENT_ADMIN,
        ),
    ),
    (
        "doctor-or-department-admin",
        dependencies.require_doctor_or_department_admin,
        (
            UserRole.DEPARTMENT_ADMIN,
            UserRole.DOCTOR,
        ),
    ),
    (
        "super-admin",
        dependencies.require_super_admin,
        (
            UserRole.SUPER_ADMIN,
        ),
    ),
)

ALLOWED_ROLE_CASES = [
    pytest.param(
        guard,
        role,
        id=f"{guard_name}-{role.value}"
    )
    for guard_name, guard, allowed_roles in ROLE_GUARD_RULES
    for role in allowed_roles
]

DENIED_ROLE_CASES = [
    pytest.param(
        guard,
        role,
        id=f"{guard_name}-rejects-{role.value}"
    )
    for guard_name, guard, allowed_roles in ROLE_GUARD_RULES
    for role in UserRole
    if role not in allowed_roles
]


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

############################
# Tests
############################


def test_user_role_changes_require_authorization_review():
    """Tests that checks if the roles have been changed (added/removed/modified)"""
    assert set(UserRole) == EXPECTED_USER_ROLES


@pytest.mark.parametrize(
    "guard,role",
    ALLOWED_ROLE_CASES
)
def test_role_guard_allows_expected_roles(guard, role):
    """Tests that a role guard allows the expected roles"""
    user = _make_user(role)

    assert guard(current_user=user) is user


@pytest.mark.parametrize(
    "guard,role",
    DENIED_ROLE_CASES
)
def test_role_guard_rejects_disallowed_roles(guard, role):
    """Tests that a role guard rejects the disallowed roles"""
    user = _make_user(role)

    with pytest.raises(HTTPException) as exc_info:
        guard(current_user=user)

    assert exc_info.value.status_code == 403
    assert exc_info.value.detail == "Insufficient permissions for this operation"
    assert exc_info.value.headers is None
