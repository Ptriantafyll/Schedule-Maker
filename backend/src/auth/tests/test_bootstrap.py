"""
Tests for the bootstrap script
"""

from src.auth import bootstrap
from src.auth.security import verify_password
from src.user import repository as user_repository
from src.user.models import UserRole


def test_create_super_admin_creates_valid_account(session):
    """Tests creating a super admin correctly """
    plain_password = "secure-test-password"

    created_user = bootstrap.create_super_admin(
        session=session,
        email="superadmin@test.com",
        full_name="Test Super Admin",
        password=plain_password,
    )

    assert created_user.role == UserRole.SUPER_ADMIN
    assert created_user.department_id is None
    assert created_user.doctor_id is None
    assert created_user.hashed_password != plain_password
    assert verify_password(
        plain_password,
        created_user.hashed_password,
    )

    persisted_user = user_repository.get_user_by_email(
        session, "superadmin@test.com")

    assert persisted_user is not None
    assert persisted_user.id == created_user.id
