"""
Tests for the bootstrap service
"""
import pytest
from src.auth import bootstrap
from src.auth.security import verify_password
from src.user import repository as user_repository
from src.user.models import UserRole
from src.user.models import User as UserModel

PLAIN_SUPER_ADMIN_PASSWORD = "secure-test-password"
SUPER_ADMIN_EMAIL = "superadmin@test.com"


@pytest.fixture
def existing_super_admin(session):
    """Creates a reusable super admin for tests"""
    return bootstrap.create_super_admin(
        session=session,
        email=SUPER_ADMIN_EMAIL,
        full_name="Test Super Admin",
        password=PLAIN_SUPER_ADMIN_PASSWORD,
    )


def test_create_super_admin_creates_valid_account(session):
    """Tests creating a super admin correctly """

    created_user = bootstrap.create_super_admin(
        session=session,
        email=SUPER_ADMIN_EMAIL,
        full_name="Test Super Admin",
        password=PLAIN_SUPER_ADMIN_PASSWORD,
    )

    assert created_user.role == UserRole.SUPER_ADMIN
    assert created_user.department_id is None
    assert created_user.doctor_id is None
    assert created_user.hashed_password != PLAIN_SUPER_ADMIN_PASSWORD
    assert verify_password(
        PLAIN_SUPER_ADMIN_PASSWORD,
        created_user.hashed_password,
    )

    persisted_user = user_repository.get_user_by_email(
        session, SUPER_ADMIN_EMAIL)

    assert persisted_user is not None
    assert persisted_user.id == created_user.id


def test_create_super_admin_rejects_duplicate_email(session, existing_super_admin):
    """Tests creating a super admin with an email that already exists"""
    with pytest.raises(bootstrap.SuperAdminAlreadyExistsError):
        bootstrap.create_super_admin(
            session=session,
            email=SUPER_ADMIN_EMAIL,
            full_name="Test Super Admin",
            password=PLAIN_SUPER_ADMIN_PASSWORD,
        )

    retrieved_user = user_repository.get_user_by_email(
        session,
        SUPER_ADMIN_EMAIL,
    )

    assert retrieved_user is not None
    assert str(retrieved_user.id) == str(existing_super_admin.id)


def test_create_super_admin_rolls_back_database_duplicate(session, existing_super_admin, monkeypatch):
    """Tests an existing user trying to be created in the db"""
    monkeypatch.setattr(
        bootstrap.user_repository,
        "get_user_by_email",
        lambda *args, **kwargs: None,
    )

    with pytest.raises(bootstrap.SuperAdminAlreadyExistsError):
        bootstrap.create_super_admin(
            session=session,
            email=SUPER_ADMIN_EMAIL,
            full_name="Test Super Admin",
            password=PLAIN_SUPER_ADMIN_PASSWORD,
        )

    retrieved_user = session.get(UserModel, existing_super_admin.id)

    assert retrieved_user is not None
    assert retrieved_user.id == existing_super_admin.id 
