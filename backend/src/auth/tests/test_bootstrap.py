"""
Tests for the bootstrap service
"""
import pytest
from sqlmodel import select

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


def test_bootstrap_stores_canonical_email(session):
    """Tests that bootstrap stores the canonical login email."""
    entered_email = " Mixed.Admin@Example.COM "
    expected_email = "mixed.admin@example.com"

    created_user = bootstrap.create_super_admin(
        session=session,
        email=entered_email,
        full_name="Mixed Case Admin",
        password=PLAIN_SUPER_ADMIN_PASSWORD,
    )

    assert created_user.email == expected_email

    persisted_user = session.get(UserModel, created_user.id)
    assert persisted_user is not None
    assert persisted_user.email == expected_email


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


def test_bootstrap_rejects_case_variant_duplicate_email(
    session,
    existing_super_admin,
):
    """Tests that bootstrap treats email case variants as one identity."""
    original_full_name = existing_super_admin.full_name
    original_password_hash = existing_super_admin.hashed_password

    with pytest.raises(bootstrap.SuperAdminAlreadyExistsError):
        bootstrap.create_super_admin(
            session=session,
            email=f" {SUPER_ADMIN_EMAIL.upper()} ",
            full_name="Duplicate Case Variant",
            password=PLAIN_SUPER_ADMIN_PASSWORD,
        )

    stored_users = session.exec(select(UserModel)).all()

    assert len(stored_users) == 1
    assert stored_users[0].id == existing_super_admin.id
    assert stored_users[0].email == SUPER_ADMIN_EMAIL
    assert stored_users[0].full_name == original_full_name
    assert stored_users[0].hashed_password == original_password_hash


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
