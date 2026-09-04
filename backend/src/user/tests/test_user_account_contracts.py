"""
Test for User account contracts.
"""

import uuid

import pytest
from pydantic import ValidationError
from sqlmodel import select


from src.user.schemas import (
    UserAccountCreate,
    UserPersistenceCreate,
)
from src.user import services as user_services
from src.user import repository as user_repository
from src.user.models import User as UserModel
from src.user.models import UserRole
from src.auth.security import hash_password, verify_password, decode_access_token
from src.auth import controllers as auth_controllers


######################
# Helpers
######################


def build_account_create_data(
    **overrides: object
) -> dict[str, object]:
    data = {
        "email": "viewer@example.com",
        "full_name": "Test Viewer",
        "password": "test-password",
        "role": UserRole.VIEWER,
        "department_id": uuid.uuid4(),
        "doctor_id": None,
    }
    data.update(overrides)
    return data

######################
# Tests
######################


def test_user_account_create_requires_explicit_role():
    """Tests that account creation cannot silently default to a role."""
    account_data = build_account_create_data()
    account_data.pop("role")

    with pytest.raises(ValidationError) as exc_info:
        UserAccountCreate(**account_data)

    role_error = next(
        (
            error
            for error in exc_info.value.errors()
            if error["loc"] == ("role",)
        ),
        None,
    )

    assert role_error is not None
    assert role_error["type"] == "missing"


def test_user_persistence_create_requires_explicit_role():
    """Tests that account persistence cannot silently default to a role."""
    account_data = build_account_create_data()
    account_data["hashed_password"] = hash_password(account_data["password"])
    account_data.pop("password")

    account_data.pop("role")

    with pytest.raises(ValidationError) as exc_info:
        UserPersistenceCreate(**account_data)

    role_error = next(
        (
            error
            for error in exc_info.value.errors()
            if error["loc"] == ("role",)
        ),
        None,
    )

    assert role_error is not None
    assert role_error["type"] == "missing"


def test_user_account_create_uses_plaintext_password_field():
    """Tests that UserAccountCreate accepts password and has no hashed_password field."""
    account_data = build_account_create_data()
    created_data = UserAccountCreate(**account_data)

    assert created_data.password == account_data["password"]
    assert "password" in created_data.model_dump()
    assert "hashed_password" not in created_data.model_dump()


def test_user_persistence_create_uses_hashed_password_field():
    """Tests UserPersistenceCreate accepts hashed_password and has no password field."""
    account_data = build_account_create_data()

    hashed_password = hash_password(account_data["password"])
    account_data["hashed_password"] = hashed_password
    account_data.pop("password")

    created_data = UserPersistenceCreate(**account_data)
    assert created_data.hashed_password == hashed_password
    assert "hashed_password" in created_data.model_dump()
    assert "password" not in created_data.model_dump()


def test_user_persistence_create_rejects_plaintext_password():
    """Tests that plaintext cannot be used as a persisted password hash."""
    account_data = build_account_create_data()
    plaintext_password = account_data.pop("password")
    account_data["hashed_password"] = plaintext_password

    with pytest.raises(ValidationError) as exc_info:
        UserPersistenceCreate(**account_data)

    hashed_password_error = next(
        (
            error
            for error in exc_info.value.errors()
            if error["loc"] == ("hashed_password",)
        ),
        None,
    )

    assert hashed_password_error is not None
    assert hashed_password_error["type"] == "value_error"
    assert "bcrypt" in hashed_password_error["msg"].lower()


def test_user_account_create_rejects_hashed_password_field():
    """Tests that plaintext input cannot accept a password hash field."""
    account_data = build_account_create_data()
    account_data["hashed_password"] = hash_password("test-password")

    with pytest.raises(ValidationError) as exc_info:
        UserAccountCreate(**account_data)

    hashed_password_error = next(
        (
            error
            for error in exc_info.value.errors()
            if error["loc"] == ("hashed_password",)
        ),
        None,
    )

    assert hashed_password_error is not None
    assert hashed_password_error["type"] == "extra_forbidden"


def test_user_persistence_create_rejects_password_field():
    """Tests that persistence input cannot accept a plaintext field."""
    persistence_data = build_account_create_data()
    persistence_data["hashed_password"] = hash_password("test-password")
    # Keep password intentionally so it is treated as the forbidden extra field.

    with pytest.raises(ValidationError) as exc_info:
        UserPersistenceCreate(**persistence_data)

    password_error = next(
        (
            error
            for error in exc_info.value.errors()
            if error["loc"] == ("password",)
        ),
        None,
    )

    assert password_error is not None
    assert password_error["type"] == "extra_forbidden"


def test_create_user_account_does_not_mutate_input(session, department_factory):
    """Tests that account creation does not mutate its plaintext input."""
    department = department_factory(name="Test Department", code="TEST")
    plaintext_password = "test-password"
    account_input = UserAccountCreate(
        **build_account_create_data(
            department_id=department.id,
            password=plaintext_password,
        )
    )
    original_input = account_input.model_dump()

    created_user = user_services.create_user_account(
        session=session,
        account_data=account_input,
    )

    assert account_input.model_dump() == original_input
    assert account_input.password == plaintext_password

    persisted_user = session.get(UserModel, created_user.id)
    assert persisted_user is not None
    assert persisted_user.hashed_password != plaintext_password
    assert verify_password(
        plaintext_password,
        persisted_user.hashed_password,
    )


def test_stage_user_account_can_be_rolled_back(session, department_factory):
    """Tests stage_user_account can be rolled back without mutating the input."""
    department = department_factory(name="Test Department", code="TEST")
    plaintext_password = "test-password"
    account_input = UserAccountCreate(
        **build_account_create_data(
            department_id=department.id,
            password=plaintext_password,
        )
    )
    original_input = account_input.model_dump()

    staged_user = user_services.stage_user_account(
        session=session,
        account_data=account_input,
    )

    assert account_input.model_dump() == original_input
    assert account_input.password == plaintext_password
    assert staged_user.hashed_password != plaintext_password
    assert verify_password(
        plaintext_password,
        staged_user.hashed_password,
    )

    staged_user_id = staged_user.id
    session.rollback()

    persisted_user = session.get(UserModel, staged_user_id)
    assert persisted_user is None


def test_create_user_account_commits_user(session, department_factory):
    """Tests that committed account creation survives a later rollback."""
    department = department_factory(
        name="Committed Account Department",
        code="COMMIT",
    )
    account_input = UserAccountCreate(
        **build_account_create_data(
            department_id=department.id,
        )
    )

    created_user = user_services.create_user_account(
        session=session,
        account_data=account_input,
    )
    created_user_id = created_user.id

    session.rollback()

    persisted_user = session.get(UserModel, created_user_id)
    assert persisted_user is not None
    assert persisted_user.id == created_user_id


def test_create_user_account_translates_database_email_conflict(
    session,
    department_factory,
    monkeypatch,
):
    """Tests that a database email conflict becomes a domain error."""
    department = department_factory(
        name="Database Failure Department",
        code="FAILURE",
    )
    account_input = UserAccountCreate(
        **build_account_create_data(
            email="database-conflict@example.com",
            department_id=department.id,
        )
    )
    existing_user = UserModel(
        email=account_input.email,
        full_name="Existing User",
        hashed_password=hash_password("existing-password"),
        role=account_input.role,
        department_id=account_input.department_id,
        doctor_id=account_input.doctor_id,
    )
    session.add(existing_user)
    session.commit()
    existing_user_id = existing_user.id
    original_full_name = existing_user.full_name
    original_password_hash = existing_user.hashed_password

    def stage_duplicate_user(*, session, account_data):
        duplicate_user = UserModel(
            email=account_data.email,
            full_name=account_data.full_name,
            hashed_password=hash_password(account_data.password),
            role=account_data.role,
            department_id=account_data.department_id,
            doctor_id=account_data.doctor_id,
        )
        session.add(duplicate_user)
        session.flush()
        return duplicate_user

    monkeypatch.setattr(
        user_services,
        "stage_user_account",
        stage_duplicate_user,
    )

    with pytest.raises(
        user_services.UserEmailAlreadyExistsError
    ) as exc_info:
        user_services.create_user_account(
            session=session,
            account_data=account_input,
        )

    assert str(exc_info.value) == "A user with this email already exists."

    stored_users = session.exec(select(UserModel)).all()
    assert len(stored_users) == 1
    assert stored_users[0].id == existing_user_id
    assert stored_users[0].full_name == original_full_name
    assert stored_users[0].hashed_password == original_password_hash


def test_create_user_account_stores_canonical_email(
    session,
    department_factory,
):
    """Tests that account creation stores the canonical email."""
    department = department_factory()
    entered_email = " CaNonical@example.COM"
    expected_email = "canonical@example.com"

    account_input = UserAccountCreate(
        **build_account_create_data(
            email=entered_email,
            department_id=department.id,
        )
    )
    created_user = user_services.create_user_account(
        session=session,
        account_data=account_input,
    )
    assert created_user.email == expected_email
    persisted_user = session.get(UserModel, created_user.id)
    assert persisted_user is not None
    assert persisted_user.email == expected_email


def test_get_user_by_email_normalizes_lookup(
    session,
    department_factory,
):
    """Tests that get_user_by_email normalizes the email for lookup."""
    department = department_factory()
    entered_email = " CaNonical@example.COM"
    expected_email = "canonical@example.com"

    account_input = UserAccountCreate(
        **build_account_create_data(
            email=expected_email,
            department_id=department.id,
        )
    )
    user_services.create_user_account(
        session=session,
        account_data=account_input,
    )
    retrieved_user = user_repository.get_user_by_email(
        session=session,
        user_email=entered_email,
    )
    assert retrieved_user is not None
    assert retrieved_user.email == expected_email


@pytest.mark.parametrize(
    "entered_email",
    [
        "CaNonical@example.COM",
        " canonical@example.com ",
    ],
)
def test_login_accepts_non_canonical_email(
    session,
    department_factory,
    entered_email,
):
    """Tests that login accepts email case variants."""
    department = department_factory()
    password = "securepassword"
    canonical_email = "canonical@example.com"

    account_input = UserAccountCreate(
        **build_account_create_data(
            email=canonical_email,
            department_id=department.id,
            password=password,
        )
    )
    created_user = user_services.create_user_account(
        session=session,
        account_data=account_input,
    )

    logged_in_user_token = auth_controllers.login_controller(
        email=entered_email,
        password=password,
        session=session,
    )

    assert logged_in_user_token is not None
    decoded_token = decode_access_token(logged_in_user_token.access_token)
    assert decoded_token["sub"] == str(created_user.id)
    assert decoded_token["token_type"] == "access"
    assert decoded_token["iss"] == "schedule-maker-api"
    assert decoded_token["aud"] == "schedule-maker-clients"


def test_create_user_account_rejects_case_variant_duplicate(
    session,
    department_factory,
):
    """Tests that account creation rejects case-variant duplicates."""
    department_a = department_factory()
    department_b = department_factory()
    entered_email = "CaNonical@example.COM"

    account_input = UserAccountCreate(
        **build_account_create_data(
            email=entered_email,
            department_id=department_a.id,
        )
    )

    original_user = user_services.create_user_account(
        session=session,
        account_data=account_input,
    )
    original_full_name = original_user.full_name
    original_password_hash = original_user.hashed_password

    entered_email_variant = " canonical@example.com "
    duplicate_account_input = UserAccountCreate(
        **build_account_create_data(
            email=entered_email_variant,
            department_id=department_b.id,
            full_name="Duplicate User",
            password="different-password",
        )
    )

    with pytest.raises(user_services.UserEmailAlreadyExistsError) as exc_info:
        user_services.create_user_account(
            session=session,
            account_data=duplicate_account_input,
        )

    assert str(exc_info.value) == "A user with this email already exists."

    stored_users = session.exec(select(UserModel)).all()
    assert len(stored_users) == 1
    assert stored_users[0].id == original_user.id
    assert stored_users[0].email == "canonical@example.com"
    assert stored_users[0].full_name == original_full_name
    assert stored_users[0].hashed_password == original_password_hash


def test_soft_deleted_user_email_remains_reserved(
    session,
    department_factory,
):
    """Tests that a soft-deleted user email remains reserved."""
    department = department_factory()
    entered_email = "reserved@example.com"

    account_input = UserAccountCreate(
        **build_account_create_data(
            email=entered_email,
            department_id=department.id,
        )
    )

    original_user = user_services.create_user_account(
        session=session,
        account_data=account_input,
    )

    # Soft delete the user
    stored_user = session.get(UserModel, original_user.id)
    assert stored_user is not None
    stored_user.is_deleted = True
    session.add(stored_user)
    session.commit()

    # Attempt to create a new account with the same email
    duplicate_account_input = UserAccountCreate(
        **build_account_create_data(
            email=entered_email,
            department_id=department.id,
        )
    )

    with pytest.raises(user_services.UserEmailAlreadyExistsError) as exc_info:
        user_services.create_user_account(
            session=session,
            account_data=duplicate_account_input,
        )

    assert str(exc_info.value) == "A user with this email already exists."
    stored_users = session.exec(select(UserModel)).all()

    assert len(stored_users) == 1
    assert stored_users[0].id == original_user.id
    assert stored_users[0].email == "reserved@example.com"
    assert stored_users[0].is_deleted is True
