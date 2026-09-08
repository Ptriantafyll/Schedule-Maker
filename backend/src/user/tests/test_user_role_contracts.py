"""
Tests for User role and relationship contracts.
"""

import uuid

import pytest
from sqlalchemy.exc import IntegrityError

from src.team import repository as team_repository
from src.doctor import repository as doctor_repository
from src.user.models import User as UserModel
from src.user.models import UserRole
from src.user.schemas import UserAccountCreate
from src.user import services as user_services
from src.user.services import InvalidUserAccountRelationshipError
from src.user import repository as user_repository


######################
# Fixtures
######################


@pytest.fixture(name="department")
def department_fixture(department_factory):
    """Creates a reusable department for tests"""
    return department_factory()


@pytest.fixture(name="team")
def team_fixture(session, department):
    """Creates a reusable team for tests"""
    return team_repository.create_team(
        session=session,
        name="Team A",
        department_id=department.id
    )


@pytest.fixture(name="doctor")
def doctor_fixture(session, department, team):
    """Creates a reusable doctor for tests"""
    return doctor_repository.create_doctor(
        session=session,
        name="Dr Test",
        email="drtest@gmail.com",
        team_id=team.id,
        department_id=department.id,
    )


#######################
# Tests
#######################


def test_user_model_requires_explicit_role():
    """Tests that the User model requires an explicit role."""
    role_field = UserModel.model_fields["role"]
    assert role_field.is_required()


@pytest.mark.parametrize(
    ("role", "include_department", "include_doctor"),
    [
        (UserRole.SUPER_ADMIN, None, None),
        (UserRole.DEPARTMENT_ADMIN, True, None),
        (UserRole.DEPARTMENT_ADMIN, True, True),
        (UserRole.DOCTOR, True, True),
        (UserRole.VIEWER, True, None),
    ]
)
def test_create_user_account_accepts_valid_role_shape(
    role,
    include_department,
    include_doctor,
    doctor,
    department,
    session,
):
    """Tests that account creation accepts valid role shapes."""
    account_data = UserAccountCreate(
        email="testuser@gmail.com",
        role=role,
        full_name="Test User",
        password="test-password",
        department_id=department.id if include_department else None,
        doctor_id=doctor.id if include_doctor else None,
    )

    created_user = user_services.create_user_account(
        session=session,
        account_data=account_data,
    )

    assert created_user.role == role
    assert created_user.department_id == (
        department.id if include_department else None
    )
    assert created_user.doctor_id == (
        doctor.id if include_doctor else None
    )


@pytest.mark.parametrize(
    ("role", "include_department", "include_doctor"),
    [
        (UserRole.SUPER_ADMIN, True, None),
        (UserRole.SUPER_ADMIN, None, True),
        (UserRole.DEPARTMENT_ADMIN, None, None),
        (UserRole.DOCTOR, None, True),
        (UserRole.DOCTOR, True, None),
        (UserRole.VIEWER, None, None),
        (UserRole.VIEWER, True, True),
    ]
)
def test_create_user_account_rejects_invalid_role_shape(
    role,
    include_department,
    include_doctor,
    doctor,
    department,
    session,
):
    """Tests that account creation rejects invalid role shapes."""
    account_data = UserAccountCreate(
        email="testuser@gmail.com",
        role=role,
        full_name="Test User",
        password="test-password",
        department_id=department.id if include_department else None,
        doctor_id=doctor.id if include_doctor else None,
    )

    with pytest.raises(InvalidUserAccountRelationshipError) as exc_info:
        user_services.create_user_account(
            session=session,
            account_data=account_data,
        )

    assert "Invalid relationship" in str(exc_info.value)
    stored_user = user_repository.get_user_by_email(
        session=session,
        user_email=account_data.email,
    )
    assert stored_user is None


@pytest.mark.parametrize(
    ("role", "include_department", "include_doctor"),
    [
        (UserRole.SUPER_ADMIN, True, None),
        (UserRole.SUPER_ADMIN, None, True),
        (UserRole.DEPARTMENT_ADMIN, None, None),
        (UserRole.DOCTOR, None, True),
        (UserRole.DOCTOR, True, None),
        (UserRole.VIEWER, None, None),
        (UserRole.VIEWER, True, True),
    ]
)
def test_database_rejects_invalid_user_role_shape(
    role,
    include_department,
    include_doctor,
    session,
    doctor,
    department,
):
    """Tests that the database rejects invalid User role shapes."""
    user = UserModel(
        email="testuser@gmail.com",
        role=role,
        full_name="Test User",
        hashed_password="hashed-password",
        department_id=department.id if include_department else None,
        doctor_id=doctor.id if include_doctor else None,
    )
    session.add(user)

    with pytest.raises(IntegrityError) as exc_info:
        session.flush()

    assert "ck_user_role_shape" in str(exc_info.value)
    session.rollback()


def test_user_account_creation_rejects_nonexistent_department(
    session,
):
    """Tests that account creation rejects a nonexistent Department."""
    account_data = UserAccountCreate(
        email="testuser@gmail.com",
        role=UserRole.DEPARTMENT_ADMIN,
        full_name="Test User",
        password="test-password",
        department_id=uuid.uuid4(),
        doctor_id=None,
    )

    with pytest.raises(InvalidUserAccountRelationshipError) as exc_info:
        user_services.create_user_account(
            session=session,
            account_data=account_data,
        )

    assert "Invalid relationship" in str(exc_info.value)
    assert user_repository.get_user_by_email(
        session=session,
        user_email=account_data.email,
    ) is None


def test_user_account_creation_rejects_soft_deleted_department(
    session,
    department,
):
    """Tests that account creation rejects a soft-deleted Department."""
    department.is_deleted = True
    session.add(department)
    session.commit()

    account_data = UserAccountCreate(
        email="testuser@gmail.com",
        role=UserRole.DEPARTMENT_ADMIN,
        full_name="Test User",
        password="test-password",
        department_id=department.id,
        doctor_id=None,
    )

    with pytest.raises(InvalidUserAccountRelationshipError) as exc_info:
        user_services.create_user_account(
            session=session,
            account_data=account_data,
        )

    assert "Invalid relationship" in str(exc_info.value)
    assert user_repository.get_user_by_email(
        session=session,
        user_email=account_data.email,
    ) is None


def test_user_account_creation_rejects_malformed_department(
    session,
    doctor,
):
    """Tests that account input rejects a malformed Department ID."""
    with pytest.raises(ValueError) as exc_info:
        UserAccountCreate(
            email="testuser@gmail.com",
            role=UserRole.DEPARTMENT_ADMIN,
            full_name="Test User",
            password="test-password",
            department_id="not-a-uuid",
            doctor_id=doctor.id,
        )

    assert "Input should be a valid UUID" in str(exc_info.value)
    assert user_repository.get_user_by_email(
        session=session,
        user_email="testuser@gmail.com",
    ) is None


@pytest.mark.parametrize(
    "role",
    [
        UserRole.DOCTOR,
        UserRole.DEPARTMENT_ADMIN,
    ]
)
def test_user_account_creation_rejects_nonexistent_doctor(
    session,
    department,
    role,
):
    """Tests that account creation rejects a nonexistent Doctor."""
    account_data = UserAccountCreate(
        email="testuser@gmail.com",
        role=role,
        full_name="Test User",
        password="test-password",
        department_id=department.id,
        doctor_id=uuid.uuid4(),
    )

    with pytest.raises(InvalidUserAccountRelationshipError) as exc_info:
        user_services.create_user_account(
            session=session,
            account_data=account_data,
        )

    assert "Invalid relationship" in str(exc_info.value)
    assert user_repository.get_user_by_email(
        session=session,
        user_email=account_data.email,
    ) is None


@pytest.mark.parametrize(
    "role",
    [
        UserRole.DOCTOR,
        UserRole.DEPARTMENT_ADMIN,
    ]
)
def test_user_account_creation_rejects_soft_deleted_doctor(
    session,
    doctor,
    department,
    role,
):
    """Tests that account creation rejects a soft-deleted Doctor."""
    doctor.is_deleted = True
    session.add(doctor)
    session.commit()

    account_data = UserAccountCreate(
        email="testuser@gmail.com",
        role=role,
        full_name="Test User",
        password="test-password",
        department_id=department.id,
        doctor_id=doctor.id,
    )

    with pytest.raises(InvalidUserAccountRelationshipError) as exc_info:
        user_services.create_user_account(
            session=session,
            account_data=account_data,
        )

    assert "Invalid relationship" in str(exc_info.value)
    assert user_repository.get_user_by_email(
        session=session,
        user_email=account_data.email,
    ) is None


@pytest.mark.parametrize(
    "role",
    [
        UserRole.DOCTOR,
        UserRole.DEPARTMENT_ADMIN,
    ]
)
def test_user_account_creation_rejects_malformed_doctor(
    session,
    department,
    role,
):
    """Tests that account input rejects a malformed Doctor ID."""
    with pytest.raises(ValueError) as exc_info:
        UserAccountCreate(
            email="testuser@gmail.com",
            role=role,
            full_name="Test User",
            password="test-password",
            department_id=department.id,
            doctor_id="not-a-uuid",
        )

    assert "Input should be a valid UUID" in str(exc_info.value)
    assert user_repository.get_user_by_email(
        session=session,
        user_email="testuser@gmail.com",
    ) is None


@pytest.mark.parametrize(
    "role",
    [
        UserRole.DOCTOR,
        UserRole.DEPARTMENT_ADMIN,
    ]
)
def test_user_account_creation_rejects_doctor_from_foreign_department(
    session,
    doctor,
    department_factory,
    role,
):
    """Tests that account creation rejects a Doctor from another Department."""
    foreign_department = department_factory()
    account_data = UserAccountCreate(
        email="testuser@gmail.com",
        role=role,
        full_name="Test User",
        password="test-password",
        department_id=foreign_department.id,
        doctor_id=doctor.id,
    )

    with pytest.raises(InvalidUserAccountRelationshipError) as exc_info:
        user_services.create_user_account(
            session=session,
            account_data=account_data,
        )

    assert "Invalid relationship" in str(exc_info.value)
    assert user_repository.get_user_by_email(
        session=session,
        user_email=account_data.email,
    ) is None
