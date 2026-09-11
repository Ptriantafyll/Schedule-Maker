"""
Tests for User role and relationship contracts.
"""

from src.auth.security import hash_password
import uuid

import pytest
from sqlalchemy.exc import IntegrityError
from sqlmodel import select

from src.team import repository as team_repository
from src.doctor import repository as doctor_repository
from src.user.models import User as UserModel
from src.user.models import UserRole
from src.user.schemas import UserAccountCreate
from src.user import services as user_services
from src.user.services import InvalidUserAccountRelationshipError, DoctorAlreadyLinkedError
from src.user import repository as user_repository
from src.doctor.models import Doctor as DoctorModel

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


def test_user_account_creation_rejects_nonexistent_doctor(
    session,
    department,
):
    """Tests that account creation rejects a nonexistent Doctor."""
    account_data = UserAccountCreate(
        email="testuser@gmail.com",
        role=UserRole.DOCTOR,
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


def test_user_account_creation_rejects_soft_deleted_doctor(
    session,
    doctor,
    department,
):
    """Tests that account creation rejects a soft-deleted Doctor."""
    doctor.is_deleted = True
    session.add(doctor)
    session.commit()

    account_data = UserAccountCreate(
        email="testuser@gmail.com",
        role=UserRole.DOCTOR,
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


def test_doctor_can_exist_without_user_account(
    session,
    doctor_factory
):
    """Tests that a doctor can exist in the DB without a user"""
    doctor = doctor_factory()

    retrieved_doctor = session.get(
        DoctorModel, doctor.id
    )

    assert retrieved_doctor is not None
    assert retrieved_doctor.id == doctor.id

    linked_users = session.exec(
        select(UserModel).where(
            UserModel.doctor_id == doctor.id
        )
    ).all()
    assert linked_users == []


def test_doctor_can_have_one_active_user_account(
    session,
    doctor_factory,
    user_factory,
):
    """Tests that a doctor can be linked to a user"""
    doctor = doctor_factory()
    user = user_factory(
        role=UserRole.DOCTOR,
        department_id=doctor.department_id,
        doctor_id=doctor.id
    )

    retrieved_doctor = session.get(
        DoctorModel, doctor.id
    )

    retrieved_user = session.get(
        UserModel, user.id
    )

    assert retrieved_doctor is not None and retrieved_user is not None
    assert retrieved_doctor.id == retrieved_user.doctor_id


def test_second_active_user_for_doctor_is_rejected(
    session,
    user_factory,
    doctor_factory,
):
    """Tests that a doctor cannot have more than 1 user linked to them"""
    doctor = doctor_factory()
    user_a = user_factory(
        role=UserRole.DOCTOR,
        department_id=doctor.department_id,
        doctor_id=doctor.id,
        email="doctor_a@example.com",
    )

    with pytest.raises(DoctorAlreadyLinkedError) as exc_info:
        user_factory(
            role=UserRole.DOCTOR,
            department_id=doctor.department_id,
            doctor_id=doctor.id,
            email="doctor_b@example.com",
        )

    assert "Doctor is already linked" in str(exc_info.value)

    stored_users = session.exec(
        select(UserModel).where(UserModel.doctor_id == doctor.id)
    ).all()
    assert len(stored_users) == 1
    assert stored_users[0].id == user_a.id
    assert stored_users[0].email == "doctor_a@example.com"

    rejected_user = user_repository.get_user_by_email(
        session, "doctor_b@example.com")
    assert rejected_user is None


def test_department_admin_can_link_same_department_doctor(
    session,
    department_factory,
    doctor_factory,
    user_factory
):
    """Tests that a DEPARTMENT_ADMIN can link to a doctor in their own department."""
    department = department_factory()
    doctor = doctor_factory(department_id=department.id)

    user = user_factory(
        role=UserRole.DEPARTMENT_ADMIN,
        department_id=department.id,
        doctor_id=doctor.id
    )

    retrieved_doctor = session.get(
        DoctorModel, doctor.id
    )

    retrieved_user = session.get(
        UserModel, user.id
    )

    assert retrieved_doctor is not None and retrieved_user is not None
    assert retrieved_doctor.id == retrieved_user.doctor_id


def test_department_admin_doctor_link_blocks_second_doctor_user(
    session,
    department_factory,
    doctor_factory,
    user_factory,
):
    """Tests that linking a user to an already linked doctor fails"""
    department = department_factory()
    doctor = doctor_factory(department_id=department.id)

    user = user_factory(
        role=UserRole.DEPARTMENT_ADMIN,
        department_id=department.id,
        doctor_id=doctor.id,
        email="doctor_a@example.com",
    )

    with pytest.raises(DoctorAlreadyLinkedError) as exc_info:
        user_factory(
            role=UserRole.DOCTOR,
            department_id=department.id,
            doctor_id=doctor.id,
            email="doctor_b@example.com",
        )

    assert "Doctor is already linked" in str(exc_info.value)

    stored_users = session.exec(
        select(UserModel).where(UserModel.doctor_id == doctor.id)
    ).all()
    assert len(stored_users) == 1
    assert stored_users[0].id == user.id
    assert stored_users[0].email == "doctor_a@example.com"

    rejected_user = user_repository.get_user_by_email(
        session, "doctor_b@example.com")
    assert rejected_user is None


def test_different_doctors_can_have_active_accounts(
    session,
    department_factory,
    doctor_factory,
    user_factory,
):
    """Tests that creating multiple doctors linked to different users works"""
    dept = department_factory()
    doctor_a = doctor_factory(department_id=dept.id)
    doctor_b = doctor_factory(department_id=dept.id)

    user_a = user_factory(
        role=UserRole.DOCTOR,
        department_id=dept.id,
        doctor_id=doctor_a.id
    )

    user_b = user_factory(
        role=UserRole.DOCTOR,
        department_id=dept.id,
        doctor_id=doctor_b.id
    )

    # Assert persistence
    retrieved_1 = session.get(UserModel, user_a.id)
    retrieved_2 = session.get(UserModel, user_b.id)
    assert retrieved_1 is not None
    assert retrieved_2 is not None

    # Assert correct ownership
    assert retrieved_1.doctor_id == doctor_a.id
    assert retrieved_2.doctor_id == doctor_b.id
    assert retrieved_1.doctor_id != retrieved_2.doctor_id


@pytest.mark.parametrize(
    "role",
    [
        UserRole.VIEWER,
        UserRole.DEPARTMENT_ADMIN,
    ]
)
def test_non_doctor_users_do_not_conflict_on_null_doctor_id(
    session,
    user_factory,
    department_factory,
    role,
):
    """Tests that creating a non-doctor user with no doctor link works"""
    department = department_factory()
    user = user_factory(
        role=role,
        department_id=department.id,
        doctor_id=None,
    )

    retrieved_user = session.get(UserModel, user.id)

    assert retrieved_user is not None
    assert retrieved_user.department_id == department.id


def test_soft_deleted_doctor_user_allows_replacement_account(
    session,
    user_factory,
    doctor_factory,
):
    """Tests that deleting a user removes the doctor link and allows doctor to be linked to another user"""
    doctor_a = doctor_factory()
    user_a = user_factory(
        role=UserRole.DOCTOR,
        department_id=doctor_a.department_id,
        doctor_id=doctor_a.id,
        email="user1@test.com",
    )

    user_a.is_deleted = True
    session.add(user_a)
    session.commit()

    user_b = user_factory(
        role=UserRole.DOCTOR,
        department_id=doctor_a.department_id,
        doctor_id=doctor_a.id,
        email="user2@test.com",
    )

    retrieved_user = session.get(UserModel, user_b.id)
    assert retrieved_user is not None
    assert retrieved_user.doctor_id == doctor_a.id


def test_database_rejects_concurrent_active_doctor_link(
    session,
    department_factory,
    doctor_factory,
):
    """Tests that the database-level partial unique index rejects duplicate active doctor links."""
    dept = department_factory()
    doctor = doctor_factory(department_id=dept.id)

    user_1 = UserModel(
        email="user1@test.com",
        full_name="User 1",
        hashed_password=hash_password("pwd"),
        role=UserRole.DOCTOR,
        department_id=dept.id,
        doctor_id=doctor.id,
        is_deleted=False,
    )
    session.add(user_1)
    session.commit()

    # Directly bypass user_services and add a second user with the same doctor_id to test raw DB index
    user_2 = UserModel(
        email="user2@test.com",
        full_name="User 2",
        hashed_password=hash_password("pwd"),
        role=UserRole.DOCTOR,
        department_id=dept.id,
        doctor_id=doctor.id,
        is_deleted=False,
    )
    session.add(user_2)
    with pytest.raises(IntegrityError):
        session.commit()

    session.rollback()


def test_department_admin_rejects_foreign_or_deleted_doctor_link(
    session,
    department_factory,
    doctor_factory,
    user_factory,
):
    """Tests that a department admin cannot link to a deleted or foreign doctor."""
    dept_a = department_factory()
    dept_b = department_factory()
    foreign_doc = doctor_factory(department_id=dept_b.id)

    # Rejects foreign doctor
    with pytest.raises(InvalidUserAccountRelationshipError):
        user_factory(
            role=UserRole.DEPARTMENT_ADMIN,
            department_id=dept_a.id,
            doctor_id=foreign_doc.id,
        )

    # Rejects deleted doctor
    local_doc = doctor_factory(department_id=dept_a.id)
    local_doc.is_deleted = True
    session.add(local_doc)
    session.commit()

    with pytest.raises(InvalidUserAccountRelationshipError):
        user_factory(
            role=UserRole.DEPARTMENT_ADMIN,
            department_id=dept_a.id,
            doctor_id=local_doc.id,
        )
