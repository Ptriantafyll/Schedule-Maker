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
from src.department import repository as department_repository
from src.department.models import Department as DepartmentModel
from src.department.schemas import DepartmentCreate

PLAIN_SUPER_ADMIN_PASSWORD = "SecurePassword123!"
SUPER_ADMIN_EMAIL = "superadmin@test.com"
PLAIN_DEPT_ADMIN_PASSWORD = "SecurePassword123!"
DEPT_ADMIN_EMAIL = "deptadmin@test.com"


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
    mock_get = iter([None, existing_super_admin])
    monkeypatch.setattr(
        "src.user.services.repository.get_user_by_email",
        lambda *args, **kwargs: next(mock_get, existing_super_admin),
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


def test_create_department_admin_creates_valid_account_and_department(session):
    """Tests creating a department admin and a new department."""
    dept, created_user = bootstrap.create_department_admin(
        session=session,
        department_name="Cardiology",
        department_code="CARD",
        email=DEPT_ADMIN_EMAIL,
        full_name="Dr. Cardio Admin",
        password=PLAIN_DEPT_ADMIN_PASSWORD,
    )

    # Invariants on the user
    assert created_user.role == UserRole.DEPARTMENT_ADMIN
    assert created_user.department_id == dept.id
    assert created_user.doctor_id is None
    assert verify_password(PLAIN_DEPT_ADMIN_PASSWORD, created_user.hashed_password)

    # Invariants on the department
    assert dept.name == "Cardiology"
    assert dept.code == "CARD"

    # Verify database persistence
    persisted_user = user_repository.get_user_by_email(session, DEPT_ADMIN_EMAIL)
    assert persisted_user is not None
    assert persisted_user.id == created_user.id

    persisted_dept = department_repository.get_department_by_name_global(
        session, "Cardiology"
    )
    assert persisted_dept is not None
    assert persisted_dept.id == dept.id


def test_create_department_admin_derives_default_code_when_omitted(session):
    """Tests that department_code defaults to uppercase prefix when omitted."""
    dept, created_user = bootstrap.create_department_admin(
        session=session,
        department_name="Neurology",
        department_code=None,
        email="neuro.admin@test.com",
        full_name="Dr. Neuro Admin",
        password=PLAIN_DEPT_ADMIN_PASSWORD,
    )

    assert dept.code == "NEUR"
    assert created_user.department_id == dept.id


def test_create_department_admin_reuses_existing_department(session):
    """Tests that an existing department is reused rather than duplicated."""
    existing_dept = department_repository.create_department(
        session,
        DepartmentCreate(name="Cardiology", code="CARD"),
    )

    dept, user = bootstrap.create_department_admin(
        session=session,
        department_name="Cardiology",
        department_code=None,
        email="another.admin@cardio.com",
        full_name="Second Admin",
        password=PLAIN_DEPT_ADMIN_PASSWORD,
    )

    assert dept.id == existing_dept.id
    all_depts = department_repository.get_active_departments_global(session)
    cardio_depts = [d for d in all_depts if d.name == "Cardiology"]
    assert len(cardio_depts) == 1
    assert user.department_id == existing_dept.id


def test_create_department_admin_rejects_duplicate_email(session):
    """Tests creating a department admin with an email that already exists."""
    bootstrap.create_department_admin(
        session=session,
        department_name="Cardiology",
        department_code="CARD",
        email=DEPT_ADMIN_EMAIL,
        full_name="First Admin",
        password=PLAIN_DEPT_ADMIN_PASSWORD,
    )

    with pytest.raises(bootstrap.DepartmentAdminAlreadyExistsError):
        bootstrap.create_department_admin(
            session=session,
            department_name="Cardiology",
            department_code="CARD",
            email=DEPT_ADMIN_EMAIL,
            full_name="Duplicate Admin",
            password=PLAIN_DEPT_ADMIN_PASSWORD,
        )


def test_create_department_admin_rolls_back_department_on_user_creation_failure(
    session, monkeypatch
):
    """Tests that a newly created department is rolled back if user creation fails."""
    def _mock_failure(*args, **kwargs):
        raise RuntimeError("Simulated failure during user creation")

    monkeypatch.setattr(
        "src.user.services.create_user_account",
        _mock_failure,
    )

    with pytest.raises(RuntimeError):
        bootstrap.create_department_admin(
            session=session,
            department_name="UnsavedDept",
            department_code="UNSV",
            email="unsaved@test.com",
            full_name="Unsaved Admin",
            password=PLAIN_DEPT_ADMIN_PASSWORD,
        )

    # Department must not be persisted
    persisted_dept = department_repository.get_department_by_name_global(
        session, "UnsavedDept"
    )
    assert persisted_dept is None
