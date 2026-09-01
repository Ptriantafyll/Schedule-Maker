"""
Tests for the user module
"""


import uuid
import datetime
import pytest
from sqlmodel import Session

from src.user.schemas import UserCreate
from src.user.models import UserRole
from src.user import repository as user_repository
from src.user import controllers as user_controllers
from src.department.schemas import DepartmentCreate
from src.department import repository as department_repository
from src.auth.security import create_access_token
from src.doctor import repository as doctor_repository
from src.doctor.schemas import DoctorCreate
from src.doctor.models import Doctor as DoctorModel
from src.team import repository as team_repository
from src.team.schemas import TeamCreate
from src.auth.security import verify_password

#####################
# Helpers
#####################


def create_new_doctor(
    session: Session,
    name: str,
    email: str,
    department_id: uuid.UUID,
    team_id: uuid.UUID
) -> DoctorModel:
    """Helper that creates a new doctor in the db"""
    doctor_data = DoctorCreate(
        name=name,
        email=email,
        department_id=department_id,
        team_id=team_id
    )
    return doctor_repository.create_doctor(session, doctor_data)
#####################
# Fixtures
#####################


@pytest.fixture(name="department")
def department_fixture(session):
    """Creates a reusable department for tests"""
    dept_data = DepartmentCreate(name="Cardiology", code="CARD")
    return department_repository.create_department(session, dept_data)


@pytest.fixture(name="department_b")
def department_b_fixture(session):
    """Creates a reusable department for tests"""
    dept_data = DepartmentCreate(name="Radiology", code="RAD")
    return department_repository.create_department(session, dept_data)


@pytest.fixture(name="department_admin_user")
def department_admin_user_fixture(user_factory, department):
    """Creates a reusable admin user for tests"""
    return user_factory(
        role=UserRole.DEPARTMENT_ADMIN,
        department_id=department.id,
        email="admin@gmail.com",
        full_name="Admin User",
        password="test-password",
    )


@pytest.fixture(name="department_b_user")
def department_b_user(session, department_b):
    """Creates a user for the department B"""
    user_data = UserCreate(
        email="user@gmail.com",
        full_name="test test",
        password="password123",
        role=UserRole.VIEWER,
        department_id=department_b.id
    )
    return user_controllers.create_user_controller(user_data, session)


@pytest.fixture(name="department_admin_headers")
def admin_headers_fixture(department_admin_user, auth_headers_factory):
    """Creates reusable admin headers"""
    return auth_headers_factory(department_admin_user)


@pytest.fixture(name="team")
def team_fixture(session, department):
    """Creates a reusable team for tests"""
    team_data = TeamCreate(name="ER Team A", department_id=department.id)
    return team_repository.create_team(session, team_data)


@pytest.fixture(name="doctor")
def doctor_fixture(session, department, team):
    """Creates a reusable doctor for tests"""
    return create_new_doctor(session, "Dr Panos", "drpanos@gmail.com", department.id, team.id)


@pytest.fixture(name="user")
def user_fixture(session, doctor, department):
    """Creates a reusable user for tests"""
    user_data = UserCreate(
        email="test@gmail.com",
        password="test123",
        role=UserRole.DOCTOR,
        full_name="Test testakis",
        doctor_id=doctor.id,
        department_id=department.id
    )
    return user_repository.create_user(session, user_data)


@pytest.fixture(name="doctor_headers")
def doctor_headers_fixture(user):
    """Creates reusable doctor headers"""
    access_token = create_access_token({"sub": str(user.id)})

    return {"Authorization": f"Bearer {access_token}"}

#####################
# Repository tests
#####################


def test_create_user(user):
    """Test creating a user and verifying their fields"""
    assert isinstance(user.id, uuid.UUID)
    assert user.email == "test@gmail.com"
    assert isinstance(user.hashed_password, str)
    assert user.role == "doctor"
    assert user.full_name == "Test testakis"
    assert user.is_deleted is False
    assert user.sync_status is False
    assert isinstance(user.created_at, datetime.datetime)
    assert isinstance(user.updated_at, datetime.datetime)


def test_get_user_by_email(session, user):
    """Test retrieving a user by email"""
    retrieved_user = user_repository.get_user_by_email(
        session, "test@gmail.com")

    assert retrieved_user is not None
    assert retrieved_user.id == user.id


def test_get_user_by_id(session, user):
    """Test retrieving a user by id"""

    retrieved_user = user_repository.get_user_by_id(
        session, user.id)

    assert retrieved_user is not None
    assert retrieved_user.id == user.id


def test_get_active_users(session, user):
    """Test listing all active users"""
    new_user_data = UserCreate(
        full_name="Test2 Testakis",
        role=UserRole.VIEWER,
        email="test2@gmail.com",
        password="test123",
        doctor_id=None,
        department_id=None
    )

    new_user = user_repository.create_user(session, new_user_data)

    new_user.is_deleted = True
    session.add(new_user)
    session.commit()

    active_users = user_repository.get_active_users(session)

    assert user in active_users
    assert new_user not in active_users


def test_get_active_users_by_department(
    session,
    department,
    department_b,
    user_factory,
):
    """Tests listing users in a department"""
    department_a_admin = user_factory(
        role=UserRole.VIEWER,
        department_id=department.id,
    )

    department_a_viewer = user_factory(
        role=UserRole.VIEWER,
        department_id=department.id,
    )

    department_b_viewer = user_factory(
        role=UserRole.VIEWER,
        department_id=department_b.id,
    )

    department_a_del_user = user_factory(
        role=UserRole.VIEWER,
        department_id=department.id,
    )
    department_a_del_user.is_deleted = True
    session.add(department_a_del_user)
    session.commit()

    super_admin_user = user_factory(
        role=UserRole.SUPER_ADMIN,
        department_id=None,
    )

    active_users = user_repository.get_active_users_by_department(
        session=session,
        department_id=department.id
    )

    returned_ids = {user.id for user in active_users}
    assert department_b_user.id not in returned_ids
    assert super_admin_user.id not in returned_ids
    assert returned_ids == {
        department_a_admin.id,
        department_a_viewer.id
    }


#####################
# Controller tests
#####################


def test_create_user_controller_duplicate_email(session, user):
    """Tests that creating a user with a duplicate email returns error"""
    new_user_data = UserCreate(
        full_name="Test2 Testakis",
        role=UserRole.DOCTOR,
        email=user.email,
        password="test123",
        doctor_id=None,
        department_id=None
    )

    with pytest.raises(Exception) as exc_info:
        user_controllers.create_user_controller(new_user_data, session)

    assert exc_info.type.__name__ == "HTTPException"
    assert exc_info.value.status_code == 400
    assert "already exists" in exc_info.value.detail


def test_user_controller_nonexistent(session, user):
    """Tests that trying to retrieve a deleted position returns error"""
    user.is_deleted = True
    session.add(user)
    session.commit()

    with pytest.raises(Exception) as exc_info:
        user_controllers.get_user_controller(user.email, session)

    assert exc_info.type.__name__ == "HTTPException"
    assert exc_info.value.status_code == 404
    assert "not found" in exc_info.value.detail


def test_create_user_controller_hashes_password(session, department):
    """Tests that a password gets hashed"""
    plain_password = "test123"
    new_user_data = UserCreate(
        full_name="Test2 Testakis",
        role=UserRole.VIEWER,
        email="test@test.com",
        password=plain_password,
        department_id=department.id
    )

    new_user = user_controllers.create_user_controller(new_user_data, session)

    assert new_user is not None
    assert new_user.hashed_password != plain_password
    assert verify_password(plain_password, new_user.hashed_password)
    assert not verify_password("wrong_password", new_user.hashed_password)


#####################
# Route tests
#####################

def test_department_admin_lists_users_only_in_own_department(client, department_admin_user, department_admin_headers, department_b_user, department):
    """Tests GET /api/v1/users route with sufficient permissions"""
    response = client.get(
        "/api/v1/users",
        headers=department_admin_headers
    )

    assert response.status_code == 200
    data = response.json()
    assert isinstance(data, list)

    returned_admin = next(
        (item for item in data if item["id"] == str(department_admin_user.id)),
        None
    )

    assert returned_admin is not None
    assert returned_admin["email"] == department_admin_user.email
    assert returned_admin["role"] == "department_admin"
    assert "hashed_password" not in returned_admin

    returned_ids = {item["id"] for item in data}

    assert str(department_admin_user.id) in returned_ids
    assert str(department_b_user.id) not in returned_ids
    assert all(
        item["department_id"] == str(department.id)
        for item in data
    )


def test_list_users_requires_authentication(client):
    """Tests that an authenticated route returns error on unauthenticated request"""
    response = client.get("/api/v1/users")

    assert response.status_code == 401
    data = response.json()
    assert data == {"detail": "Unauthorized"}
    assert response.headers.get("WWW-Authenticate") == "Bearer"


def test_list_users_rejects_invalid_token(client):
    """Tests that an authenticated route rejects an invalid token"""
    response = client.get(
        "/api/v1/users",
        headers={"Authorization": "Bearer invalid token"}
    )

    assert response.status_code == 401
    data = response.json()
    assert data == {"detail": "Unauthorized"}
    assert response.headers.get("WWW-Authenticate") == "Bearer"


def test_deleted_user_token_is_rejected(session, client, department_admin_user, department_admin_headers):
    """Test that a deleted user's token is rejected"""
    department_admin_user.is_deleted = True
    session.add(department_admin_user)
    session.commit()

    response = client.get(
        "/api/v1/departments",
        headers=department_admin_headers
    )

    assert response.status_code == 401
    data = response.json()

    assert data == {"detail": "User account no longer active"}
    assert response.headers.get("WWW-Authenticate") == "Bearer"


def test_department_admin_can_create_team(client, department_admin_headers, department, session):
    """Tests that a department admin can create a team"""
    response = client.post(
        "/api/v1/teams",
        json={"name": "Rad Team E", "department_id": str(department.id)},
        headers=department_admin_headers
    )

    assert response.status_code == 201
    data = response.json()
    assert data["name"] == "Rad Team E"
    assert data["department_id"] == str(department.id)
    new_team = team_repository.get_team_by_name(session, "Rad Team E")
    assert new_team is not None
    assert str(new_team.id) == data["id"]


@pytest.mark.parametrize(
    "role",
    [
        UserRole.DOCTOR,
        UserRole.VIEWER,
        UserRole.SUPER_ADMIN,
    ]
)
def test_user_list_rejects_non_department_admin_roles(client, user_factory, auth_headers_factory, role, department, doctor):
    """GET /api/v1/users rejects non department_admin_roles"""
    test_user = user_factory(
        role=role,
        department_id=(
            None
            if role == UserRole.SUPER_ADMIN
            else department.id
        ),
        doctor_id=(
            doctor.id
            if role == UserRole.DOCTOR
            else None
        ),
    )

    headers = auth_headers_factory(test_user)

    response = client.get(
        "/api/v1/users",
        headers=headers,
    )

    assert response.status_code == 403
    assert response.json() == {
        "detail": "Insufficient permissions for this operation"
    }
    assert response.headers.get("WWW-Authenticate") is None
