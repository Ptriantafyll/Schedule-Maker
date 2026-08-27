"""
Tests for the user module
"""


import uuid
import pytest
import datetime
from fastapi.testclient import TestClient
from sqlalchemy.pool import StaticPool
from sqlmodel import SQLModel, create_engine, Session

from src.main import app
from src.user.schemas import UserCreate
from src.user.models import UserRole
from src.user import repository as user_repository
from src.user import controllers as user_controllers
from src.department.schemas import DepartmentCreate
from src.department import repository as department_repository
from src.auth.security import create_access_token


#####################
# Helpers
#####################

#####################
# Fixtures
#####################


@pytest.fixture(name="session")
def session_fixture():
    """Creates a fresh in-memory database session for each test."""
    engine = create_engine(
        "sqlite:///:memory:",
        connect_args={"check_same_thread": False},
        poolclass=StaticPool,
    )
    SQLModel.metadata.create_all(engine)
    with Session(engine) as session:
        yield session


@pytest.fixture(name="user")
def user_fixture(session):
    """Creates a reusable user for tests"""
    user_data = UserCreate(
        email="test@gmail.com",
        password="test123",
        role="doctor",
        full_name="Test testakis",
        doctor_id=None,
        department_id=None
    )
    return user_repository.create_user(session, user_data)


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
def admin_user_fixture(session, department):
    """Creates a reusable admin user for tests"""
    admin_user_data = UserCreate(
        email="admin@gmail.com",
        full_name="admin admin",
        password="password123",
        role=UserRole.DEPARTMENT_ADMIN,
        department_id=department.id
    )

    return user_controllers.create_user_controller(admin_user_data, session)


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
def admin_headers_fixture(department_admin_user):
    """Creates reusable admin headers"""
    access_token = create_access_token({"sub": str(department_admin_user.id)})

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


#####################
# Route tests
#####################

@pytest.fixture(name="client")
def client_fixture(session):
    """Creates a TestClient for the FastAPI app with dependency override."""
    from src.db.connection import get_session

    def override_get_session():
        yield session
    app.dependency_overrides[get_session] = override_get_session

    with TestClient(app) as test_client:
        yield test_client

    app.dependency_overrides.clear()


def test_create_user_route(client):
    """Tests POST /api/v1/users route"""
    response = client.post(
        "api/v1/users/signup",
        json={
            "full_name": "Test2 Testakis",
            "email": "test@gmail.com",
            "password": "test123",
        }
    )

    assert response.status_code == 201
    data = response.json()
    assert data["full_name"] == "Test2 Testakis"
    assert data["role"] == "doctor"
    assert data["email"] == "test@gmail.com"
    assert "id" in data
    assert "created_at" in data
    assert "updated_at" in data


def test_create_user_route_invalid_payload(client):
    """Tests POST /api/v1/users route with invalid payload returns error"""
    response = client.post(
        "api/v1/users/signup",
        json={
            "password": "test123",
        }
    )

    assert response.status_code == 422


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


def test_get_user_route_nonexistent_email(client):
    """Tests GET /api/v1/users/{user_email} with invalid payload"""
    response = client.get("/api/v1/users/123")

    assert response.status_code == 404


@pytest.mark.parametrize(
    "role",
    [
        UserRole.DEPARTMENT_ADMIN,
        UserRole.SUPER_ADMIN,
    ],
)
def test_public_signup_rejects_privileged_roles(client, session, role):
    """Test POST /api/v1/users/signup should reject privileged roles"""
    email = f"{role.value}@test.com"
    response = client.post(
        "api/v1/users/signup",
        json={
            "full_name": "Test2 Testakis",
            "email": email,
            "password": "test123",
            "role": role.value,
        }
    )

    assert response.status_code == 422
    assert user_repository.get_user_by_email(session, "test@gmail.com") is None
