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
from src.user.models import User as UserModel
from src.user import repository as user_repository
from src.user import controllers as user_controllers


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
        role="admin",
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
        role="admin",
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
            "role": "admin",
            "email": "test@gmail.com",
            "password": "test123",
        }
    )

    assert response.status_code == 201
    data = response.json()
    assert data["full_name"] == "Test2 Testakis"
    assert data["role"] == "admin"
    assert data["email"] == "test@gmail.com"
    assert "id" in data
    assert "created_at" in data
    assert "updated_at" in data


def test_create_user_route_invalid_payload(client):
    """Tests POST /api/v1/users route with invalid payload returns error"""
    response = client.post(
        "api/v1/users/signup",
        json={
            "role": "admin",
            "password": "test123",
        }
    )

    assert response.status_code == 422


def test_list_users_route(client, user):
    """Tests GET /api/v1/users route"""
    response = client.get("/api/v1/users")

    assert response.status_code == 200
    data = response.json()
    assert isinstance(data, list)
    assert data[0]["id"] == str(user.id)


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
