"""
Configuration for tests
"""

import uuid
import pytest
from sqlmodel import SQLModel, create_engine, Session
from sqlalchemy.pool import StaticPool
from fastapi.testclient import TestClient

from src.db.connection import get_session
from src.main import app
from src.user.models import UserRole
from src.user.models import User as UserModel
from src.user.schemas import UserCreate
from src.user import controllers as user_controllers
from src.auth.security import create_access_token


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


@pytest.fixture(name="client")
def client_fixture(session, monkeypatch):
    """Creates a TestClient for the FastAPI app with dependency override."""
    try:
        def override_get_session():
            yield session

        app.dependency_overrides[get_session] = override_get_session
        monkeypatch.setattr("src.main.init_db", lambda: None)

        with TestClient(app) as test_client:
            yield test_client
    finally:
        app.dependency_overrides.clear()


@pytest.fixture(name="user_factory")
def user_factory_fixture(session):
    """Create persisted users with customizable roles and relationships"""

    def create_user(
        *,
        role: UserRole,
        department_id: uuid.UUID | None = None,
        doctor_id: uuid.UUID | None = None,
        email: str | None = None,
        full_name: str = "Test User",
        password: str = "test-password",
    ) -> UserModel:
        user_email = email or f"user-{uuid.uuid4().hex}@test.com"

        user_data = UserCreate(
            email=user_email,
            full_name=full_name,
            password=password,
            role=role,
            department_id=department_id,
            doctor_id=doctor_id,
        )

        return user_controllers.create_user_controller(
            user_data=user_data,
            session=session,
        )

    return create_user


@pytest.fixture(name="auth_headers_factory")
def auth_headers_factory_fixture():
    """Create auth headers factory fixture"""

    def create_auth_headers(user: UserModel) -> dict[str, str]:
        access_token = create_access_token({"sub": str(user.id)})

        return {"Authorization": f"Bearer {access_token}"}

    return create_auth_headers
