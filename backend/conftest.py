"""
Configuration for tests
"""

import pytest
from sqlmodel import SQLModel, create_engine, Session
from sqlalchemy.pool import StaticPool
from fastapi.testclient import TestClient

from src.db.connection import get_session
from src.main import app


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
def client_fixture(session):
    """Creates a TestClient for the FastAPI app with dependency override."""
    try:
        def override_get_session():
            yield session
        app.dependency_overrides[get_session] = override_get_session

        with TestClient(app) as test_client:
            yield test_client
    finally:
        app.dependency_overrides.clear()
