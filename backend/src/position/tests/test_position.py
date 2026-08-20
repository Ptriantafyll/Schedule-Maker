"""
Tests for the position module
"""

import uuid
import datetime
import pytest
from fastapi.testclient import TestClient
from sqlalchemy.pool import StaticPool
from sqlmodel import SQLModel, create_engine, Session

from src.main import app
from src.position.schemas import PositionCreate
from src.position.models import Position as PositionModel
from src.position import repository as position_repository
from src.position import controllers as position_controllers
from src.department.schemas import DepartmentCreate
from src.department import repository as department_repository

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


@pytest.fixture(name="department")
def department_fixture(session):
    """Creates a reusable department for tests"""
    dept_data = DepartmentCreate(name="Cardiology", code="CARD")
    return department_repository.create_department(session, dept_data)


@pytest.fixture(name="position")
def position_fixture(session, department):
    """Creates a reusable position for tests"""
    position_data = PositionCreate(
        name="ER",
        department_id=department.id,
        duty_days=[1, 3, 5],
    )

    return position_repository.create_position(session, position_data)
#####################
# Repository tests
#####################


def test_create_position(position):
    """Tests creating a position in the db"""
    assert isinstance(position.id, uuid.UUID)
    assert position.is_deleted is False
    assert position.sync_status is False
    assert isinstance(position.created_at, datetime.datetime)
    assert isinstance(position.updated_at, datetime.datetime)


def test_get_position_by_id(session, position):
    """Tests retrieving a position from the db"""
    retrieved_position = position_repository.get_position_by_id(
        session=session,
        position_id=position.id
    )

    assert retrieved_position is not None
    assert retrieved_position.id == position.id


def test_get_position_by_name(session, position):
    """Tests retrieving a position by its name"""
    retrieved_position = position_repository.get_position_by_name(
        session=session,
        position_name=position.name
    )

    assert retrieved_position is not None
    assert retrieved_position.id == position.id


def test_get_active_positions(session, position):
    """Tests retrieving all active positions"""
    position_data = PositionCreate(
        name="Clinic",
        department_id=position.department_id,
        duty_days=[4, 5, 6]
    )
    position2 = position_repository.create_position(session, position_data)

    position2.is_deleted = True
    session.add(position2)
    session.commit()

    retrieved_positions = position_repository.get_active_positions(session)

    assert isinstance(retrieved_positions, list)
    assert position in retrieved_positions
    assert position2 not in retrieved_positions

#####################
# Controller tests
#####################


def test_create_position_controller_duplicate_name(session, position):
    """Tests that creating a position with a duplicate name returns error"""
    position_data = PositionCreate(
        name=position.name,
        department_id=position.department_id,
        duty_days=position.duty_days
    )

    with pytest.raises(Exception) as exc_info:
        position_controllers.create_position_controller(position_data, session)

    assert exc_info.type.__name__ == "HTTPException"
    assert exc_info.value.status_code == 400
    assert "already exists" in exc_info.value.detail


def test_create_position_controller_nonexistent_department(session):
    """Tests that creating a position with a non existent position id returns error"""
    position_data = PositionCreate(
        name="ER",
        department_id=uuid.uuid4(),
        duty_days=[1, 2, 3]
    )

    with pytest.raises(Exception) as exc_info:
        position_controllers.create_position_controller(position_data, session)

    assert exc_info.type.__name__ == "HTTPException"
    assert exc_info.value.status_code == 422
    assert "does not exist" in exc_info.value.detail


def test_get_position_controller_nonexistent(session):
    """Tests that trying to retrieve a non existent position returns error"""
    with pytest.raises(Exception) as exc_info:
        position_controllers.get_position_controller("test", session)

    assert exc_info.type.__name__ == "HTTPException"
    assert exc_info.value.status_code == 404
    assert "not found" in exc_info.value.detail


def test_get_position_controller_deleted(session, position):
    """Tests that trying to retrieve a deleted position returns error"""
    position.is_deleted = True
    session.add(position)
    session.commit()

    with pytest.raises(Exception) as exc_info:
        position_controllers.get_position_controller(position.name, session)

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


def test_create_position_route(client, department):
    """Tests post /api/v1/positions route"""
    response = client.post(
        "api/v1/positions",
        json={
            "name": "ER",
            "department_id": str(department.id),
            "duty_days": [1, 2, 3]
        }
    )

    assert response.status_code == 201
    data = response.json()
    assert data["name"] == "ER"
    assert "duty_days" in data
    assert "id" in data
    assert "created_at" in data
    assert "updated_at" in data


def test_create_position_route_invalid_payload(client):
    """Tests post /api/v1/positions route with invalid payload"""
    response = client.post(
        "api/v1/positions",
        json={
            "name": "ER",
        }
    )

    assert response.status_code == 422


def test_list_positions_route(client, position):
    """Tests get /api/v1/positionsroute"""
    response = client.get("/api/v1/positions")

    assert response.status_code == 200
    data = response.json()
    assert isinstance(data, list)
    assert data[0]["id"] == str(position.id)


def test_get_position_route(client, position):
    """Tests get /api/v1/positions/{position_name} route"""
    response = client.get(f"/api/v1/positions/{position.name}")

    assert response.status_code == 200
    data = response.json()
    assert data["id"] == str(position.id)


def test_get_position_route_nonexistent_name(client):
    """Tests get /api/v1/positions/{position_name} route with invalid payload"""
    response = client.get("/api/v1/positions/test")

    assert response.status_code == 404
