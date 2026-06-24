"""
Tests for the team module
"""

import uuid
import datetime
import pytest
from fastapi.testclient import TestClient
from sqlalchemy.pool import StaticPool
from sqlmodel import SQLModel, create_engine, Session

from src.main import app
from src.team.schemas import TeamCreate
from src.team.repository import create_team, get_active_teams, get_team_by_name
from src.department.schemas import DepartmentCreate
from src.department.repository import create_department


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


################################
# Repository tests
################################
def test_create_team(session):
    """Test creating a team and verifying its fields."""

    dept_data = DepartmentCreate(name="Cardiology", code="CARD")
    new_dept = create_department(session, dept_data)

    team_data = TeamCreate(name="ER Team A", department_id=new_dept.id)
    new_team = create_team(session, team_data)

    assert isinstance(new_team.id, uuid.UUID)
    assert new_team.name == "ER Team A"
    assert new_team.department_id == new_dept.id
    assert new_team.is_deleted is False
    assert new_team.sync_status is False
    assert isinstance(new_team.created_at, datetime.datetime)
    assert isinstance(new_team.updated_at, datetime.datetime)


def test_get_team_by_name(session):
    """ Tests retrieving a team by its name."""
    dept_data = DepartmentCreate(name="Neurology", code="NEURO")
    new_dept = create_department(session, dept_data)

    team_data = TeamCreate(name="Neuro Team B", department_id=new_dept.id)
    new_team = create_team(session, team_data)

    retrieved_team = get_team_by_name(session, "Neuro Team B")

    assert retrieved_team is not None
    assert retrieved_team.id == new_team.id


def test_get_active_teams(session):
    """ Tests retrieving only active (non-deleted) teams."""
    dept_data = DepartmentCreate(name="Oncology", code="ONC")
    new_dept = create_department(session, dept_data)

    team1_data = TeamCreate(name="Onco Team C", department_id=new_dept.id)
    team1 = create_team(session, team1_data)

    team2_data = TeamCreate(name="Onco Team D", department_id=new_dept.id)
    team2 = create_team(session, team2_data)

    # Mark team2 as deleted
    team2.is_deleted = True
    session.add(team2)
    session.commit()

    active_teams = get_active_teams(session)

    assert team1 in active_teams
    assert team2 not in active_teams


#######################
# Controller tests
#######################
def test_create_team_controller_duplicate_name(session):
    """Test that creating a team with a duplicate name raises an error."""
    dept_data = DepartmentCreate(name="Radiology", code="RAD")
    new_dept = create_department(session, dept_data)

    team_data = TeamCreate(name="Rad Team E", department_id=new_dept.id)

    create_team(session, team_data)

    with pytest.raises(Exception) as exc_info:
        from src.team.controllers import create_team_controller
        create_team_controller(team_data, session)

    assert exc_info.type.__name__ == "HTTPException"
    assert exc_info.value.status_code == 400
    assert "already exists" in exc_info.value.detail


def test_get_team_controller_nonexistent(session):
    """Test that fetching a non-existent team raises a 404 error"""
    non_existent_id = uuid.uuid4()

    with pytest.raises(Exception) as exc_info:
        from src.team.controllers import get_team_controller
        get_team_controller(non_existent_id, session)

    assert exc_info.type.__name__ == "HTTPException"
    assert exc_info.value.status_code == 404
    assert "not found" in exc_info.value.detail


def test_get_team_controller_deleted(session):
    """Test that fetching a deleted team raises a 404 error"""

    dept_data = DepartmentCreate(name="Radiology", code="RAD")
    new_dept = create_department(session, dept_data)

    team_data = TeamCreate(name="Rad Team E", department_id=new_dept.id)
    new_team = create_team(session, team_data)

    new_team.is_deleted = True
    session.add(new_team)
    session.commit()

    with pytest.raises(Exception) as exc_info:
        from src.team.controllers import get_team_controller
        get_team_controller(new_team.id, session)

    assert exc_info.type.__name__ == "HTTPException"
    assert exc_info.value.status_code == 404
    assert "not found" in exc_info.value.detail

###############################
# Route tests
###############################


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


def test_create_team_route(session, client):
    """Test the POST /teams/ route for creating a team."""
    dept_data = DepartmentCreate(name="Radiology", code="RAD")
    new_dept = create_department(session, dept_data)

    response = client.post(
        "/api/v1/teams", json={"name": "Rad Team E", "department_id": str(new_dept.id)}
    )

    assert response.status_code == 201
    data = response.json()
    assert data["name"] == "Rad Team E"
    assert data["department_id"] == str(new_dept.id)
    assert "id" in data
    assert "created_at" in data
    assert "updated_at" in data


def test_create_team_route_invalid_payload(client):
    """Tests that the POST /teams/ route rejects invalid payloads."""

    response = client.post(
        "/api/v1/teams/", json={"name": "Team A"}
    )
    assert response.status_code == 422


def test_get_team_by_id_route(session, client):
    """Tests that the GET /teams/{team_id} route returns a team"""
    dept_data = DepartmentCreate(name="Radiology", code="RAD")
    new_dept = create_department(session, dept_data)

    team_data = TeamCreate(name="Rad Team E", department_id=new_dept.id)
    new_team = create_team(session, team_data)

    response = client.get(
        f"/api/v1/teams/{new_team.id}"
    )

    assert response.status_code == 200
    data = response.json()
    assert data["name"] == "Rad Team E"
    assert data["department_id"] == str(new_dept.id)
    assert "id" in data
    assert "created_at" in data
    assert "updated_at" in data
