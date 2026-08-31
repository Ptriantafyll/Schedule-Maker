"""
Tests for the team module
"""

import uuid
import datetime
import pytest
from fastapi import HTTPException

from src.team.schemas import TeamCreate
from src.team import repository as team_repository
from src.team import controllers as team_controllers
from src.department.schemas import DepartmentCreate
from src.department.repository import create_department
from src.user.models import UserRole


@pytest.fixture(name="department")
def department_fixture(session):
    """Creates a reusable department for team tests."""
    dept_data = DepartmentCreate(name="Cardiology", code="CARD")
    return create_department(session, dept_data)


@pytest.fixture(name="team")
def team_fixture(session, department):
    """Creates a reusable team for tests"""

    team_data = TeamCreate(name="ER Team A", department_id=department.id)
    return team_repository.create_team(session, team_data)


@pytest.fixture(name="department_admin_user")
def department_admin_user_fixture(user_factory, department):
    """Creates a reusable department admin user for tests"""
    return user_factory(
        role=UserRole.DEPARTMENT_ADMIN,
        department_id=department.id,
        doctor_id=None
    )


@pytest.fixture(name="department_admin_headers")
def department_admin_headers_fixture(department_admin_user, auth_headers_factory):
    """Creates reusable department admin auth headers for tests"""

    return auth_headers_factory(department_admin_user)
################################
# Repository tests
################################


def test_create_team(department, team):
    """Test creating a team and verifying its fields."""

    assert isinstance(team.id, uuid.UUID)
    assert team.name == "ER Team A"
    assert team.department_id == department.id
    assert team.is_deleted is False
    assert team.sync_status is False
    assert isinstance(team.created_at, datetime.datetime)
    assert isinstance(team.updated_at, datetime.datetime)


def test_get_team_by_name(session, team):
    """ Tests retrieving a team by its name."""

    retrieved_team = team_repository.get_team_by_name(session, team.name)

    assert retrieved_team is not None
    assert retrieved_team.id == team.id


def test_get_active_teams(session, department):
    """ Tests retrieving only active (non-deleted) teams."""

    team1_data = TeamCreate(name="Onco Team C", department_id=department.id)
    team1 = team_repository.create_team(session, team1_data)

    team2_data = TeamCreate(name="Onco Team D", department_id=department.id)
    team2 = team_repository.create_team(session, team2_data)

    # Mark team2 as deleted
    team2.is_deleted = True
    session.add(team2)
    session.commit()

    active_teams = team_repository.get_active_teams(session)

    assert team1 in active_teams
    assert team2 not in active_teams


#######################
# Controller tests
#######################
def test_create_team_controller_duplicate_name(session, department, team):
    """Test that creating a team with a duplicate name raises an error."""

    team_data = TeamCreate(name=team.name, department_id=department.id)
    with pytest.raises(HTTPException) as exc_info:
        team_controllers.create_team_controller(team_data, session)

    assert exc_info.value.status_code == 400
    assert "already exists" in exc_info.value.detail


def test_get_team_controller_nonexistent(session):
    """Test that fetching a non-existent team raises a 404 error"""
    non_existent_id = uuid.uuid4()

    with pytest.raises(HTTPException) as exc_info:
        team_controllers.get_team_controller(non_existent_id, session)

    assert exc_info.value.status_code == 404
    assert "not found" in exc_info.value.detail


def test_get_team_controller_deleted(session, team):
    """Test that fetching a deleted team raises a 404 error"""

    team.is_deleted = True
    session.add(team)
    session.commit()

    with pytest.raises(HTTPException) as exc_info:
        team_controllers.get_team_controller(team.id, session)

    assert exc_info.value.status_code == 404
    assert "not found" in exc_info.value.detail

###############################
# Route tests
###############################


def test_create_team_route(client, department, department_admin_headers):
    """Test the POST /teams/ route for creating a team."""

    response = client.post(
        "/api/v1/teams",
        json={"name": "Rad Team E", "department_id": str(department.id)},
        headers=department_admin_headers,
    )

    assert response.status_code == 201
    data = response.json()
    assert data["name"] == "Rad Team E"
    assert data["department_id"] == str(department.id)
    assert "id" in data
    assert "created_at" in data
    assert "updated_at" in data


def test_create_team_route_invalid_payload(client, department_admin_headers):
    """Tests that the POST /teams/ route rejects invalid payloads."""

    response = client.post(
        "/api/v1/teams/",
        json={"name": "Team A"},
        headers=department_admin_headers,
    )
    assert response.status_code == 422


def test_get_team_by_id_route(client, department, team):
    """Tests that the GET /teams/{team_id} route returns a team"""
    response = client.get(
        f"/api/v1/teams/{team.id}"
    )

    assert response.status_code == 200
    data = response.json()
    assert data["name"] == team.name
    assert data["department_id"] == str(department.id)
    assert "id" in data
    assert "created_at" in data
    assert "updated_at" in data


def test_list_teams_route():
    """Tests the GET /teams/ route"""
    # TODO
