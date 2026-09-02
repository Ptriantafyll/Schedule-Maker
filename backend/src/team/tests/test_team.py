"""
Tests for the team module
"""

import uuid
import datetime
import pytest
from fastapi import HTTPException
from sqlalchemy.exc import IntegrityError

from src.team.schemas import TeamCreate
from src.team import repository as team_repository
from src.team import controllers as team_controllers
from src.department.schemas import DepartmentCreate
from src.department import repository as department_repository
from src.user.models import UserRole


@pytest.fixture(name="department")
def department_fixture(session):
    """Creates a reusable department for team tests."""
    dept_data = DepartmentCreate(name="Cardiology", code="CARD")
    return department_repository.create_department(session, dept_data)


@pytest.fixture(name="team")
def team_fixture(session, department):
    """Creates a reusable team for tests"""

    return team_repository.create_team(
        session=session,
        name="ER Team A",
        department_id=department.id,
    )


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


@pytest.fixture(name="viewer_user")
def viewer_user_fixture(user_factory, department):
    """Creates a reusable viewer user for tests"""
    return user_factory(
        role=UserRole.VIEWER,
        department_id=department.id,
        doctor_id=None
    )


@pytest.fixture(name="viewer_headers")
def viewer_headers_fixture(viewer_user, auth_headers_factory):
    """Creates reusable viewer auth headers for tests"""

    return auth_headers_factory(viewer_user)


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


def test_get_team_by_name_for_department(
    session,
    team,
):
    """Tests retrieving a team by its name and department id."""
    department_b = department_repository.create_department(
        session=session,
        department_data=DepartmentCreate(
            name="Department B",
            code="DEPT B",
        )
    )

    same_name_team = team_repository.create_team(
        session=session,
        name=team.name,
        department_id=department_b.id,
    )

    retrieved_team = team_repository.get_team_by_name_for_department(
        session=session,
        name=team.name,
        department_id=team.department_id,
    )
    retrieved_team_b = team_repository.get_team_by_name_for_department(
        session=session,
        name=same_name_team.name,
        department_id=same_name_team.department_id,
    )

    assert retrieved_team is not None
    assert retrieved_team.id == team.id
    assert retrieved_team_b is not None
    assert retrieved_team_b.id == same_name_team.id


def test_get_active_teams_by_department(session, department):
    """Tests retrieving only active (non-deleted) teams."""

    team1 = team_repository.create_team(
        session=session,
        name="Onco Team C",
        department_id=department.id
    )

    team2 = team_repository.create_team(
        session=session,
        name="Onco team C",
        department_id=department.id,
    )

    # Mark team2 as deleted
    team2.is_deleted = True
    session.add(team2)
    session.commit()

    active_teams = team_repository.get_active_teams_by_department(
        session=session,
        department_id=department.id,
    )

    assert team1 in active_teams
    assert team2 not in active_teams


def test_team_name_can_repeat_across_departments(
    session,
    department,
):
    """Tests that team names are unique only within a department."""
    department_b = department_repository.create_department(
        session,
        DepartmentCreate(name="Radiology", code="RAD"),
    )
    team_name = "Shared team"

    team_a = team_repository.create_team(
        session,
        name=team_name,
        department_id=department.id,
    )

    team_b = team_repository.create_team(
        session,
        name=team_name,
        department_id=department_b.id,
    )

    assert team_a.id != team_b.id
    assert team_a.department_id == department.id
    assert team_b.department_id == department_b.id
    assert team_a.name == team_b.name == team_name


def test_team_name_cannot_repeat_within_department(
    session,
    team,
):
    """Tests that team name remains unique within a department."""
    with pytest.raises(IntegrityError):
        team_repository.create_team(
            session=session,
            name=team.name,
            department_id=team.department_id,
        )

    session.rollback()


def test_soft_deleted_team_name_remains_reserved_within_department(
    session,
    team,
):
    """Tests that soft deletion does not release a team name."""
    team.is_deleted = True
    session.add(team)
    session.commit()

    with pytest.raises(IntegrityError):
        team_repository.create_team(
            session=session,
            name=team.name,
            department_id=team.department_id,
        )

    session.rollback()
    session.refresh(team)

    assert team.is_deleted is True


#######################
# Controller tests
#######################


def test_create_team_controller_duplicate_name(session, department, team):
    """Test that creating a team with a duplicate name raises an error."""

    with pytest.raises(HTTPException) as exc_info:
        team_controllers.create_team_controller(
            team_data=TeamCreate(name=team.name),
            department_id=team.department_id,
            session=session,
        )

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
        json={
            "name": "Rad Team E",
        },
        headers=department_admin_headers,
    )

    assert response.status_code == 201
    data = response.json()
    assert data["name"] == "Rad Team E"
    assert data["department_id"] == str(department.id)
    assert "id" in data
    assert "created_at" in data
    assert "updated_at" in data


def test_create_team_route_rejects_supplied_department_id(
    client,
    department,
    department_admin_headers,
    session,
):
    """Tests that the POST /teams/ route rejects client-supplied department scope."""

    response = client.post(
        "/api/v1/teams/",
        json={
            "name": "Team A",
            "department_id": str(department.id),
        },
        headers=department_admin_headers,
    )
    assert response.status_code == 422
    assert team_repository.get_team_by_name_for_department(
        session=session,
        name="Team A",
        department_id=department.id
    ) is None

    validation_errors = response.json()["detail"]
    department_id_error = next(
        (
            error for error in validation_errors
            if error["loc"] == ["body", "department_id"]
        ),
        None,
    )

    assert department_id_error is not None
    assert department_id_error["type"] == "extra_forbidden"


def test_get_team_by_id_route(client, department, team, viewer_headers):
    """Tests that the GET /teams/{team_id} route returns a team"""
    response = client.get(
        f"/api/v1/teams/{team.id}",
        headers=viewer_headers,
    )

    assert response.status_code == 200
    data = response.json()
    assert data["name"] == team.name
    assert data["department_id"] == str(department.id)
    assert "id" in data
    assert "created_at" in data
    assert "updated_at" in data


def test_list_teams_route(
    client,
    session,
    department,
    team,
    viewer_headers,
):
    """Tests the GET /teams/ route"""
    department_b = department_repository.create_department(
        session=session,
        department_data=DepartmentCreate(
            name="Department B",
            code="DEPT B",
        )
    )
    department_b_team = team_repository.create_team(
        session=session,
        name=team.name,
        department_id=department_b.id,
    )

    response = client.get(
        "/api/v1/teams",
        headers=viewer_headers,
    )

    assert response.status_code == 200

    data = response.json()
    assert isinstance(data, list)

    returned_ids = {item["id"] for item in data}

    assert str(team.id) in returned_ids
    assert str(department_b_team.id) not in returned_ids
    assert all(
        item["department_id"] == str(department.id)
        for item in data
    )


@pytest.mark.parametrize(
    "role",
    [
        UserRole.DOCTOR,
        UserRole.VIEWER,
        UserRole.SUPER_ADMIN,
    ]
)
def test_non_department_admin_cannot_create_team(
    client,
    department,
    session,
    user_factory,
    auth_headers_factory,
    role
):
    """Tests that a doctor cannot perform an admin action"""
    user = user_factory(
        role=role,
        department_id=(
            None
            if role == UserRole.SUPER_ADMIN
            else department.id
        )
    )

    headers = auth_headers_factory(user)

    response = client.post(
        "/api/v1/teams",
        json={"name": "Rad Team E"},
        headers=headers
    )

    assert response.status_code == 403
    data = response.json()
    assert data == {"detail": "Insufficient permissions for this operation"}
    assert response.headers.get("WWW-Authenticate") is None
    assert team_repository.get_team_by_name_for_department(
        session=session,
        name="Rad Team E",
        department_id=department.id,
    ) is None


@pytest.mark.parametrize(
    "path_template",
    [
        pytest.param(
            "/api/v1/teams/",
            id="list-teams",
        ),
        pytest.param(
            "/api/v1/teams/{team_id}",
            id="get-team",
        ),
    ],
)
def test_team_read_routes_require_authentication(
    client,
    team,
    path_template,
):
    """Tests that the team read routes require auth"""
    path = path_template.format(team_id=team.id)

    response = client.get(path)

    assert response.status_code == 401
    assert response.json() == {"detail": "Unauthorized"}
    assert response.headers.get("WWW-Authenticate") == "Bearer"


@pytest.mark.parametrize(
    "role",
    [
        UserRole.VIEWER,
        UserRole.DOCTOR,
        UserRole.DEPARTMENT_ADMIN,
    ]
)
def test_department_member_cannot_get_team_from_another_department(
    client,
    session,
    user_factory,
    auth_headers_factory,
    role,
    team,
):
    """Tests that a department member cannot get a team from another department"""
    department_b = department_repository.create_department(
        session=session,
        department_data=DepartmentCreate(
            name="Department B",
            code="DEPT B",
        )
    )

    user = user_factory(
        role=role,
        department_id=department_b.id
    )

    headers = auth_headers_factory(user)

    response = client.get(
        f"/api/v1/teams/{team.id}",
        headers=headers,
    )

    assert response.status_code == 404
    assert response.json() == {"detail": "Team not found."}


@pytest.mark.parametrize(
    "role",
    [
        UserRole.VIEWER,
        UserRole.DOCTOR,
        UserRole.DEPARTMENT_ADMIN,
    ]
)
def test_department_member_can_get_team_from_their_department(
    client,
    session,
    team,
    user_factory,
    auth_headers_factory,
):
    """Tests that a user can retrieve a team from their department"""
    user = user_factory(
        role=role,
        department_id=department_b.id
    )
    headers = auth_headers_factory(user)

    response = client.get(
        f"/api/v1/teams/{team.id}",
        headers=headers,
    )

    assert response.status_code = 200
