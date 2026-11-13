"""
Tests for the team module
"""

import uuid
import datetime
import pytest
from fastapi import HTTPException
from sqlalchemy.exc import IntegrityError

from src.team.schemas import TeamCreate, TeamUpdate
from src.team import repository as team_repository
from src.team import controllers as team_controllers
from src.department.schemas import DepartmentCreate
from src.department import repository as department_repository
from src.doctor import repository as doctor_repository
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


@pytest.fixture(name="doctor")
def doctor_fixture(session, department, team):
    """Creates a reusable doctor for tests"""
    return doctor_repository.create_doctor(
        session=session,
        department_id=department.id,
        name="Dr Panos",
        team_id=team.id,
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


@pytest.fixture(name="department_b")
def department_b_fixture(session):
    """Creates a secondary department for cross-tenant isolation tests."""
    return department_repository.create_department(
        session=session,
        department_data=DepartmentCreate(name="Department B", code="DEPT B"),
    )


@pytest.fixture(name="team_factory")
def team_factory_fixture(session, department):
    """Factory fixture to create teams with customizable status flags."""

    def _create_team(
        name: str = "Factory Team",
        department_id: uuid.UUID | None = None,
        is_deleted: bool = False,
        sync_status: bool = False,
    ):
        created = team_repository.create_team(
            session=session,
            name=name,
            department_id=department_id or department.id,
        )
        if is_deleted or sync_status:
            created.is_deleted = is_deleted
            created.sync_status = sync_status
            session.add(created)
            session.commit()
            session.refresh(created)
        return created

    return _create_team


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
        department_id=department.id,
    )

    team2 = team_repository.create_team(
        session=session,
        name="Onco team D",
        department_id=department.id,
    )

    department_b = department_repository.create_department(
        session=session,
        department_data=DepartmentCreate(
            name="Department B",
            code="DEPT B",
        )
    )

    foreign_team = team_repository.create_team(
        session=session,
        name=team1.name,
        department_id=department_b.id,
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
    assert foreign_team not in active_teams
    assert all(
        team.department_id == department.id
        for team in active_teams
    )


def test_get_team_by_id_for_department_returns_own_active_team(
    session,
    team,
):
    """Tests retrieving an active team from its department."""
    retrieved_team = team_repository.get_team_by_id_for_department(
        session=session,
        team_id=team.id,
        department_id=team.department_id,
    )

    assert retrieved_team == team


def test_get_team_by_id_for_department_hides_foreign_team(
    session,
    team
):
    """Tests that a department-scoped lookup hides foreign teams."""
    department_b = department_repository.create_department(
        session=session,
        department_data=DepartmentCreate(
            name="Department B",
            code="DEPT B",
        )
    )

    foreign_team = team_repository.create_team(
        session=session,
        name=team.name,
        department_id=department_b.id,
    )

    retrieved_team = team_repository.get_team_by_id_for_department(
        session=session,
        team_id=foreign_team.id,
        department_id=team.department_id,
    )

    assert retrieved_team is None


def test_get_team_by_id_for_department_hides_deleted_team(
    session,
    team,
):
    """Tests tjat a department-scoped lookup hides deleted teams."""
    team.is_deleted = True
    session.add(team)
    session.commit()

    retrieved_team = team_repository.get_team_by_id_for_department(
        session=session,
        team_id=team.id,
        department_id=team.department_id,
    )

    assert retrieved_team is None


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


def test_create_team_controller_duplicate_name(session, team):
    """Test that creating a team with a duplicate name raises an error."""

    with pytest.raises(HTTPException) as exc_info:
        team_controllers.create_team_controller(
            team_data=TeamCreate(name=team.name),
            department_id=team.department_id,
            session=session,
        )

    assert exc_info.value.status_code == 400
    assert "already exists" in exc_info.value.detail


def test_get_team_controller_nonexistent(session, department):
    """Test that fetching a non-existent team raises a 404 error"""
    non_existent_id = uuid.uuid4()

    with pytest.raises(HTTPException) as exc_info:
        team_controllers.get_team_controller(
            session=session,
            team_id=non_existent_id,
            department_id=department.id,
        )

    assert exc_info.value.status_code == 404
    assert "not found" in exc_info.value.detail


def test_get_team_controller_deleted(session, team):
    """Test that fetching a deleted team raises a 404 error"""

    team.is_deleted = True
    session.add(team)
    session.commit()

    with pytest.raises(HTTPException) as exc_info:
        team_controllers.get_team_controller(
            session=session,
            team_id=team.id,
            department_id=team.department_id,
        )

    assert exc_info.value.status_code == 404
    assert "not found" in exc_info.value.detail


def test_create_team_controller_restores_soft_deleted_team(
    session,
    department,
    team_factory,
):
    """Test creating a team with a soft-deleted team's name restores it in place."""
    deleted_team = team_factory(
        name="Archived ER Team",
        department_id=department.id,
        is_deleted=True,
        sync_status=True,
    )
    previous_updated_at = deleted_team.updated_at

    restored = team_controllers.create_team_controller(
        team_data=TeamCreate(name="Archived ER Team"),
        department_id=department.id,
        session=session,
    )

    assert restored.id == deleted_team.id
    assert restored.name == "Archived ER Team"
    assert restored.is_deleted is False
    assert restored.sync_status is False
    assert restored.updated_at >= previous_updated_at


def test_update_team_controller_renames_team(session, team_factory):
    """Test updating a team name updates persisted fields and sync metadata."""
    target_team = team_factory(name="Before Rename", sync_status=True)
    previous_updated_at = target_team.updated_at

    updated = team_controllers.update_team_controller(
        team_id=target_team.id,
        team_data=TeamUpdate(name="After Rename"),
        department_id=target_team.department_id,
        session=session,
    )

    assert updated.id == target_team.id
    assert updated.name == "After Rename"
    assert updated.sync_status is False
    assert updated.updated_at >= previous_updated_at


@pytest.mark.parametrize("other_is_deleted", [False, True])
def test_update_team_controller_rejects_colliding_name(
    session,
    team_factory,
    other_is_deleted,
):
    """Test renaming a team to an existing (active or soft-deleted) name raises 400."""
    team_a = team_factory(name="Team Alpha")
    team_b = team_factory(name="Team Beta", is_deleted=other_is_deleted)

    with pytest.raises(HTTPException) as exc_info:
        team_controllers.update_team_controller(
            team_id=team_a.id,
            team_data=TeamUpdate(name=team_b.name),
            department_id=team_a.department_id,
            session=session,
        )

    assert exc_info.value.status_code == 400
    assert "already exists" in exc_info.value.detail


def test_update_team_controller_nonexistent_or_deleted_raises_404(
    session,
    department,
    team_factory,
):
    """Test updating a missing or soft-deleted team raises 404."""
    deleted_team = team_factory(name="Deleted Team", is_deleted=True)

    for missing_id in (uuid.uuid4(), deleted_team.id):
        with pytest.raises(HTTPException) as exc_info:
            team_controllers.update_team_controller(
                team_id=missing_id,
                team_data=TeamUpdate(name="New Name"),
                department_id=department.id,
                session=session,
            )

        assert exc_info.value.status_code == 404
        assert "not found" in exc_info.value.detail.lower()


def test_delete_team_controller_soft_deletes_and_detaches_doctors(
    session,
    department,
    team_factory,
):
    """Test soft-deleting a team detaches assigned doctors while keeping other teams intact."""
    target_team = team_factory(name="Team To Delete", sync_status=True)
    other_team = team_factory(name="Unaffected Team")

    affected_doctor = doctor_repository.create_doctor(
        session=session,
        department_id=department.id,
        name="Dr Affected",
        team_id=target_team.id,
    )
    affected_doctor.sync_status = True
    session.add(affected_doctor)
    session.commit()

    unaffected_doctor = doctor_repository.create_doctor(
        session=session,
        department_id=department.id,
        name="Dr Unaffected",
        team_id=other_team.id,
    )

    team_controllers.delete_team_controller(
        team_id=target_team.id,
        department_id=department.id,
        session=session,
    )

    session.refresh(target_team)
    session.refresh(affected_doctor)
    session.refresh(unaffected_doctor)

    assert target_team.is_deleted is True
    assert target_team.sync_status is False
    assert affected_doctor.team_id is None
    assert affected_doctor.is_deleted is False
    assert affected_doctor.sync_status is False
    assert unaffected_doctor.team_id == other_team.id


def test_delete_team_controller_nonexistent_or_deleted_raises_404(
    session,
    department,
    team_factory,
):
    """Test deleting a missing or already soft-deleted team raises 404."""
    deleted_team = team_factory(name="Already Deleted", is_deleted=True)

    for missing_id in (uuid.uuid4(), deleted_team.id):
        with pytest.raises(HTTPException) as exc_info:
            team_controllers.delete_team_controller(
                team_id=missing_id,
                department_id=department.id,
                session=session,
            )

        assert exc_info.value.status_code == 404
        assert "not found" in exc_info.value.detail.lower()


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


def test_create_team_rejects_duplicate_name_within_department(
    client,
    session,
    team,
    department_admin_headers,
):
    """Tests that an admin cannot dupliacte a team name in their department"""
    response = client.post(
        "/api/v1/teams",
        json={"name": team.name},
        headers=department_admin_headers,
    )

    assert response.status_code == 400
    assert response.json() == {
        "detail": f"A team named '{team.name}' already exists."
    }

    stored_teams = team_repository.get_active_teams_by_department(
        session=session,
        department_id=team.department_id,
    )

    assert {stored_team.id for stored_team in stored_teams} == {team.id}


def test_create_team_allows_same_name_in_another_department(
    client,
    session,
    team,
    user_factory,
    auth_headers_factory,
):
    """Tests that another department may use the same team name."""
    department_b = department_repository.create_department(
        session=session,
        department_data=DepartmentCreate(
            name="Department B",
            code="DEPT B",
        )
    )

    department_b_admin = user_factory(
        role=UserRole.DEPARTMENT_ADMIN,
        department_id=department_b.id
    )

    response = client.post(
        "/api/v1/teams",
        json={"name": team.name},
        headers=auth_headers_factory(department_b_admin),
    )

    assert response.status_code == 201
    data = response.json()

    assert data["name"] == team.name
    assert data["department_id"] == str(department_b.id)

    stored_team = team_repository.get_team_by_name_for_department(
        session=session,
        name=team.name,
        department_id=department_b.id,
    )

    assert stored_team is not None
    assert str(stored_team.id) == data["id"]


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
    role,
    doctor,
):
    """Tests that a doctor cannot perform an admin action"""
    user = user_factory(
        role=role,
        department_id=(
            None
            if role == UserRole.SUPER_ADMIN
            else department.id
        ),
        doctor_id=doctor.id if role == UserRole.DOCTOR else None
    )

    headers = auth_headers_factory(user)

    response = client.post(
        "/api/v1/teams",
        json={"name": "Rad Team E"},
        headers=headers
    )

    assert response.status_code == 403
    data = response.json()
    assert data == {"detail": "Insufficient permissions for this operation."}
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
        ),
    )
    doctor_b = None
    team_b = team_repository.create_team(
        session=session,
        name="Team B",
        department_id=department_b.id
    )

    if role == UserRole.DOCTOR:
        doctor_b = doctor_repository.create_doctor(
            session=session,
            name="Dr Dept B",
            team_id=team_b.id,
            department_id=department_b.id,
        )

    user = user_factory(
        role=role,
        department_id=department_b.id,
        doctor_id=doctor_b.id if role == UserRole.DOCTOR else None,
    )

    headers = auth_headers_factory(user)

    response = client.get(
        f"/api/v1/teams/{team.id}",
        headers=headers,
    )

    assert response.status_code == 404
    assert response.json() == {"detail": "Team not found."}
    assert response.headers.get("WWW-Authenticate") is None


def test_restore_soft_deleted_team_route(
    client,
    session,
    department,
    team_factory,
    department_admin_headers,
):
    """Tests POST /teams/ restores a previously soft-deleted team with the same name."""
    deleted_team = team_factory(
        name="Restorable Team",
        department_id=department.id,
        is_deleted=True,
        sync_status=True,
    )

    response = client.post(
        "/api/v1/teams/",
        json={"name": "Restorable Team"},
        headers=department_admin_headers,
    )

    assert response.status_code == 201
    data = response.json()
    assert data["id"] == str(deleted_team.id)
    assert data["name"] == "Restorable Team"
    assert data["is_deleted"] is False
    assert data["sync_status"] is False

    session.refresh(deleted_team)
    assert deleted_team.is_deleted is False


def test_update_team_route(
    client,
    session,
    team_factory,
    department_admin_headers,
):
    """Tests PATCH /teams/{team_id} renames an active team in the admin's department."""
    target_team = team_factory(name="Original Name", sync_status=True)

    response = client.patch(
        f"/api/v1/teams/{target_team.id}",
        json={"name": "Renamed Team"},
        headers=department_admin_headers,
    )

    assert response.status_code == 200
    data = response.json()
    assert data["id"] == str(target_team.id)
    assert data["name"] == "Renamed Team"
    assert data["department_id"] == str(target_team.department_id)
    assert data["sync_status"] is False

    session.refresh(target_team)
    assert target_team.name == "Renamed Team"
    assert target_team.sync_status is False


@pytest.mark.parametrize("use_same_name", [True, False])
def test_update_team_route_idempotent_same_name_or_empty_payload(
    client,
    team,
    department_admin_headers,
    use_same_name,
):
    """Tests PATCH /teams/{team_id} succeeds when keeping the same name or omitting fields."""
    payload = {"name": team.name} if use_same_name else {}

    response = client.patch(
        f"/api/v1/teams/{team.id}",
        json=payload,
        headers=department_admin_headers,
    )

    assert response.status_code == 200
    assert response.json()["name"] == team.name


def test_update_team_route_rejects_supplied_department_id(
    client,
    session,
    team,
    department_b,
    department_admin_headers,
):
    """Tests PATCH /teams/{team_id} rejects department_id in payload (extra='forbid')."""
    response = client.patch(
        f"/api/v1/teams/{team.id}",
        json={
            "name": "Attempted Move",
            "department_id": str(department_b.id),
        },
        headers=department_admin_headers,
    )

    assert response.status_code == 422
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

    session.refresh(team)
    assert team.name == "ER Team A"
    assert team.department_id != department_b.id


def test_update_team_route_rejects_duplicate_name(
    client,
    session,
    team,
    team_factory,
    department_admin_headers,
):
    """Tests PATCH /teams/{team_id} returns 400 when renaming to another team's name."""
    other_team = team_factory(name="ICU Team B")

    response = client.patch(
        f"/api/v1/teams/{team.id}",
        json={"name": other_team.name},
        headers=department_admin_headers,
    )

    assert response.status_code == 400
    assert "already exists" in response.json()["detail"]

    session.refresh(team)
    assert team.name == "ER Team A"


def test_delete_team_route(
    client,
    session,
    team,
    doctor,
    department_admin_headers,
):
    """Tests DELETE /teams/{team_id} soft-deletes the team and detaches assigned doctors."""
    assert doctor.team_id == team.id

    response = client.delete(
        f"/api/v1/teams/{team.id}",
        headers=department_admin_headers,
    )

    assert response.status_code == 204
    assert response.content == b""

    session.refresh(team)
    session.refresh(doctor)
    assert team.is_deleted is True
    assert team.sync_status is False
    assert doctor.team_id is None
    assert doctor.is_deleted is False

    list_response = client.get("/api/v1/teams/", headers=department_admin_headers)
    assert list_response.status_code == 200
    assert all(item["id"] != str(team.id) for item in list_response.json())

    get_response = client.get(
        f"/api/v1/teams/{team.id}",
        headers=department_admin_headers,
    )
    assert get_response.status_code == 404


@pytest.mark.parametrize("method", ["PATCH", "DELETE"])
@pytest.mark.parametrize("target_kind", ["missing", "deleted", "foreign_department"])
def test_team_mutation_routes_return_404_for_missing_deleted_or_foreign_team(
    client,
    team_factory,
    department_b,
    department_admin_headers,
    method,
    target_kind,
):
    """Tests PATCH/DELETE return 404 for nonexistent, soft-deleted, or foreign teams."""
    if target_kind == "missing":
        target_id = uuid.uuid4()
    elif target_kind == "deleted":
        target_id = team_factory(name="Deleted Target", is_deleted=True).id
    else:
        target_id = team_factory(
            name="Foreign Team",
            department_id=department_b.id,
        ).id

    request_kwargs = {"json": {"name": "New Name"}} if method == "PATCH" else {}
    response = client.request(
        method,
        f"/api/v1/teams/{target_id}",
        headers=department_admin_headers,
        **request_kwargs,
    )

    assert response.status_code == 404
    assert response.json() == {"detail": "Team not found."}
    assert response.headers.get("WWW-Authenticate") is None


@pytest.mark.parametrize("method", ["PATCH", "DELETE"])
def test_team_mutation_routes_require_authentication(client, team, method):
    """Tests PATCH/DELETE /teams/{team_id} require authentication."""
    request_kwargs = {"json": {"name": "No Auth"}} if method == "PATCH" else {}
    response = client.request(method, f"/api/v1/teams/{team.id}", **request_kwargs)

    assert response.status_code == 401
    assert response.json() == {"detail": "Unauthorized"}
    assert response.headers.get("WWW-Authenticate") == "Bearer"


@pytest.mark.parametrize("method", ["PATCH", "DELETE"])
@pytest.mark.parametrize(
    "role",
    [
        UserRole.DOCTOR,
        UserRole.VIEWER,
        UserRole.SUPER_ADMIN,
    ],
)
def test_non_department_admin_cannot_update_or_delete_team(
    client,
    session,
    department,
    team,
    doctor,
    user_factory,
    auth_headers_factory,
    method,
    role,
):
    """Tests that only DEPARTMENT_ADMIN can PATCH or DELETE teams."""
    user = user_factory(
        role=role,
        department_id=None if role == UserRole.SUPER_ADMIN else department.id,
        doctor_id=doctor.id if role == UserRole.DOCTOR else None,
    )
    headers = auth_headers_factory(user)
    request_kwargs = {"json": {"name": "Blocked Rename"}} if method == "PATCH" else {}

    response = client.request(
        method,
        f"/api/v1/teams/{team.id}",
        headers=headers,
        **request_kwargs,
    )

    assert response.status_code == 403
    assert response.json() == {"detail": "Insufficient permissions for this operation."}
    assert response.headers.get("WWW-Authenticate") is None

    session.refresh(team)
    assert team.name == "ER Team A"
    assert team.is_deleted is False
