"""
Tests for the position module
"""

import uuid
import datetime
import pytest
from fastapi import HTTPException
from pydantic import ValidationError
from sqlalchemy.exc import IntegrityError

from src.position.models import Position as PositionModel
from src.position.schemas import PositionCreate, PositionUpdate
from src.position import repository as position_repository
from src.doctor import repository as doctor_repository
from src.team import repository as team_repository
from src.position import controllers as position_controllers
from src.department.schemas import DepartmentCreate
from src.department import repository as department_repository
from src.user.models import UserRole

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
    """Creates a second department for tenant-isolation tests."""
    dept_data = DepartmentCreate(name="Radiology", code="RAD")
    return department_repository.create_department(session, dept_data)


@pytest.fixture(name="position")
def position_fixture(session, department):
    """Creates a reusable position for tests"""
    return position_repository.create_position(
        session=session,
        position_name="ER",
        department_id=department.id,
        duty_days=[1, 3, 5],
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


@pytest.fixture(name="team")
def team_fixture(session, department):
    """Creates a reusable team for tests"""
    return team_repository.create_team(
        session=session,
        name="Test team",
        department_id=department.id
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


def test_stage_position_flushes_without_committing(session, department):
    """Verify stage_position creates a position queryable in-session that rolls back cleanly."""
    staged = position_repository.stage_position(
        session=session,
        position_name="Staged Position",
        duty_days=[1, 2, 3],
        department_id=department.id,
    )

    assert staged.id is not None
    assert staged.name == "Staged Position"
    assert staged.duty_days == [1, 2, 3]

    fetched = session.get(PositionModel, staged.id)
    assert fetched is not None

    session.rollback()
    assert session.get(PositionModel, staged.id) is None


def test_get_position_by_id_for_department_returns_own_active_position(
    session,
    position,
):
    """Tests retrieving an active Position from its department."""
    retrieved_position = position_repository.get_position_by_id_for_department(
        session=session,
        position_id=position.id,
        department_id=position.department_id,
    )

    assert retrieved_position is not None
    assert retrieved_position.id == position.id


def test_get_position_by_id_for_department_hides_foreign_position(
    session,
    department,
    department_b,
):
    """Tests that a scoped ID lookup hides foreign Positions."""
    foreign_position = position_repository.create_position(
        session=session,
        position_name="Radiology Position",
        department_id=department_b.id,
        duty_days=[1, 3, 5],
    )

    retrieved_position = position_repository.get_position_by_id_for_department(
        session=session,
        position_id=foreign_position.id,
        department_id=department.id,
    )

    assert retrieved_position is None


def test_get_position_by_id_for_department_hides_deleted_position(
    session,
    position,
):
    """Tests that a scoped ID lookup hides deleted Positions."""
    position.is_deleted = True
    session.add(position)
    session.commit()

    retrieved_position = position_repository.get_position_by_id_for_department(
        session=session,
        position_id=position.id,
        department_id=position.department_id,
    )

    assert retrieved_position is None


def test_get_position_by_name_for_department(
    session,
    department_b,
    position,
):
    """Tests resolving same-named Positions by department."""
    position_b = position_repository.create_position(
        session=session,
        position_name=position.name,
        department_id=department_b.id,
        duty_days=[2, 4, 6],
    )

    retrieved_position_a = position_repository.get_position_by_name_for_department(
        session=session,
        department_id=position.department_id,
        position_name=position.name,
    )
    retrieved_position_b = position_repository.get_position_by_name_for_department(
        session=session,
        department_id=position_b.department_id,
        position_name=position_b.name,
    )

    assert retrieved_position_a is not None
    assert retrieved_position_b is not None
    assert retrieved_position_a.id == position.id
    assert retrieved_position_b.id == position_b.id


def test_get_active_positions_for_department_excludes_foreign_and_deleted(
    session,
    department,
    department_b,
    position,
):
    """Tests listing only active Positions within department scope."""
    deleted_position = position_repository.create_position(
        session=session,
        position_name="Clinic",
        department_id=position.department_id,
        duty_days=[4, 5, 6],
    )
    foreign_position = position_repository.create_position(
        session=session,
        position_name=position.name,
        department_id=department_b.id,
        duty_days=[2, 4, 6],
    )

    deleted_position.is_deleted = True
    session.add(deleted_position)
    session.commit()

    retrieved_positions = position_repository.get_active_positions_for_department(
        session=session,
        department_id=department.id,
    )

    returned_ids = {
        retrieved_position.id
        for retrieved_position in retrieved_positions
    }
    assert position.id in returned_ids
    assert deleted_position.id not in returned_ids
    assert foreign_position.id not in returned_ids
    assert all(
        retrieved_position.department_id == department.id
        for retrieved_position in retrieved_positions
    )


def test_position_name_can_repeat_across_departments(
    session,
    department,
):
    """Tests that position names are unique only within a department."""
    department_b = department_repository.create_department(
        session=session,
        department_data=DepartmentCreate(
            name="Radiology",
            code="RAD",
        ),
    )

    position_name = "Shared Position"

    position_a = position_repository.create_position(
        session=session,
        position_name=position_name,
        department_id=department.id,
        duty_days=[1, 2, 3],
    )

    position_b = position_repository.create_position(
        session=session,
        position_name=position_name,
        department_id=department_b.id,
        duty_days=[4, 5, 6],
    )

    assert position_a.id != position_b.id
    assert position_a.department_id == department.id
    assert position_b.department_id == department_b.id
    assert position_a.name == position_b.name == position_name


def test_position_name_cannot_repeat_within_department(
    session,
    position,
):
    """Tests that a position name is unique within a department."""
    with pytest.raises(IntegrityError):
        position_repository.create_position(
            session=session,
            position_name=position.name,
            department_id=position.department_id,
            duty_days=position.duty_days,
        )

    session.rollback()


def test_soft_deleted_position_name_remains_reserved_within_department(
    session,
    position,
):
    """Tests that a soft deletion does not release a position name."""
    position.is_deleted = True
    session.add(position)
    session.commit()

    with pytest.raises(IntegrityError):
        position_repository.create_position(
            session=session,
            position_name=position.name,
            department_id=position.department_id,
            duty_days=position.duty_days,
        )

    session.rollback()
    session.refresh(position)

    assert position.is_deleted is True

#####################
# Controller tests
#####################


def test_create_position_controller_duplicate_name(session, position):
    """Tests that creating a position with a duplicate name returns error"""
    position_data = PositionCreate(
        name=position.name,
        duty_days=position.duty_days,
    )

    with pytest.raises(HTTPException) as exc_info:
        position_controllers.create_position_controller(
            position_data=position_data,
            department_id=position.department_id,
            session=session,
        )

    assert exc_info.value.status_code == 400
    assert "already exists" in exc_info.value.detail


def test_get_position_controller_nonexistent(session, position):
    """Tests that trying to retrieve a non existent position returns error"""
    with pytest.raises(HTTPException) as exc_info:
        position_controllers.get_position_controller(
            position_id=uuid.uuid4(),
            department_id=position.department_id,
            session=session,
        )

    assert exc_info.value.status_code == 404
    assert "not found" in exc_info.value.detail


def test_get_position_controller_deleted(session, position):
    """Tests that trying to retrieve a deleted position returns error"""
    position.is_deleted = True
    session.add(position)
    session.commit()

    with pytest.raises(HTTPException) as exc_info:
        position_controllers.get_position_controller(
            position_id=position.id,
            department_id=position.department_id,
            session=session,
        )

    assert exc_info.value.status_code == 404
    assert "not found" in exc_info.value.detail

#####################
# Route tests
#####################


def test_create_position_route(
    client,
    department,
    department_admin_headers,
):
    """Tests post /api/v1/positions route"""
    response = client.post(
        "api/v1/positions",
        json={
            "name": "ER",
            "duty_days": [1, 2, 3]
        },
        headers=department_admin_headers,
    )

    assert response.status_code == 201
    data = response.json()
    assert data["name"] == "ER"
    assert data["department_id"] == str(department.id)
    assert "duty_days" in data
    assert "id" in data
    assert "created_at" in data
    assert "updated_at" in data


def test_create_position_route_rejects_supplied_department_id(
    client,
    session,
    department,
    department_admin_headers,
):
    """Tests post /api/v1/positions route"""
    position_name = "Client scoped position"
    response = client.post(
        "api/v1/positions",
        json={
            "name": position_name,
            "duty_days": [1, 2, 3],
            "department_id": str(department.id),
        },
        headers=department_admin_headers,
    )

    assert response.status_code == 422
    validation_errors = response.json()["detail"]
    department_id_error = next(
        (
            error
            for error in validation_errors
            if error["loc"] == ["body", "department_id"]
        ),
        None,
    )

    assert department_id_error is not None
    assert department_id_error["type"] == "extra_forbidden"
    assert position_repository.get_position_by_name_for_department(
        session=session,
        position_name=position_name,
        department_id=department.id,
    ) is None


def test_create_position_route_missing_duty_days(client, department_admin_headers):
    """Tests post /api/v1/positions route with invalid payload"""
    response = client.post(
        "api/v1/positions",
        json={
            "name": "ER",
        },
        headers=department_admin_headers,
    )

    assert response.status_code == 422


def test_create_position_allows_same_name_in_another_department(
    client,
    session,
    department_b,
    position,
    user_factory,
    auth_headers_factory,
):
    """Tests that separate departments may use the same Position name."""
    department_b_admin = user_factory(
        role=UserRole.DEPARTMENT_ADMIN,
        department_id=department_b.id,
    )

    response = client.post(
        "/api/v1/positions/",
        json={
            "name": position.name,
            "duty_days": [2, 4, 6],
        },
        headers=auth_headers_factory(department_b_admin),
    )

    assert response.status_code == 201
    data = response.json()
    assert data["name"] == position.name
    assert data["department_id"] == str(department_b.id)

    stored_position = position_repository.get_position_by_name_for_department(
        session=session,
        position_name=position.name,
        department_id=department_b.id,
    )
    assert stored_position is not None
    assert str(stored_position.id) == data["id"]


def test_list_positions_route(
    client,
    session,
    department,
    department_b,
    position,
    viewer_headers,
):
    """Tests listing only Positions in the authenticated department."""
    foreign_position = position_repository.create_position(
        session=session,
        position_name=position.name,
        department_id=department_b.id,
        duty_days=[2, 4, 6],
    )

    response = client.get(
        "/api/v1/positions/",
        headers=viewer_headers,
    )

    assert response.status_code == 200
    data = response.json()
    returned_ids = {item["id"] for item in data}
    assert str(position.id) in returned_ids
    assert str(foreign_position.id) not in returned_ids
    assert all(
        item["department_id"] == str(department.id)
        for item in data
    )


def test_get_position_route(client, position, viewer_headers):
    """Tests get /api/v1/positions/{position_id} route"""
    response = client.get(
        f"/api/v1/positions/{position.id}",
        headers=viewer_headers,
    )

    assert response.status_code == 200
    data = response.json()
    assert data["id"] == str(position.id)


def test_get_position_route_malformed_id(client, viewer_headers):
    """Tests get /api/v1/positions/{position_id} route with invalid payload"""
    response = client.get(
        "/api/v1/positions/test",
        headers=viewer_headers,
    )

    assert response.status_code == 422


def test_get_position_route_nonexistent_id(client, viewer_headers):
    """Tests get /api/v1/positions/{position_id} route with nonexistent id"""
    response = client.get(
        f"/api/v1/positions/{uuid.uuid4()}",
        headers=viewer_headers,
    )

    assert response.status_code == 404


@pytest.mark.parametrize(
    "role",
    [
        UserRole.DEPARTMENT_ADMIN,
        UserRole.DOCTOR,
        UserRole.VIEWER,
    ],
)
def test_department_member_cannot_get_position_from_another_department(
    client,
    session,
    department,
    department_b,
    role,
    user_factory,
    auth_headers_factory,
    doctor,
):
    """Tests that foreign Position IDs are hidden from department members."""
    foreign_position = position_repository.create_position(
        session=session,
        position_name="Radiology Position",
        department_id=department_b.id,
        duty_days=[1, 3, 5],
    )
    department_user = user_factory(
        role=role,
        department_id=department.id,
        doctor_id=doctor.id if role == UserRole.DOCTOR else None
    )

    response = client.get(
        f"/api/v1/positions/{foreign_position.id}",
        headers=auth_headers_factory(department_user),
    )

    assert response.status_code == 404
    assert response.json() == {"detail": "Position not found"}
    assert response.headers.get("WWW-Authenticate") is None


@pytest.mark.parametrize(
    "role",
    [
        UserRole.DOCTOR,
        UserRole.VIEWER,
        UserRole.SUPER_ADMIN,
    ]
)
def test_non_department_admin_cannot_create_position(
    client,
    user_factory,
    auth_headers_factory,
    department,
    session,
    role,
    doctor
):
    """Test that role except department_admin cannot create a position"""
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
        "/api/v1/positions",
        json={
            "name": "ICU",
            "duty_days": [1, 2],
        },
        headers=headers,
    )

    assert response.status_code == 403
    assert response.json() == {
        "detail": "Insufficient permissions for this operation."}
    assert response.headers.get("WWW-Authenticate") is None
    assert position_repository.get_position_by_name_for_department(
        session=session,
        position_name="ICU",
        department_id=department.id,
    ) is None


def test_create_position_requires_authentication(client, department, session):
    """Tests post /api/v1/positions route requires auth"""

    response = client.post(
        "/api/v1/positions",
        json={
            "name": "ICU",
            "duty_days": [1, 2],
        },
    )

    assert response.status_code == 401
    assert response.json() == {"detail": "Unauthorized"}
    assert response.headers.get("WWW-Authenticate") == "Bearer"


@pytest.mark.parametrize(
    "path_template",
    [
        pytest.param(
            "/api/v1/positions/",
            id="list-positions",
        ),
        pytest.param(
            "/api/v1/positions/{position_id}",
            id="get-position",
        ),
    ],
)
def test_position_read_routes_require_authentication(
    client,
    position,
    path_template,
):
    """Tests that the position read routes require auth"""
    path = path_template.format(position_id=position.id)

    response = client.get(path)

    assert response.status_code == 401
    assert response.json() == {"detail": "Unauthorized"}
    assert response.headers.get("WWW-Authenticate") == "Bearer"


def test_position_create_schema_duty_days_validation():
    """Tests Pydantic validation for PositionCreate duty_days."""
    # Out of range (7 is invalid; valid is 0..6)
    with pytest.raises(ValidationError):
        PositionCreate(name="ER", duty_days=[0, 1, 7])

    # Negative number
    with pytest.raises(ValidationError):
        PositionCreate(name="ER", duty_days=[-1, 2])

    # Empty list
    with pytest.raises(ValidationError):
        PositionCreate(name="ER", duty_days=[])

    # Valid: deduplicates and sorts
    pos = PositionCreate(name="ER", duty_days=[5, 2, 2, 0])
    assert pos.duty_days == [0, 2, 5]


def test_position_update_schema_duty_days_validation():
    """Tests Pydantic validation for PositionUpdate duty_days."""
    # Out of range
    with pytest.raises(ValidationError):
        PositionUpdate(duty_days=[7])

    # Empty list
    with pytest.raises(ValidationError):
        PositionUpdate(duty_days=[])

    # None is allowed for partial update
    pos_update = PositionUpdate(duty_days=None)
    assert pos_update.duty_days is None

    # Valid: deduplicates and sorts
    pos_update_valid = PositionUpdate(duty_days=[6, 0, 0])
    assert pos_update_valid.duty_days == [0, 6]


def test_create_position_route_rejects_invalid_duty_days(client, department_admin_headers):
    """Tests that POST /api/v1/positions returns 422 for out-of-range or empty duty_days."""
    # Out-of-range day (7)
    response_out_of_range = client.post(
        "/api/v1/positions/",
        json={"name": "Out of Range", "duty_days": [1, 2, 7]},
        headers=department_admin_headers,
    )
    assert response_out_of_range.status_code == 422

    # Empty duty days
    response_empty = client.post(
        "/api/v1/positions/",
        json={"name": "Empty Days", "duty_days": []},
        headers=department_admin_headers,
    )
    assert response_empty.status_code == 422


def test_create_position_route_normalizes_duty_days(
    client, session, department, department_admin_headers
):
    """Tests that valid duty_days are sorted and deduplicated upon creation."""
    response = client.post(
        "/api/v1/positions/",
        json={"name": "Normalized Position", "duty_days": [5, 1, 1, 3]},
        headers=department_admin_headers,
    )
    assert response.status_code == 201
    data = response.json()
    assert data["duty_days"] == [1, 3, 5]
