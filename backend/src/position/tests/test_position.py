"""
Tests for the position module
"""

import uuid
import datetime
import pytest

from src.position.schemas import PositionCreate
from src.position import repository as position_repository
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


@pytest.fixture(name="position")
def position_fixture(session, department):
    """Creates a reusable position for tests"""
    position_data = PositionCreate(
        name="ER",
        department_id=department.id,
        duty_days=[1, 3, 5],
    )

    return position_repository.create_position(session, position_data)


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
            "department_id": str(department.id),
            "duty_days": [1, 2, 3]
        },
        headers=department_admin_headers,
    )

    assert response.status_code == 201
    data = response.json()
    assert data["name"] == "ER"
    assert "duty_days" in data
    assert "id" in data
    assert "created_at" in data
    assert "updated_at" in data


def test_create_position_route_invalid_payload(client, department_admin_headers):
    """Tests post /api/v1/positions route with invalid payload"""
    response = client.post(
        "api/v1/positions",
        json={
            "name": "ER",
        },
        headers=department_admin_headers,
    )

    assert response.status_code == 422


def test_list_positions_route(client, position, viewer_headers):
    """Tests get /api/v1/positionsroute"""
    response = client.get(
        "/api/v1/positions",
        headers=viewer_headers,
    )

    assert response.status_code == 200
    data = response.json()
    assert isinstance(data, list)

    returned_position = next(
        (
            item
            for item in data
            if item["id"] == str(position.id)
        ),
        None
    )
    assert returned_position is not None
    assert returned_position["name"] == position.name


def test_get_position_route(client, position, viewer_headers):
    """Tests get /api/v1/positions/{position_name} route"""
    response = client.get(
        f"/api/v1/positions/{position.name}",
        headers=viewer_headers,
    )

    assert response.status_code == 200
    data = response.json()
    assert data["id"] == str(position.id)


def test_get_position_route_nonexistent_name(client, viewer_headers):
    """Tests get /api/v1/positions/{position_name} route with invalid payload"""
    response = client.get(
        "/api/v1/positions/test",
        headers=viewer_headers,
    )

    assert response.status_code == 404


@pytest.mark.parametrize(
    "role",
    [
        UserRole.DOCTOR,
        UserRole.VIEWER,
        UserRole.SUPER_ADMIN,
    ]
)
def test_non_allowed_roles_cannot_create_position(
    client,
    user_factory,
    auth_headers_factory,
    department,
    session,
    role,
):
    """Test that role except department_admin cannot create a position"""
    user = user_factory(
        role=role,
        department_id=department.id
    )
    headers = auth_headers_factory(user)

    response = client.post(
        "/api/v1/positions",
        json={
            "name": "ICU",
            "department_id": str(department.id),
            "duty_days": [1, 2],
        },
        headers=headers,
    )

    assert response.status_code == 403
    assert response.json() == {
        "detail": "Insufficient permissions for this operation"}
    assert response.headers.get("WWW-Authenticate") is None
    assert position_repository.get_position_by_name(session, "ICU") is None
