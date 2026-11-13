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
from src.shift import repository as shift_repository
from src.shift.schemas import ShiftAssignmentCreate, ShiftCreate
from src.doctor.schemas import DoctorPositionCreate, DoctorPreAssignmentCreate
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


@pytest.fixture(name="position_factory")
def position_factory_fixture(session, department):
    """Factory fixture to create positions with customizable fields and flags."""

    def _create_position(
        name: str = "Factory Position",
        duty_days: list[int] | None = None,
        department_id: uuid.UUID | None = None,
        is_deleted: bool = False,
        sync_status: bool = False,
    ) -> PositionModel:
        created = position_repository.create_position(
            session=session,
            position_name=name,
            duty_days=duty_days if duty_days is not None else [0, 2, 4],
            department_id=department_id or department.id,
        )
        if is_deleted or sync_status:
            created.is_deleted = is_deleted
            created.sync_status = sync_status
            session.add(created)
            session.commit()
            session.refresh(created)
        return created

    return _create_position


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


def test_create_position_controller_restores_soft_deleted_position(
    session,
    department,
    position_factory,
):
    """Tests creating a position with a soft-deleted position's name restores it and updates duty_days."""
    deleted_position = position_factory(
        name="Archived ICU",
        duty_days=[0, 1],
        department_id=department.id,
        is_deleted=True,
        sync_status=True,
    )
    previous_updated_at = deleted_position.updated_at

    restored = position_controllers.create_position_controller(
        position_data=PositionCreate(name="Archived ICU", duty_days=[5, 2, 2]),
        department_id=department.id,
        session=session,
    )

    assert restored.id == deleted_position.id
    assert restored.name == "Archived ICU"
    assert restored.duty_days == [2, 5]
    assert restored.is_deleted is False
    assert restored.sync_status is False
    assert restored.updated_at >= previous_updated_at


@pytest.mark.parametrize(
    ("update_kwargs", "expected_name", "expected_duty_days"),
    [
        pytest.param({"name": "Updated ICU"}, "Updated ICU", [1, 3, 5], id="rename-only"),
        pytest.param({"duty_days": [6, 0, 6]}, "Initial ICU", [0, 6], id="duty-days-only"),
        pytest.param({"name": "Trauma", "duty_days": [4, 2]}, "Trauma", [2, 4], id="both"),
    ],
)
def test_update_position_controller_updates_fields(
    session,
    position_factory,
    update_kwargs,
    expected_name,
    expected_duty_days,
):
    """Tests partial and full updates via update_position_controller."""
    target_position = position_factory(
        name="Initial ICU",
        duty_days=[1, 3, 5],
        sync_status=True,
    )
    previous_updated_at = target_position.updated_at

    updated = position_controllers.update_position_controller(
        position_id=target_position.id,
        position_data=PositionUpdate(**update_kwargs),
        department_id=target_position.department_id,
        session=session,
    )

    assert updated.id == target_position.id
    assert updated.name == expected_name
    assert updated.duty_days == expected_duty_days
    assert updated.sync_status is False
    assert updated.updated_at >= previous_updated_at


@pytest.mark.parametrize("other_is_deleted", [False, True])
def test_update_position_controller_rejects_colliding_name(
    session,
    position_factory,
    other_is_deleted,
):
    """Tests renaming a position to an existing (active or soft-deleted) name raises 400."""
    pos_a = position_factory(name="Position Alpha")
    pos_b = position_factory(name="Position Beta", is_deleted=other_is_deleted)

    with pytest.raises(HTTPException) as exc_info:
        position_controllers.update_position_controller(
            position_id=pos_a.id,
            position_data=PositionUpdate(name=pos_b.name),
            department_id=pos_a.department_id,
            session=session,
        )

    assert exc_info.value.status_code == 400
    assert "already exists" in exc_info.value.detail


def test_update_position_controller_nonexistent_or_deleted_raises_404(
    session,
    department,
    position_factory,
):
    """Tests updating a missing or soft-deleted position raises 404."""
    deleted_position = position_factory(name="Deleted Pos", is_deleted=True)

    for missing_id in (uuid.uuid4(), deleted_position.id):
        with pytest.raises(HTTPException) as exc_info:
            position_controllers.update_position_controller(
                position_id=missing_id,
                position_data=PositionUpdate(name="New Name"),
                department_id=department.id,
                session=session,
            )

        assert exc_info.value.status_code == 404
        assert "not found" in exc_info.value.detail.lower()


def test_delete_position_controller_cascades_and_preserves_history(
    session,
    department,
    doctor,
    position_factory,
):
    """Tests soft-deleting a position cascades to shifts, doctor links, and upcoming pre-assignments."""
    target_pos = position_factory(name="Cascade Target", sync_status=True)
    sibling_pos = position_factory(name="Sibling Position")

    shift_day = shift_repository.create_shift(
        session=session,
        shift_data=ShiftCreate(
            name="Day",
            position_id=target_pos.id,
            doctors_per_shift=1,
            grants_day_off=False,
        ),
    )
    shift_day.sync_status = True
    session.add(shift_day)
    session.commit()

    shift_night = shift_repository.create_shift(
        session=session,
        shift_data=ShiftCreate(
            name="Night",
            position_id=target_pos.id,
            doctors_per_shift=1,
            grants_day_off=True,
        ),
    )
    sibling_shift = shift_repository.create_shift(
        session=session,
        shift_data=ShiftCreate(
            name="Sibling Day",
            position_id=sibling_pos.id,
            doctors_per_shift=1,
            grants_day_off=False,
        ),
    )

    doc_pos_target = doctor_repository.create_doctor_position(
        session=session,
        doctor_id=doctor.id,
        doctor_pos_data=DoctorPositionCreate(position_id=target_pos.id),
    )
    doc_pos_target.sync_status = True
    session.add(doc_pos_target)
    session.commit()

    doc_pos_sibling = doctor_repository.create_doctor_position(
        session=session,
        doctor_id=doctor.id,
        doctor_pos_data=DoctorPositionCreate(position_id=sibling_pos.id),
    )

    today = datetime.date.today()
    upcoming_today_pre = doctor_repository.create_doctor_pre_assignment(
        session=session,
        doctor_id=doctor.id,
        pre_assignment_data=DoctorPreAssignmentCreate(
            shift_id=shift_day.id,
            date=today,
        ),
    )
    upcoming_today_pre.sync_status = True
    session.add(upcoming_today_pre)
    session.commit()

    upcoming_future_pre = doctor_repository.create_doctor_pre_assignment(
        session=session,
        doctor_id=doctor.id,
        pre_assignment_data=DoctorPreAssignmentCreate(
            shift_id=shift_night.id,
            date=today + datetime.timedelta(days=5),
        ),
    )
    historical_pre = doctor_repository.create_doctor_pre_assignment(
        session=session,
        doctor_id=doctor.id,
        pre_assignment_data=DoctorPreAssignmentCreate(
            shift_id=shift_day.id,
            date=today - datetime.timedelta(days=1),
        ),
    )
    sibling_pre = doctor_repository.create_doctor_pre_assignment(
        session=session,
        doctor_id=doctor.id,
        pre_assignment_data=DoctorPreAssignmentCreate(
            shift_id=sibling_shift.id,
            date=today + datetime.timedelta(days=7),
        ),
    )

    future_assignment = shift_repository.create_shift_assignment(
        session=session,
        shift_id=shift_day.id,
        shift_assignment_data=ShiftAssignmentCreate(
            doctor_id=doctor.id,
            date=today + datetime.timedelta(days=2),
        ),
    )
    past_assignment = shift_repository.create_shift_assignment(
        session=session,
        shift_id=shift_night.id,
        shift_assignment_data=ShiftAssignmentCreate(
            doctor_id=doctor.id,
            date=today - datetime.timedelta(days=3),
        ),
    )

    position_controllers.delete_position_controller(
        position_id=target_pos.id,
        department_id=department.id,
        session=session,
    )

    for obj in (
        target_pos,
        shift_day,
        shift_night,
        doc_pos_target,
        upcoming_today_pre,
        upcoming_future_pre,
        historical_pre,
        future_assignment,
        past_assignment,
        sibling_pos,
        sibling_shift,
        doc_pos_sibling,
        sibling_pre,
    ):
        session.refresh(obj)

    assert target_pos.is_deleted is True
    assert target_pos.sync_status is False
    assert shift_day.is_deleted is True
    assert shift_day.sync_status is False
    assert shift_night.is_deleted is True
    assert doc_pos_target.is_deleted is True
    assert doc_pos_target.sync_status is False
    assert upcoming_today_pre.is_deleted is True
    assert upcoming_today_pre.sync_status is False
    assert upcoming_future_pre.is_deleted is True

    # Historical pre-assignments, all ShiftAssignments, and sibling entities stay active
    assert historical_pre.is_deleted is False
    assert future_assignment.is_deleted is False
    assert past_assignment.is_deleted is False
    assert sibling_pos.is_deleted is False
    assert sibling_shift.is_deleted is False
    assert doc_pos_sibling.is_deleted is False
    assert sibling_pre.is_deleted is False


def test_delete_position_controller_nonexistent_or_deleted_raises_404(
    session,
    department,
    position_factory,
):
    """Tests deleting a missing or already soft-deleted position raises 404."""
    deleted_position = position_factory(name="Already Gone", is_deleted=True)

    for missing_id in (uuid.uuid4(), deleted_position.id):
        with pytest.raises(HTTPException) as exc_info:
            position_controllers.delete_position_controller(
                position_id=missing_id,
                department_id=department.id,
                session=session,
            )

        assert exc_info.value.status_code == 404
        assert "not found" in exc_info.value.detail.lower()

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


def test_position_update_schema_rejects_department_id():
    """Tests PositionUpdate forbids client-supplied department_id."""
    with pytest.raises(ValidationError):
        PositionUpdate(name="ER", department_id=uuid.uuid4())


def test_restore_soft_deleted_position_route(
    client,
    session,
    department,
    position_factory,
    department_admin_headers,
):
    """Tests POST /positions/ restores a soft-deleted position and updates duty_days."""
    deleted_pos = position_factory(
        name="Restorable ER",
        duty_days=[0, 1],
        department_id=department.id,
        is_deleted=True,
        sync_status=True,
    )

    response = client.post(
        "/api/v1/positions/",
        json={"name": "Restorable ER", "duty_days": [6, 4, 4]},
        headers=department_admin_headers,
    )

    assert response.status_code == 201
    data = response.json()
    assert data["id"] == str(deleted_pos.id)
    assert data["name"] == "Restorable ER"
    assert data["duty_days"] == [4, 6]
    assert data["is_deleted"] is False
    assert data["sync_status"] is False

    session.refresh(deleted_pos)
    assert deleted_pos.is_deleted is False
    assert deleted_pos.duty_days == [4, 6]


@pytest.mark.parametrize(
    ("payload", "expected_name", "expected_duty_days"),
    [
        pytest.param({"name": "Renamed ER"}, "Renamed ER", [1, 3, 5], id="rename-only"),
        pytest.param({"duty_days": [6, 2, 2]}, "ER", [2, 6], id="duty-days-only"),
        pytest.param({"name": "Trauma ER", "duty_days": [0, 4]}, "Trauma ER", [0, 4], id="both"),
        pytest.param({"name": "ER", "duty_days": [1, 3, 5]}, "ER", [1, 3, 5], id="same-values"),
        pytest.param({}, "ER", [1, 3, 5], id="empty-body"),
    ],
)
def test_update_position_route(
    client,
    session,
    position,
    department_admin_headers,
    payload,
    expected_name,
    expected_duty_days,
):
    """Tests PATCH /positions/{position_id} supports partial, full, and idempotent updates."""
    response = client.patch(
        f"/api/v1/positions/{position.id}",
        json=payload,
        headers=department_admin_headers,
    )

    assert response.status_code == 200
    data = response.json()
    assert data["id"] == str(position.id)
    assert data["name"] == expected_name
    assert data["duty_days"] == expected_duty_days
    assert data["department_id"] == str(position.department_id)
    assert data["sync_status"] is False

    session.refresh(position)
    assert position.name == expected_name
    assert position.duty_days == expected_duty_days


def test_update_position_route_rejects_supplied_department_id(
    client,
    session,
    position,
    department_b,
    department_admin_headers,
):
    """Tests PATCH /positions/{position_id} rejects department_id in payload (extra='forbid')."""
    response = client.patch(
        f"/api/v1/positions/{position.id}",
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

    session.refresh(position)
    assert position.name == "ER"
    assert position.department_id != department_b.id


@pytest.mark.parametrize("bad_duty_days", [[], [7], [-1, 2]])
def test_update_position_route_rejects_invalid_duty_days(
    client,
    position,
    department_admin_headers,
    bad_duty_days,
):
    """Tests PATCH /positions/{position_id} returns 422 for invalid duty_days."""
    response = client.patch(
        f"/api/v1/positions/{position.id}",
        json={"duty_days": bad_duty_days},
        headers=department_admin_headers,
    )

    assert response.status_code == 422


def test_update_position_route_rejects_duplicate_name(
    client,
    session,
    position,
    position_factory,
    department_admin_headers,
):
    """Tests PATCH /positions/{position_id} returns 400 when renaming to an existing position name."""
    other_position = position_factory(name="ICU")

    response = client.patch(
        f"/api/v1/positions/{position.id}",
        json={"name": other_position.name},
        headers=department_admin_headers,
    )

    assert response.status_code == 400
    assert "already exists" in response.json()["detail"]

    session.refresh(position)
    assert position.name == "ER"


def test_delete_position_route(
    client,
    session,
    position,
    doctor,
    department_admin_headers,
):
    """Tests DELETE /positions/{position_id} cascades soft-delete and hides the position from reads."""
    shift = shift_repository.create_shift(
        session=session,
        shift_data=ShiftCreate(
            name="Day",
            position_id=position.id,
            doctors_per_shift=1,
            grants_day_off=False,
        ),
    )
    doc_pos = doctor_repository.create_doctor_position(
        session=session,
        doctor_id=doctor.id,
        doctor_pos_data=DoctorPositionCreate(position_id=position.id),
    )
    today = datetime.date.today()
    upcoming_pre = doctor_repository.create_doctor_pre_assignment(
        session=session,
        doctor_id=doctor.id,
        pre_assignment_data=DoctorPreAssignmentCreate(
            shift_id=shift.id,
            date=today + datetime.timedelta(days=3),
        ),
    )
    assignment = shift_repository.create_shift_assignment(
        session=session,
        shift_id=shift.id,
        shift_assignment_data=ShiftAssignmentCreate(
            doctor_id=doctor.id,
            date=today + datetime.timedelta(days=4),
        ),
    )

    response = client.delete(
        f"/api/v1/positions/{position.id}",
        headers=department_admin_headers,
    )

    assert response.status_code == 204
    assert response.content == b""

    for obj in (position, shift, doc_pos, upcoming_pre, assignment):
        session.refresh(obj)

    assert position.is_deleted is True
    assert shift.is_deleted is True
    assert doc_pos.is_deleted is True
    assert upcoming_pre.is_deleted is True
    assert assignment.is_deleted is False

    list_response = client.get("/api/v1/positions/", headers=department_admin_headers)
    assert list_response.status_code == 200
    assert all(item["id"] != str(position.id) for item in list_response.json())

    get_response = client.get(
        f"/api/v1/positions/{position.id}",
        headers=department_admin_headers,
    )
    assert get_response.status_code == 404


@pytest.mark.parametrize("method", ["PATCH", "DELETE"])
@pytest.mark.parametrize("target_kind", ["missing", "deleted", "foreign_department"])
def test_position_mutation_routes_return_404_for_missing_deleted_or_foreign_position(
    client,
    position_factory,
    department_b,
    department_admin_headers,
    method,
    target_kind,
):
    """Tests PATCH/DELETE return 404 for nonexistent, soft-deleted, or foreign positions."""
    if target_kind == "missing":
        target_id = uuid.uuid4()
    elif target_kind == "deleted":
        target_id = position_factory(name="Deleted Target", is_deleted=True).id
    else:
        target_id = position_factory(
            name="Foreign Position",
            department_id=department_b.id,
        ).id

    request_kwargs = {"json": {"name": "New Name"}} if method == "PATCH" else {}
    response = client.request(
        method,
        f"/api/v1/positions/{target_id}",
        headers=department_admin_headers,
        **request_kwargs,
    )

    assert response.status_code == 404
    assert response.json() == {"detail": "Position not found"}
    assert response.headers.get("WWW-Authenticate") is None


@pytest.mark.parametrize("method", ["PATCH", "DELETE"])
def test_position_mutation_routes_require_authentication(client, position, method):
    """Tests PATCH/DELETE /positions/{position_id} require authentication."""
    request_kwargs = {"json": {"name": "No Auth"}} if method == "PATCH" else {}
    response = client.request(method, f"/api/v1/positions/{position.id}", **request_kwargs)

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
def test_non_department_admin_cannot_update_or_delete_position(
    client,
    session,
    department,
    position,
    doctor,
    user_factory,
    auth_headers_factory,
    method,
    role,
):
    """Tests that only DEPARTMENT_ADMIN can PATCH or DELETE positions."""
    user = user_factory(
        role=role,
        department_id=None if role == UserRole.SUPER_ADMIN else department.id,
        doctor_id=doctor.id if role == UserRole.DOCTOR else None,
    )
    headers = auth_headers_factory(user)
    request_kwargs = {"json": {"name": "Blocked Rename"}} if method == "PATCH" else {}

    response = client.request(
        method,
        f"/api/v1/positions/{position.id}",
        headers=headers,
        **request_kwargs,
    )

    assert response.status_code == 403
    assert response.json() == {"detail": "Insufficient permissions for this operation."}
    assert response.headers.get("WWW-Authenticate") is None

    session.refresh(position)
    assert position.name == "ER"
    assert position.is_deleted is False
