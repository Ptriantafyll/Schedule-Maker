"""
Tests for the shift module
"""

import uuid
import datetime
import pytest
from sqlmodel import Session, select
from sqlalchemy.exc import IntegrityError

from src.shift.schemas import ShiftCreate, ShiftAssignmentCreate
from src.shift.models import Shift as ShiftModel
from src.shift.models import ShiftAssignment as ShiftAssignmentModel
from src.shift import repository as shift_repository
from src.shift import controllers as shift_controllers
from src.position.schemas import PositionCreate
from src.position import repository as position_repository
from src.doctor.models import Doctor as DoctorModel
from src.doctor.schemas import DoctorCreate, DoctorUnavailabilityCreate
from src.doctor import repository as doctor_repository
from src.department.schemas import DepartmentCreate
from src.department import repository as department_repository
from src.team import repository as team_repository
from src.user.models import UserRole

#####################
# Helpers
#####################


def create_new_doctor(
    session: Session,
    name: str,
    email: str,
    team_id: uuid.UUID
) -> DoctorModel:
    """Helper that creates a new doctor in the db"""
    return doctor_repository.create_doctor(
        session=session,
        name=name,
        email=email,
        team_id=team_id
    )


def create_new_shift(
    session: Session,
    name: str,
    position_id: uuid.UUID,
    grants_day_off: bool,
    doctors_per_shift: int
) -> ShiftModel:
    """Helper that creates a shift in the db"""

    shift_data = ShiftCreate(
        name=name,
        position_id=position_id,
        grants_day_off=grants_day_off,
        doctors_per_shift=doctors_per_shift
    )

    return shift_repository.create_shift(session, shift_data)


def create_new_shift_assignment(
    session: Session,
    doctor_id: uuid.UUID,
    shift_id: uuid.UUID,
    date: datetime.date
) -> ShiftAssignmentModel:
    """Helper that creates a shift assignment in the db"""
    shift_assignment_data = ShiftAssignmentCreate(
        doctor_id=doctor_id,
        date=date
    )

    return shift_repository.create_shift_assignment(
        session=session,
        shift_id=shift_id,
        shift_assignment_data=shift_assignment_data
    )


def create_test_unavailability(
    session: Session,
    doctor: DoctorModel,
    date: datetime.date
):
    """Helper that creates a new doctor pre assignment in the db"""
    unavailability_data = DoctorUnavailabilityCreate(
        date=date
    )
    return doctor_repository.create_doctor_unavailability(
        session=session,
        doctor_id=doctor.id,
        doctor_unavailability_data=unavailability_data
    )

#####################
# Fixtures
#####################


@pytest.fixture(name="shift")
def shift_fixture(session, position):
    """Creates a reusable shift for tests"""
    return create_new_shift(session, "ER 1", position.id, False, 2)


@pytest.fixture(name="department")
def department_fixture(session):
    """Creates a reusable department for tests"""
    dept_data = DepartmentCreate(name="Cardiology", code="CARD")
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


@pytest.fixture(name="team")
def team_fixture(session, department):
    """Creates a reusable team for tests"""
    return team_repository.create_team(
        session=session,
        name="ER Team A",
        department_id=department.id
    )


@pytest.fixture(name="new_doctor")
def doctor_fixture(session, department, team):
    """Creates a reusable doctor for tests"""
    return create_new_doctor(session, "Dr Panos", "drpanos@gmail.com", department.id, team.id)


@pytest.fixture(name="shift_assignment")
def shift_assignment_fixture(session, new_doctor, shift):
    """Creates a reusable shift assignment for tests"""
    return create_new_shift_assignment(session, new_doctor.id, shift.id, datetime.date(2026, 8, 12))


@pytest.fixture(name="unavailability")
def unavailability_fixture(session, new_doctor):
    """Creates a reusable pre-assignment for tests"""
    return create_test_unavailability(session, new_doctor, datetime.date(2026, 8, 13))


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
    """Creates a second department for tenant-isolation tests."""
    dept_data = DepartmentCreate(name="Radiology", code="RAD")
    return department_repository.create_department(session, dept_data)


@pytest.fixture(name="position_b")
def position_b_fixture(session, department_b):
    """Creates a reusable Department B position for tenant-isolation tests."""
    return position_repository.create_position(
        session=session,
        position_name="ICU",
        department_id=department_b.id,
        duty_days=[2, 4, 6],
    )


@pytest.fixture(name="team_b")
def team_b_fixture(session, department_b):
    """Creates a reusable Department B team for tenant-isolation tests."""
    return team_repository.create_team(
        session=session,
        name="ICU Team B",
        department_id=department_b.id
    )


@pytest.fixture(name="doctor_b")
def doctor_b_fixture(session, department_b, team_b):
    """Creates a reusable Department B doctor for tenant-isolation tests."""
    return create_new_doctor(
        session, "Dr Radiology", "drradiology@gmail.com", department_b.id, team_b.id
    )


@pytest.fixture(name="shift_b")
def shift_b_fixture(session, position_b):
    """Creates a reusable Department B shift for tenant-isolation tests."""
    return create_new_shift(session, "ICU 1", position_b.id, False, 2)


#####################
# Repository Tests
#####################


def test_create_shift(shift, position):
    """Tests creating a shift in the db"""
    assert isinstance(shift.id, uuid.UUID)
    assert shift.position_id == position.id
    assert shift.is_deleted is False
    assert shift.sync_status is False
    assert isinstance(shift.created_at, datetime.datetime)
    assert isinstance(shift.updated_at, datetime.datetime)


def test_get_shift_by_name_for_position_resolves_by_position(
    session,
    shift,
    position,
    position_b,
):
    """Tests that the same shift name resolves independently under each position."""
    shift_under_b = create_new_shift(
        session=session,
        name=shift.name,
        position_id=position_b.id,
        grants_day_off=False,
        doctors_per_shift=1,
    )

    retrieved_a = shift_repository.get_shift_by_name_for_position(
        session=session,
        position_id=position.id,
        shift_name=shift.name,
    )
    retrieved_b = shift_repository.get_shift_by_name_for_position(
        session=session,
        position_id=position_b.id,
        shift_name=shift.name,
    )

    assert retrieved_a is not None
    assert retrieved_b is not None
    assert retrieved_a.id == shift.id
    assert retrieved_b.id == shift_under_b.id
    assert retrieved_a.id != retrieved_b.id


def test_get_shift_by_name_for_position_includes_deleted_for_reservation(
    session,
    shift,
):
    """Tests that a soft-deleted shift is still found to reserve its name."""
    shift.is_deleted = True
    session.add(shift)
    session.commit()

    retrieved_shift = shift_repository.get_shift_by_name_for_position(
        session=session,
        position_id=shift.position_id,
        shift_name=shift.name,
    )

    assert retrieved_shift is not None
    assert retrieved_shift.id == shift.id
    assert retrieved_shift.is_deleted is True


def test_get_shift_by_id_for_department_returns_own_active_shift(
    session,
    shift,
    position,
):
    """Tests retrieving an active Shift scoped to its own department."""
    retrieved_shift = shift_repository.get_shift_by_id_for_department(
        session=session,
        shift_id=shift.id,
        department_id=position.department_id,
    )

    assert retrieved_shift is not None
    assert retrieved_shift.id == shift.id


def test_get_shift_by_id_for_department_hides_foreign_shift(
    session,
    department,
    shift_b,
):
    """Tests that a scoped ID lookup hides Shifts from another department."""
    retrieved_shift = shift_repository.get_shift_by_id_for_department(
        session=session,
        shift_id=shift_b.id,
        department_id=department.id,
    )

    assert retrieved_shift is None


def test_get_shift_by_id_for_department_hides_deleted_shift(
    session,
    shift,
    position,
):
    """Tests that a scoped ID lookup hides deleted Shifts."""
    shift.is_deleted = True
    session.add(shift)
    session.commit()

    retrieved_shift = shift_repository.get_shift_by_id_for_department(
        session=session,
        shift_id=shift.id,
        department_id=position.department_id,
    )

    assert retrieved_shift is None


def test_get_active_shifts_for_department_excludes_foreign_and_deleted(
    session,
    department,
    position,
    shift,
    shift_b,
):
    """Tests listing only active Shifts within department scope."""
    deleted_shift = create_new_shift(session, "ER 2", position.id, False, 1)
    deleted_shift.is_deleted = True
    session.add(deleted_shift)
    session.commit()

    retrieved_shifts = shift_repository.get_active_shifts_for_department(
        session=session,
        department_id=department.id,
    )

    returned_ids = {retrieved_shift.id for retrieved_shift in retrieved_shifts}
    assert shift.id in returned_ids
    assert deleted_shift.id not in returned_ids
    assert shift_b.id not in returned_ids


def test_create_shift_assignment(session, new_doctor, shift):
    """Tests creating a shift assignment in the db"""
    new_shift_assignment = create_new_shift_assignment(
        session, new_doctor.id, shift.id, datetime.date(2026, 8, 12))

    assert isinstance(new_shift_assignment.id, uuid.UUID)
    assert new_shift_assignment.shift_id == shift.id
    assert new_shift_assignment.is_deleted is False
    assert new_shift_assignment.sync_status is False
    assert isinstance(new_shift_assignment.created_at, datetime.datetime)
    assert isinstance(new_shift_assignment.updated_at, datetime.datetime)


def test_get_shift_assignments_by_date(session, shift_assignment):
    """Tests retrieving a shift assignemtnt by its date"""
    retrieved_shift_assignments = shift_repository.get_shift_assignments_by_date(
        session=session,
        shift_id=shift_assignment.shift_id,
        target_date=datetime.date(2026, 8, 12)
    )

    assert isinstance(retrieved_shift_assignments, list)
    assert shift_assignment.id == retrieved_shift_assignments[0].id


def test_get_shift_assignment_by_id(session, shift_assignment):
    """Tests retrieving a shift assignemtnt by its date"""
    retrieved_shift_assignment = shift_repository.get_shift_assignment_by_id(
        session=session,
        shift_assignment_id=shift_assignment.id
    )

    assert isinstance(retrieved_shift_assignment.id, uuid.UUID)
    assert shift_assignment.id == retrieved_shift_assignment.id


def test_get_active_shift_assignments_for_department_excludes_foreign_and_deleted(
    session,
    department,
    new_doctor,
    shift,
    shift_assignment,
    shift_b,
    doctor_b,
):
    """Tests listing only active ShiftAssignments within department scope."""
    deleted_assignment = create_new_shift_assignment(
        session=session,
        doctor_id=new_doctor.id,
        shift_id=shift.id,
        date=datetime.date(2026, 8, 13),
    )
    deleted_assignment.is_deleted = True
    session.add(deleted_assignment)
    session.commit()

    foreign_assignment = create_new_shift_assignment(
        session=session,
        doctor_id=doctor_b.id,
        shift_id=shift_b.id,
        date=datetime.date(2026, 8, 12),
    )

    retrieved_assignments = shift_repository.get_active_shift_assignments_for_department(
        session=session,
        department_id=department.id,
    )

    returned_ids = {
        retrieved_assignment.id
        for retrieved_assignment in retrieved_assignments
    }
    assert shift_assignment.id in returned_ids
    assert deleted_assignment.id not in returned_ids
    assert foreign_assignment.id not in returned_ids


def test_shift_name_can_repeat_across_positions(
    session,
    position,
):
    """Tests that a shift name is allowed under different positions."""
    position_b = position_repository.create_position(
        session=session,
        position_name="Position B",
        department_id=position.department_id,
        duty_days=position.duty_days,
    )

    shift_name = "Shared shift"

    shift_a = create_new_shift(
        session=session,
        name=shift_name,
        position_id=position.id,
        grants_day_off=False,
        doctors_per_shift=1,
    )

    shift_b = create_new_shift(
        session=session,
        name=shift_name,
        position_id=position_b.id,
        grants_day_off=False,
        doctors_per_shift=1,
    )

    assert shift_a.id != shift_b.id
    assert shift_a.position_id == position.id
    assert shift_b.position_id == position_b.id
    assert shift_a.name == shift_b.name == shift_name


def test_shift_name_cannot_repeat_within_position(
    session,
    shift,
):
    """Tests that a shift name is unique within a position."""
    duplicate_shift = ShiftCreate(
        name=shift.name,
        doctors_per_shift=shift.doctors_per_shift,
        grants_day_off=shift.grants_day_off,
        position_id=shift.position_id,
    )

    with pytest.raises(IntegrityError):
        shift_repository.create_shift(
            session=session,
            shift_data=duplicate_shift,
        )

    session.rollback()


def test_soft_deleted_shift_name_remains_reserved_within_position(
    session,
    shift,
):
    """Tests that soft deletion does not release the shift name."""
    shift.is_deleted = True
    session.add(shift)
    session.commit()
    replacement_shift = ShiftCreate(
        name=shift.name,
        doctors_per_shift=shift.doctors_per_shift,
        grants_day_off=shift.grants_day_off,
        position_id=shift.position_id,
    )

    with pytest.raises(IntegrityError):
        shift_repository.create_shift(
            session=session,
            shift_data=replacement_shift,
        )

    session.rollback()
    session.refresh(shift)

    assert shift.is_deleted is True


#####################
# Controller Tests
#####################


def test_create_shift_controller_duplicate_name(session, department, shift):
    """Tests that creating a shift with a duplicate name returns error and no extra row is created."""
    shift2_data = ShiftCreate(
        name=shift.name,
        position_id=shift.position_id,
        grants_day_off=False,
        doctors_per_shift=1
    )

    with pytest.raises(Exception) as exc_info:
        shift_controllers.create_shift_controller(
            shift_data=shift2_data,
            department_id=department.id,
            session=session,
        )

    assert exc_info.type.__name__ == "HTTPException"
    assert exc_info.value.status_code == 400
    assert "already exists" in exc_info.value.detail

    stored_shifts = session.exec(
        select(ShiftModel).where(ShiftModel.position_id == shift.position_id)
    ).all()
    assert len(stored_shifts) == 1


def test_create_shift_controller_nonexistent_position(session, department):
    """Tests that creating a shift under a non existent position returns 404 Position not found."""
    shift_data = ShiftCreate(
        name="ER 1",
        position_id=uuid.uuid4(),
        grants_day_off=False,
        doctors_per_shift=1
    )

    with pytest.raises(Exception) as exc_info:
        shift_controllers.create_shift_controller(
            shift_data=shift_data,
            department_id=department.id,
            session=session,
        )

    assert exc_info.type.__name__ == "HTTPException"
    assert exc_info.value.status_code == 404
    assert "Position not found" in exc_info.value.detail


def test_create_shift_controller_foreign_position(session, department, position_b):
    """Tests that creating a shift under a Department B position from Department A returns 404 and creates no Shift."""
    shift_data = ShiftCreate(
        name="Foreign Position Shift",
        position_id=position_b.id,
        grants_day_off=False,
        doctors_per_shift=1
    )

    with pytest.raises(Exception) as exc_info:
        shift_controllers.create_shift_controller(
            shift_data=shift_data,
            department_id=department.id,
            session=session,
        )

    assert exc_info.type.__name__ == "HTTPException"
    assert exc_info.value.status_code == 404
    assert "Position not found" in exc_info.value.detail

    stored_shifts = session.exec(
        select(ShiftModel).where(
            ShiftModel.position_id == position_b.id,
            ShiftModel.name == "Foreign Position Shift",
        )
    ).all()
    assert stored_shifts == []


def test_get_shift_controller_nonexistent(session, department):
    """Tests that trying to retrieve a non existent shift returns error"""
    with pytest.raises(Exception) as exc_info:
        shift_controllers.get_shift_controller(
            shift_id=uuid.uuid4(),
            department_id=department.id,
            session=session,
        )

    assert exc_info.type.__name__ == "HTTPException"
    assert exc_info.value.status_code == 404
    assert "not found" in exc_info.value.detail


def test_get_shift_controller_deleted(session, shift, position):
    """Tests that trying to retrieve a deleted shift returns error"""
    shift.is_deleted = True
    session.add(shift)
    session.commit()

    with pytest.raises(Exception) as exc_info:
        shift_controllers.get_shift_controller(
            shift_id=shift.id,
            department_id=position.department_id,
            session=session,
        )

    assert exc_info.type.__name__ == "HTTPException"
    assert exc_info.value.status_code == 404
    assert "not found" in exc_info.value.detail


def test_get_shift_controller_foreign_shift(session, department, shift_b):
    """Tests that a Department A caller cannot retrieve a Department B shift by id."""
    with pytest.raises(Exception) as exc_info:
        shift_controllers.get_shift_controller(
            shift_id=shift_b.id,
            department_id=department.id,
            session=session,
        )

    assert exc_info.type.__name__ == "HTTPException"
    assert exc_info.value.status_code == 404
    assert "not found" in exc_info.value.detail


def test_create_shift_assignment_controller_same_doctor_duplicate(session, shift_assignment, department):
    """Tests that creating a new shift assignment with the same doctor returns error"""
    new_shift_assignment_data = ShiftAssignmentCreate(
        doctor_id=shift_assignment.doctor_id,
        date=shift_assignment.date
    )

    with pytest.raises(Exception) as exc_info:
        shift_controllers.create_shift_assignment_controller(
            shift_id=shift_assignment.shift_id,
            department_id=department.id,
            session=session,
            shift_assignment_data=new_shift_assignment_data
        )

    assert exc_info.type.__name__ == "HTTPException"
    assert exc_info.value.status_code == 400
    assert "Doctor is already assigned on this dat" in exc_info.value.detail


def test_create_shift_assignment_controller_foreign_shift(session, department, shift_b, new_doctor):
    """Tests that creating an assignment against a Department B shift returns 404 and creates no assignment."""
    shift_assignment_data = ShiftAssignmentCreate(
        doctor_id=new_doctor.id,
        date=datetime.date(2026, 8, 12),
    )

    with pytest.raises(Exception) as exc_info:
        shift_controllers.create_shift_assignment_controller(
            shift_id=shift_b.id,
            department_id=department.id,
            session=session,
            shift_assignment_data=shift_assignment_data,
        )

    assert exc_info.type.__name__ == "HTTPException"
    assert exc_info.value.status_code == 404
    assert "Shift not found" in exc_info.value.detail
    assert shift_repository.get_shift_assignments_by_date(
        session=session,
        shift_id=shift_b.id,
        target_date=datetime.date(2026, 8, 12),
    ) == []


def test_create_shift_assignment_controller_foreign_doctor(session, department, shift, doctor_b):
    """Tests that creating an assignment for a Department B doctor returns 404 and creates no assignment."""
    shift_assignment_data = ShiftAssignmentCreate(
        doctor_id=doctor_b.id,
        date=datetime.date(2026, 8, 12),
    )

    with pytest.raises(Exception) as exc_info:
        shift_controllers.create_shift_assignment_controller(
            shift_id=shift.id,
            department_id=department.id,
            session=session,
            shift_assignment_data=shift_assignment_data,
        )

    assert exc_info.type.__name__ == "HTTPException"
    assert exc_info.value.status_code == 404
    assert "Doctor not found" in exc_info.value.detail
    assert shift_repository.get_shift_assignments_by_date(
        session=session,
        shift_id=shift.id,
        target_date=datetime.date(2026, 8, 12),
    ) == []


def test_create_shift_assignment_controller_scope_check_precedes_capacity_check(
    session,
    department,
    position_b,
    new_doctor,
):
    """Tests that the foreign-shift scope check runs before the capacity business rule."""
    zero_capacity_foreign_shift = create_new_shift(
        session=session,
        name="Zero Capacity",
        position_id=position_b.id,
        grants_day_off=False,
        doctors_per_shift=0,
    )
    shift_assignment_data = ShiftAssignmentCreate(
        doctor_id=new_doctor.id,
        date=datetime.date(2026, 8, 12),
    )

    with pytest.raises(Exception) as exc_info:
        shift_controllers.create_shift_assignment_controller(
            shift_id=zero_capacity_foreign_shift.id,
            department_id=department.id,
            session=session,
            shift_assignment_data=shift_assignment_data,
        )

    assert exc_info.type.__name__ == "HTTPException"
    assert exc_info.value.status_code == 404
    assert "Shift not found" in exc_info.value.detail


# def test_create_shift_assignment_controller_different_doctor_conflict(session, department, team, new_doctor, shift):
#     """Tests that creating a new shift assignment where another doctor is assigned returns error"""

#     new_doctor2 = create_new_doctor(
#         session=session,
#         name="Dr panostest",
#         email="drpanostest@gmail.com",
#         department_id=department.id,
#         team_id=team.id
#     )

#     create_new_shift_assignment(
#         session=session,
#         doctor_id=new_doctor.id,
#         date=datetime.date(2026, 8, 12),
#         shift_id=shift.id,
#     )

#     new_shift_assignment_data = ShiftAssignmentCreate(
#         date=datetime.date(2026, 8, 12),
#         doctor_id=new_doctor2.id
#     )

#     with pytest.raises(Exception) as exc_info:
#         shift_controllers.create_shift_assignment_controller(
#             shift_id=shift.id,
#             session=session,
#             shift_assignment_data=new_shift_assignment_data
#         )

#     assert exc_info.type.__name__ == "HTTPException"
#     assert exc_info.value.status_code == 422
#     assert "Another doctor is assigned on this shift" in exc_info.value.detail


def test_create_shift_assignment_unavailability_conflict(session, unavailability, shift, department):
    """Tests that creating a shift assignment when a doctor is unavailable returns error"""

    shift_assignment_data = ShiftAssignmentCreate(
        doctor_id=unavailability.doctor_id,
        date=unavailability.date
    )

    with pytest.raises(Exception) as exc_info:
        shift_controllers.create_shift_assignment_controller(
            session=session,
            shift_id=shift.id,
            department_id=department.id,
            shift_assignment_data=shift_assignment_data
        )

    assert exc_info.type.__name__ == "HTTPException"
    assert exc_info.value.status_code == 400
    assert "Doctor is unavailable" in exc_info.value.detail


def test_create_shift_assignment_capacity_limit(session, shift, new_doctor, department, team):
    """Tests that a shift assignment cannot be created when a shift already has reached the assignment capacity"""
    new_doctor2 = create_new_doctor(
        session, "2nd doc", "2nddoc@gmail.com", department.id, team.id)
    new_doctor3 = create_new_doctor(
        session, "3rd doc", "3rddoc@gmail.com", department.id, team.id)

    target_date = datetime.date(2026, 8, 12)

    create_new_shift_assignment(session, new_doctor.id, shift.id, target_date)
    create_new_shift_assignment(session, new_doctor2.id, shift.id, target_date)

    shift_assignment_data = ShiftAssignmentCreate(
        doctor_id=new_doctor3.id,
        date=target_date
    )

    with pytest.raises(Exception) as exc_info:
        shift_controllers.create_shift_assignment_controller(
            session=session,
            shift_id=shift.id,
            department_id=department.id,
            shift_assignment_data=shift_assignment_data
        )

    assert exc_info.type.__name__ == "HTTPException"
    assert exc_info.value.status_code == 400
    assert "Shift already has max assignments" in exc_info.value.detail


#####################
# Route Tests
#####################


def test_create_shift_route(client, position, department_admin_headers):
    """Tests post /api/v1/shifts route"""
    response = client.post(
        "api/v1/shifts",
        json={
            "name": "ER 1",
            "doctors_per_shift": 1,
            "grants_day_off": False,
            "position_id": str(position.id)
        },
        headers=department_admin_headers
    )

    assert response.status_code == 201
    data = response.json()
    assert data["name"] == "ER 1"
    assert data["position_id"] == str(position.id)
    assert "id" in data
    assert "created_at" in data
    assert "updated_at" in data


def test_create_shift_route_invalid_payload(client, position, department_admin_headers):
    """Tests post /api/v1/shifts route with invalid payload"""
    response = client.post(
        "api/v1/shifts",
        json={
            "doctors_per_shift": 1,
            "grants_day_off": False,
            "position_id": str(position.id)
        },
        headers=department_admin_headers
    )

    assert response.status_code == 422


def test_list_shifts_route(client, session, department, department_b, shift, shift_b, viewer_headers):
    """Tests that get /api/v1/shifts only returns shifts scoped to the authenticated department."""
    response = client.get(
        "/api/v1/shifts",
        headers=viewer_headers,
    )

    assert response.status_code == 200
    data = response.json()
    assert isinstance(data, list)

    returned_ids = {item["id"] for item in data}
    assert str(shift.id) in returned_ids
    assert str(shift_b.id) not in returned_ids


def test_create_shift_route_same_name_under_different_position_in_department(
    client,
    session,
    department,
    shift,
    department_admin_headers,
):
    """Tests that a shift name may repeat under a different Position within the same department."""
    other_position = position_repository.create_position(
        session=session,
        position_name="ICU",
        department_id=department.id,
        duty_days=[2, 4, 6],
    )

    response = client.post(
        "api/v1/shifts",
        json={
            "name": shift.name,
            "doctors_per_shift": 1,
            "grants_day_off": False,
            "position_id": str(other_position.id)
        },
        headers=department_admin_headers,
    )

    assert response.status_code == 201
    data = response.json()
    assert data["name"] == shift.name
    assert data["position_id"] == str(other_position.id)


def test_create_shift_route_rejects_admin_without_department(
    client,
    session,
    position,
    user_factory,
    auth_headers_factory,
):
    """Tests that an unscoped department admin cannot create a Shift."""
    shift_name = "Unscoped Shift"
    department_admin = user_factory(
        role=UserRole.DEPARTMENT_ADMIN,
        department_id=None,
    )

    response = client.post(
        "api/v1/shifts",
        json={
            "name": shift_name,
            "doctors_per_shift": 1,
            "grants_day_off": False,
            "position_id": str(position.id)
        },
        headers=auth_headers_factory(department_admin),
    )

    assert response.status_code == 403
    assert response.json() == {"detail": "Invalid account scope."}
    assert response.headers.get("WWW-Authenticate") is None

    stored_shifts = session.exec(
        select(ShiftModel).where(ShiftModel.name == shift_name)
    ).all()
    assert stored_shifts == []


def test_get_shift_route(client, shift, viewer_headers):
    """Tests get /api/v1/shifts/{shift_id} route"""
    response = client.get(
        f"/api/v1/shifts/{shift.id}",
        headers=viewer_headers,
    )

    assert response.status_code == 200
    data = response.json()
    assert data["id"] == str(shift.id)


def test_get_shift_route_malformed_id(client, viewer_headers):
    """Tests get /api/v1/shifts/{shift_id} route with a malformed UUID"""
    response = client.get(
        "/api/v1/shifts/not-a-uuid",
        headers=viewer_headers,
    )

    assert response.status_code == 422


def test_get_shift_route_nonexistent_id(client, viewer_headers):
    """Tests get /api/v1/shifts/{shift_id} route with invalid payload"""
    response = client.get(
        f"/api/v1/shifts/{uuid.uuid4()}",
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
def test_department_member_cannot_get_shift_from_another_department(
    client,
    department,
    shift_b,
    role,
    user_factory,
    auth_headers_factory,
):
    """Tests that foreign Shift IDs are hidden from department members."""
    department_user = user_factory(
        role=role,
        department_id=department.id,
    )

    response = client.get(
        f"/api/v1/shifts/{shift_b.id}",
        headers=auth_headers_factory(department_user),
    )

    assert response.status_code == 404
    assert response.json() == {"detail": "Shift not found"}
    assert response.headers.get("WWW-Authenticate") is None


def test_create_shift_assignment_route(client, shift, new_doctor, department_admin_headers):
    """Tests post /api/v1/shifts/{shift_id}/assignments"""
    response = client.post(
        f"/api/v1/shifts/{shift.id}/assignments",
        json={
            "doctor_id": str(new_doctor.id),
            "date": str(datetime.date(2026, 8, 12))
        },
        headers=department_admin_headers,
    )

    assert response.status_code == 201
    data = response.json()
    assert data["doctor_id"] == str(new_doctor.id)
    assert "id" in data
    assert "created_at" in data
    assert "updated_at" in data


def test_create_shift_assignment_route_invalid_payload(client, shift, new_doctor, department_admin_headers):
    """Tests post /api/v1/shifts/{shift_id}/assignments with invalid payload"""
    response = client.post(
        f"/api/v1/shifts/{shift.id}/assignments",
        json={
            "doctor_id": str(new_doctor.id),
        },
        headers=department_admin_headers,
    )

    assert response.status_code == 422


def test_create_shift_assignment_route_malformed_shift_id(client, new_doctor, department_admin_headers):
    """Tests post /api/v1/shifts/{shift_id}/assignments with a malformed shift id."""
    response = client.post(
        "/api/v1/shifts/not-a-uuid/assignments",
        json={
            "doctor_id": str(new_doctor.id),
            "date": str(datetime.date(2026, 8, 12))
        },
        headers=department_admin_headers,
    )

    assert response.status_code == 422


def test_create_shift_assignment_route_duplicate(client, shift, new_doctor, department_admin_headers):
    """Tests post /api/v1/shifts/{shift_id}/assignments with invalid payload"""
    client.post(
        f"/api/v1/shifts/{shift.id}/assignments",
        json={
            "doctor_id": str(new_doctor.id),
            "date": str(datetime.date(2026, 8, 12))
        },
        headers=department_admin_headers,
    )

    response = client.post(
        f"/api/v1/shifts/{shift.id}/assignments",
        json={
            "doctor_id": str(new_doctor.id),
            "date": str(datetime.date(2026, 8, 12))
        },
        headers=department_admin_headers,
    )

    assert response.status_code == 400


def test_create_shift_assignment_route_mixed_department_shift_and_doctor(
    client,
    session,
    shift,
    doctor_b,
    department_admin_headers,
):
    """Tests that a Department A shift with a Department B doctor returns 404 and creates no assignment."""
    response = client.post(
        f"/api/v1/shifts/{shift.id}/assignments",
        json={
            "doctor_id": str(doctor_b.id),
            "date": str(datetime.date(2026, 8, 12))
        },
        headers=department_admin_headers,
    )

    assert response.status_code == 404
    assert shift_repository.get_shift_assignments_by_date(
        session=session,
        shift_id=shift.id,
        target_date=datetime.date(2026, 8, 12),
    ) == []


def test_create_shift_assignment_route_rejects_admin_without_department(
    client,
    session,
    shift,
    new_doctor,
    user_factory,
    auth_headers_factory,
):
    """Tests that an unscoped department admin cannot create a Shift assignment."""
    department_admin = user_factory(
        role=UserRole.DEPARTMENT_ADMIN,
        department_id=None,
    )

    response = client.post(
        f"/api/v1/shifts/{shift.id}/assignments",
        json={
            "doctor_id": str(new_doctor.id),
            "date": str(datetime.date(2026, 8, 12))
        },
        headers=auth_headers_factory(department_admin),
    )

    assert response.status_code == 403
    assert response.json() == {"detail": "Invalid account scope."}
    assert response.headers.get("WWW-Authenticate") is None
    assert shift_repository.get_shift_assignments_by_date(
        session=session,
        shift_id=shift.id,
        target_date=datetime.date(2026, 8, 12),
    ) == []


def test_list_shift_assignments_route(client, shift_assignment, viewer_headers):
    """Tests get /api/v1/shifts/assignments"""
    response = client.get(
        "api/v1/shifts/assignments",
        headers=viewer_headers
    )

    assert response.status_code == 200
    data = response.json()
    assert isinstance(data, list)

    returned_shift_assignment = next(
        (
            item
            for item in data
            if item["id"] == str(shift_assignment.id)
        ),
        None
    )

    assert returned_shift_assignment is not None
    assert returned_shift_assignment["id"] == str(shift_assignment.id)


@pytest.mark.parametrize(
    "role",
    [
        UserRole.VIEWER,
        UserRole.DOCTOR,
        UserRole.SUPER_ADMIN,
    ]
)
def test_non_department_admin_cannot_create_shift(
    client,
    role,
    position,
    user_factory,
    auth_headers_factory,
    department,
    session,
):
    """Tests post /api/v1/shifts route with non department admin headers"""
    user = user_factory(
        role=role,
        department_id=(
            None
            if role == UserRole.SUPER_ADMIN
            else department.id
        ),
    )
    headers = auth_headers_factory(user)

    response = client.post(
        "/api/v1/shifts",
        json={
            "name": "ER 1",
            "doctors_per_shift": 1,
            "grants_day_off": False,
            "position_id": str(position.id)
        },
        headers=headers,
    )

    assert response.status_code == 403
    assert response.json() == {
        "detail": "Insufficient permissions for this operation"}
    assert response.headers.get("WWW-Authenticate") is None
    assert shift_repository.get_shift_by_name(session, "ER 1") is None


def test_create_shift_requires_authentication(client, session, position):
    """Tests post /api/v1/shifts route without auth"""
    response = client.post(
        "/api/v1/shifts",
        json={
            "name": "ER 1",
            "doctors_per_shift": 1,
            "grants_day_off": False,
            "position_id": str(position.id)
        }
    )

    assert response.status_code == 401
    assert response.json() == {"detail": "Unauthorized"}
    assert response.headers.get("WWW-Authenticate") == "Bearer"
    assert shift_repository.get_shift_by_name(session, "ER 1") is None


@pytest.mark.parametrize(
    "role",
    [
        UserRole.VIEWER,
        UserRole.DOCTOR,
        UserRole.SUPER_ADMIN,
    ]
)
def test_non_department_admin_cannot_create_shift_assignment(
    client,
    role,
    user_factory,
    auth_headers_factory,
    department,
    new_doctor,
    shift,
    session,
):
    """tests post /api/v1/shifts/{shift_id}/assignments with non department admin headers"""
    user = user_factory(
        role=role,
        department_id=(
            None
            if role == UserRole.SUPER_ADMIN
            else department.id
        ),
    )
    headers = auth_headers_factory(user)

    response = client.post(
        f"/api/v1/shifts/{shift.id}/assignments",
        json={
            "doctor_id": str(new_doctor.id),
            "date": str(datetime.date(2026, 8, 12))
        },
        headers=headers,
    )

    assert response.status_code == 403
    assert response.json() == {
        "detail": "Insufficient permissions for this operation"}
    assert response.headers.get("WWW-Authenticate") is None
    retrieved_shift_assignments = shift_repository.get_shift_assignments_by_date(
        session=session,
        shift_id=shift.id,
        target_date=datetime.date(2026, 8, 12)
    )
    assert retrieved_shift_assignments == []


def test_create_shift_assignment_requires_authentication(client, shift, new_doctor, session):
    """Tests post /api/v1/shifts/{shift_id}/assignments without auth"""

    response = client.post(
        f"/api/v1/shifts/{shift.id}/assignments",
        json={
            "doctor_id": str(new_doctor.id),
            "date": str(datetime.date(2026, 8, 12))
        },
    )

    assert response.status_code == 401
    assert response.json() == {"detail": "Unauthorized"}
    assert response.headers.get("WWW-Authenticate") == "Bearer"
    retrieved_shift_assignments = shift_repository.get_shift_assignments_by_date(
        session=session,
        shift_id=shift.id,
        target_date=datetime.date(2026, 8, 12)
    )
    assert retrieved_shift_assignments == []


@pytest.mark.parametrize(
    "path_template",
    [
        pytest.param(
            "/api/v1/shifts/",
            id="list-shifts",
        ),
        pytest.param(
            "/api/v1/shifts/{shift_id}",
            id="get-shift",
        ),
        pytest.param(
            "/api/v1/shifts/assignments",
            id="list-shift-assignments",
        ),
    ],
)
def test_shift_read_routes_require_authentication(
    client,
    shift,
    path_template,
):
    """Tests that the shift read routes require auth"""
    path = path_template.format(shift_id=shift.id)

    response = client.get(path)

    assert response.status_code == 401
    assert response.json() == {"detail": "Unauthorized"}
    assert response.headers.get("WWW-Authenticate") == "Bearer"


@pytest.mark.parametrize(
    "path_template",
    [
        pytest.param(
            "/api/v1/shifts/",
            id="list-shifts",
        ),
        pytest.param(
            "/api/v1/shifts/{shift_id}",
            id="get-shift",
        ),
        pytest.param(
            "/api/v1/shifts/assignments",
            id="list-shift-assignments",
        ),
    ],
)
def test_shift_read_routes_reject_member_without_department(
    client,
    shift,
    path_template,
    user_factory,
    auth_headers_factory,
):
    """Tests that shift reads reject accounts without tenant scope."""
    viewer = user_factory(
        role=UserRole.VIEWER,
        department_id=None,
    )
    path = path_template.format(shift_id=shift.id)

    response = client.get(
        path,
        headers=auth_headers_factory(viewer),
    )

    assert response.status_code == 403
    assert response.json() == {"detail": "Invalid account scope."}
    assert response.headers.get("WWW-Authenticate") is None


def test_list_shift_assignments_route_excludes_foreign_department(
    client,
    session,
    shift_assignment,
    shift_b,
    doctor_b,
    viewer_headers,
):
    """Tests that the assignment list excludes assignments from another department."""
    foreign_assignment = create_new_shift_assignment(
        session=session,
        doctor_id=doctor_b.id,
        shift_id=shift_b.id,
        date=datetime.date(2026, 8, 12),
    )

    response = client.get(
        "api/v1/shifts/assignments",
        headers=viewer_headers,
    )

    assert response.status_code == 200
    returned_ids = {item["id"] for item in response.json()}
    assert str(shift_assignment.id) in returned_ids
    assert str(foreign_assignment.id) not in returned_ids
