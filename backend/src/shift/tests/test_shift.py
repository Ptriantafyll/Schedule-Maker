"""
Tests for the shift module
"""

import uuid
import datetime
import pytest
from sqlmodel import Session

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
from src.team.schemas import TeamCreate
from src.team import repository as team_repository
from src.user.models import UserRole

#####################
# Helpers
#####################


def create_new_doctor(
    session: Session,
    name: str,
    email: str,
    department_id: uuid.UUID,
    team_id: uuid.UUID
) -> DoctorModel:
    """Helper that creates a new doctor in the db"""
    doctor_data = DoctorCreate(
        name=name,
        email=email,
        department_id=department_id,
        team_id=team_id
    )
    return doctor_repository.create_doctor(session, doctor_data)


def create_new_shift(
    session: Session,
    name: str,
    position_id: uuid.UUID,
    grants_day_off: bool,
    doctor_per_shift: int
) -> ShiftModel:
    """Helper that creates a shift in the db"""

    shift_data = ShiftCreate(
        name=name,
        position_id=position_id,
        grants_day_off=grants_day_off,
        doctors_per_shift=doctor_per_shift
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
    position_data = PositionCreate(
        name="ER",
        department_id=department.id,
        duty_days=[1, 3, 5],
    )

    return position_repository.create_position(session, position_data)


@pytest.fixture(name="team")
def team_fixture(session, department):
    """Creates a reusable team for tests"""
    team_data = TeamCreate(name="ER Team A", department_id=department.id)
    return team_repository.create_team(session, team_data)


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


def test_get_shift_by_id(session, shift):
    """Tests retrieving a shift by its id"""
    retrieved_shift = shift_repository.get_shift_by_id(
        session=session,
        shift_id=shift.id
    )

    assert retrieved_shift is not None
    assert retrieved_shift.id == shift.id


def test_get_shift_by_name(session, shift):
    """Tests retrieving a shift by its name"""
    retrieved_shift = shift_repository.get_shift_by_name(
        session=session,
        shift_name=shift.name
    )

    assert retrieved_shift is not None
    assert retrieved_shift.id == shift.id


def test_get_active_shifts(session, shift, position):
    """Tests retrieving all active shifts"""
    shift2 = create_new_shift(session, "ER 2", position.id, False, 1)

    shift2.is_deleted = True
    session.add(shift2)
    session.commit()

    retrieved_shifts = shift_repository.get_active_shifts(session)

    assert isinstance(retrieved_shifts, list)
    assert shift in retrieved_shifts
    assert shift2 not in retrieved_shifts


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


def test_get_active_shift_assignments(session, new_doctor, shift):
    """Tests retrieving active shift assignments"""
    shift_assignment1 = create_new_shift_assignment(
        session=session,
        doctor_id=new_doctor.id,
        shift_id=shift.id,
        date=datetime.date(2026, 8, 12)
    )

    shift_assignment2 = create_new_shift_assignment(
        session=session,
        doctor_id=new_doctor.id,
        shift_id=shift.id,
        date=datetime.date(2026, 8, 12)
    )
    shift_assignment2.is_deleted = True
    session.add(shift_assignment2)
    session.commit()

    retrieved_shift_assignments = shift_repository.get_active_shift_assignments(
        session)

    assert isinstance(retrieved_shift_assignments, list)
    assert shift_assignment1 in retrieved_shift_assignments
    assert shift_assignment2 not in retrieved_shift_assignments

#####################
# Controller Tests
#####################


def test_create_shift_controller_duplicate_name(session, shift):
    """Tests that creating a shift with a duplicate name returns error"""
    shift2_data = ShiftCreate(
        name=shift.name,
        position_id=shift.position_id,
        grants_day_off=False,
        doctor_per_shift=1
    )

    with pytest.raises(Exception) as exc_info:
        shift_controllers.create_shift_controller(shift2_data, session)

    assert exc_info.type.__name__ == "HTTPException"
    assert exc_info.value.status_code == 400
    assert "already exists" in exc_info.value.detail


def test_create_shift_controller_nonexistent_position(session):
    """Tests that creating a shift with a non existent position id returns error"""
    shift_data = ShiftCreate(
        name="ER 1",
        position_id=uuid.uuid4(),
        grants_day_off=False,
        doctors_per_shift=1
    )

    with pytest.raises(Exception) as exc_info:
        shift_controllers.create_shift_controller(shift_data, session)

    assert exc_info.type.__name__ == "HTTPException"
    assert exc_info.value.status_code == 422
    assert "does not exist" in exc_info.value.detail


def test_get_shift_controller_nonexistent(session):
    """Tests that trying to retrieve a non existent shift returns error"""
    with pytest.raises(Exception) as exc_info:
        shift_controllers.get_shift_controller("test", session)

    assert exc_info.type.__name__ == "HTTPException"
    assert exc_info.value.status_code == 404
    assert "not found" in exc_info.value.detail


def test_get_shift_controller_deleted(session, shift):
    """Tests that trying to retrieve a deleted shift returns error"""
    shift.is_deleted = True
    session.add(shift)
    session.commit()

    with pytest.raises(Exception) as exc_info:
        shift_controllers.get_shift_controller(shift.name, session)

    assert exc_info.type.__name__ == "HTTPException"
    assert exc_info.value.status_code == 404
    assert "not found" in exc_info.value.detail


def test_create_shift_assignment_controller_same_doctor_duplicate(session, shift_assignment):
    """Tests that creating a new shift assignment with the same doctor returns error"""
    new_shift_assignment_data = ShiftAssignmentCreate(
        doctor_id=shift_assignment.doctor_id,
        date=shift_assignment.date
    )

    with pytest.raises(Exception) as exc_info:
        shift_controllers.create_shift_assignment_controller(
            shift_id=shift_assignment.shift_id,
            session=session,
            shift_assignment_data=new_shift_assignment_data
        )

    assert exc_info.type.__name__ == "HTTPException"
    assert exc_info.value.status_code == 400
    assert "Doctor is already assigned on this dat" in exc_info.value.detail


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


def test_create_shift_assignment_unavailability_conflict(session, unavailability, shift):
    """Tests that creating a shift assignment when a doctor is unavailable returns error"""

    shift_assignment_data = ShiftAssignmentCreate(
        doctor_id=unavailability.doctor_id,
        date=unavailability.date
    )

    with pytest.raises(Exception) as exc_info:
        shift_controllers.create_shift_assignment_controller(
            session=session,
            shift_id=shift.id,
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


def test_list_shifts_route(client, shift, viewer_headers):
    """Tests get /api/v1/shifts route"""
    response = client.get(
        "/api/v1/shifts",
        headers=viewer_headers,
    )

    assert response.status_code == 200
    data = response.json()
    assert isinstance(data, list)

    returned_shift = next(
        (
            item
            for item in data
            if item["id"] == str(shift.id)
        ),
        None
    )

    assert returned_shift is not None
    assert returned_shift["id"] == str(shift.id)


def test_get_shift_route(client, shift, viewer_headers):
    """Tests get /api/v1/shifts/{shift_name} route"""
    response = client.get(
        f"/api/v1/shifts/{shift.name}",
        headers=viewer_headers,
    )

    assert response.status_code == 200
    data = response.json()
    assert data["id"] == str(shift.id)


def test_get_shift_route_nonexistent_id(client, viewer_headers):
    """Tests get /api/v1/shifts/{shift_id} route with invalid payload"""
    response = client.get(
        f"/api/v1/shifts/{uuid.uuid4()}",
        headers=viewer_headers,
    )

    assert response.status_code == 404


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
    shift
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
