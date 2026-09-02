"""
Tests for the doctor module
"""

import uuid
import datetime
import pytest
from sqlmodel import Session


from src.department.schemas import DepartmentCreate
from src.department import repository as department_repository
from src.team.schemas import TeamCreate
from src.team.repository import create_team
from src.shift import repository as shift_repository
from src.shift.schemas import ShiftCreate
from src.position import repository as position_repository
from src.position.schemas import PositionCreate
from src.position.models import Position as PositionModel
from src.doctor.models import Doctor as DoctorModel
from src.doctor.schemas import (
    DoctorCreate,
    DoctorPreAssignmentCreate,
    DoctorUnavailabilityCreate,
    DoctorPositionCreate,
)
from src.doctor import repository as doctor_repository
from src.doctor import controllers as doctor_controllers
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


def create_test_pre_assignment(
    session: Session,
    shift_id: str,
    doctor: DoctorModel,
    date: datetime.date
):
    """Helper that creates a new doctor pre assignment in the db"""
    pre_assignment_data = DoctorPreAssignmentCreate(
        date=date,
        shift_id=shift_id
    )

    return doctor_repository.create_doctor_pre_assignment(
        session, doctor.id, pre_assignment_data
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


def create_test_doctor_position(session: Session, doctor: DoctorModel, position: PositionModel):
    """Helper that creates a doctor-position"""
    doctor_pos_data = DoctorPositionCreate(
        position_id=position.id
    )
    return doctor_repository.create_doctor_position(
        session=session,
        doctor_id=doctor.id,
        doctor_pos_data=doctor_pos_data
    )


#####################
# Fixtures
#####################


@pytest.fixture(name="department")
def department_fixture(session):
    """Creates a reusable department for tests"""
    dept_data = DepartmentCreate(name="Cardiology", code="CARD")
    return department_repository.create_department(session, dept_data)


@pytest.fixture(name="team")
def team_fixture(session, department):
    """Creates a reusable team for tests"""
    team_data = TeamCreate(name="ER Team A", department_id=department.id)
    return create_team(
        session=session,
        name=team_data.name,
        department_id=team_data.department_id,
    )


@pytest.fixture(name="position")
def position_fixture(session, department):
    """Creates a reusable position for tests"""
    position_data = PositionCreate(
        name="ER",
        department_id=department.id,
        duty_days=[1, 3, 5],
    )

    return position_repository.create_position(session, position_data)


@pytest.fixture(name="shift")
def shift_fixture(session, position):
    """Creates a reusable shift for tests"""
    shift_data = ShiftCreate(
        name="ER 1",
        position_id=position.id,
        grants_day_off=False,
        doctors_per_shift=2
    )

    return shift_repository.create_shift(session, shift_data)


@pytest.fixture(name="new_doctor")
def doctor_fixture(session, department, team):
    """Creates a reusable doctor for tests"""
    return create_new_doctor(session, "Dr Panos", "drpanos@gmail.com", department.id, team.id)


@pytest.fixture(name="pre_assignment")
def pre_assignment_fixture(session, new_doctor, shift):
    """Creates a reusable pre-assignment for tests"""
    return create_test_pre_assignment(session, shift.id, new_doctor, datetime.date(2026, 8, 12))


@pytest.fixture(name="unavailability")
def unavailability_fixture(session, new_doctor):
    """Creates a reusable pre-assignment for tests"""
    return create_test_unavailability(session, new_doctor, datetime.date(2026, 8, 12))


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


@pytest.fixture(name="doctor_user")
def doctor_user_fixture(user_factory, department, new_doctor):
    """Creates a reusable doctor user for tests"""
    return user_factory(
        role=UserRole.DOCTOR,
        department_id=department.id,
        doctor_id=new_doctor.id
    )


@pytest.fixture(name="doctor_headers")
def doctor_headers_fixture(doctor_user, auth_headers_factory):
    """Creates reusable doctor auth headers for tests"""

    return auth_headers_factory(doctor_user)

#####################
# Repository tests
#####################


def test_create_doctor(department, team, new_doctor):
    """Test creating a doctor and verifying their fields"""
    assert isinstance(new_doctor.id, uuid.UUID)
    assert new_doctor.name == "Dr Panos"
    assert new_doctor.email == "drpanos@gmail.com"
    assert new_doctor.department_id == department.id
    assert new_doctor.team_id == team.id
    assert new_doctor.is_deleted is False
    assert new_doctor.sync_status is False
    assert isinstance(new_doctor.created_at, datetime.datetime)
    assert isinstance(new_doctor.updated_at, datetime.datetime)


def test_get_doctor_by_email(session, new_doctor):
    """Test retrieving a doctor by email"""
    retrieved_doctor = doctor_repository.get_doctor_by_email(
        session, "drpanos@gmail.com")

    assert retrieved_doctor is not None
    assert retrieved_doctor.id == new_doctor.id


def test_get_doctor_by_id(session, new_doctor):
    """Test retrieving a doctor by id"""

    retrieved_doctor = doctor_repository.get_doctor_by_id(
        session, new_doctor.id)

    assert retrieved_doctor is not None
    assert retrieved_doctor.id == new_doctor.id


def test_get_active_doctors(session, team, department, new_doctor):
    """Test retrieving all active doctors"""
    new_doctor2 = create_new_doctor(
        session, "Dr Panagiotis", "drpanagiotis@gmail.com", department.id, team.id)

    new_doctor2.is_deleted = True
    session.add(new_doctor2)
    session.commit()

    active_doctors = doctor_repository.get_active_doctors(session)

    assert new_doctor in active_doctors
    assert new_doctor2 not in active_doctors


def test_create_doctor_pre_assignment(shift, new_doctor, pre_assignment):
    """Test creating pre assignments for a doctor"""
    assert isinstance(pre_assignment.id, uuid.UUID)
    assert pre_assignment.date == datetime.date(2026, 8, 12)
    assert pre_assignment.shift_id == shift.id
    assert pre_assignment.doctor_id == new_doctor.id


def test_get_doctor_pre_assignment(session, new_doctor, pre_assignment):
    """Test retrieving a doctor's pre assignments"""
    pre_assignments = doctor_repository.get_doctor_pre_assignments(
        session, new_doctor.id)

    assert isinstance(pre_assignments, list)
    assert pre_assignment in pre_assignments


def test_get_doctor_pre_assignment_by_date(session, new_doctor, pre_assignment):
    """Test retrieving a doctor's pre assignment by date"""
    retrieved_pre_assignment = doctor_repository.get_doctor_pre_assignment_by_date(
        session=session,
        doctor_id=new_doctor.id,
        target_date=datetime.date(2026, 8, 12)
    )

    assert isinstance(retrieved_pre_assignment.id, uuid.UUID)
    assert pre_assignment.id == retrieved_pre_assignment.id


def test_get_doctor_pre_assignment_by_id(session, pre_assignment):
    """Test retrieving a pre assignment by its id"""
    retrieved_pre_assignment = doctor_repository.get_doctor_pre_assignment_by_id(
        session=session,
        pre_assignment_id=pre_assignment.id
    )

    assert isinstance(retrieved_pre_assignment.id, uuid.UUID)
    assert pre_assignment.id == retrieved_pre_assignment.id


def test_create_doctor_unavailability(session, new_doctor, unavailability):
    """Test creating unavailability dates for a doctor"""
    assert isinstance(unavailability.id, uuid.UUID)
    assert unavailability.date == datetime.date(2026, 8, 12)
    assert unavailability.doctor_id == new_doctor.id


def test_get_doctor_unavailability(session, new_doctor, unavailability):
    """Test retrieving unavailability dates for a doctor"""
    doctor_unavailabilities = doctor_repository.get_doctor_unavailability(
        session=session,
        doctor_id=new_doctor.id
    )

    assert isinstance(doctor_unavailabilities, list)
    assert unavailability in doctor_unavailabilities


def test_get_doctor_unavailability_by_date(session, new_doctor, unavailability):
    """Test retrieving 1 doctor unavailability by date"""
    retrieved_unavailability = doctor_repository.get_doctor_unavailability_by_date(
        session=session,
        doctor_id=new_doctor.id,
        target_date=datetime.date(2026, 8, 12)
    )

    assert isinstance(retrieved_unavailability.id, uuid.UUID)
    assert unavailability.id == retrieved_unavailability.id


def test_create_doctor_position(session, new_doctor, position):
    """Test creating a position for a doctor"""
    new_doctor_pos = create_test_doctor_position(session, new_doctor, position)

    assert isinstance(new_doctor_pos.id, uuid.UUID)
    assert new_doctor_pos.position_id == position.id
    assert new_doctor_pos.doctor_id == new_doctor.id


def test_get_doctor_positions(session, new_doctor, position):
    """Tests getting all the positions of a doctor"""
    new_doctor_pos = create_test_doctor_position(session, new_doctor, position)

    doctor_pos = doctor_repository.get_doctor_positions(
        session=session,
        doctor_id=new_doctor.id
    )

    assert isinstance(doctor_pos, list)
    assert new_doctor_pos in doctor_pos


def test_get_doctor_position_by_id(session, new_doctor, position):
    """Test retrieving a doctor-position by its id"""
    new_doctor_pos = create_test_doctor_position(
        session, new_doctor, position
    )

    retrieved_doctor_pos = doctor_repository.get_doctor_position_by_id(
        session=session,
        doctor_id=new_doctor.id,
        position_id=position.id
    )

    assert isinstance(retrieved_doctor_pos.id, uuid.UUID)
    assert retrieved_doctor_pos.id == new_doctor_pos.id


def test_doctor_has_department_foreign_key(session, new_doctor, department, team):
    """Test that doctors can be associated with a department and retrieved correctly."""

    assert isinstance(new_doctor.id, uuid.UUID)
    assert new_doctor.department_id == department.id
    assert new_doctor.team_id == team.id


#######################
# Controller tests
#######################


def test_create_doctor_controller_duplicate_name(session, department, team, new_doctor):
    """Test that creating a doctor with a duplicate email raises an error."""
    doctor_data = DoctorCreate(
        name=new_doctor.name,
        email=new_doctor.email,
        department_id=department.id,
        team_id=team.id
    )

    with pytest.raises(Exception) as exc_info:
        doctor_controllers.create_doctor_controller(doctor_data, session)

    assert exc_info.type.__name__ == "HTTPException"
    assert exc_info.value.status_code == 400
    assert "already exists" in exc_info.value.detail


def test_get_doctor_controller_nonexistent(session):
    """Test that retrieving a non-existent doctor raises a 404 error"""
    non_existent_id = uuid.uuid4()
    with pytest.raises(Exception) as exc_info:
        doctor_controllers.get_doctor_controller(
            session=session, doctor_id=non_existent_id)

    assert exc_info.type.__name__ == "HTTPException"
    assert exc_info.value.status_code == 404
    assert "not found" in exc_info.value.detail


def test_get_doctor_controller_deleted(session, new_doctor):
    """Test that retrieving a deleted team raises a 404 error"""
    new_doctor.is_deleted = True
    session.add(new_doctor)
    session.commit()

    with pytest.raises(Exception) as exc_info:
        doctor_controllers.get_doctor_controller(session, new_doctor.id)

    assert exc_info.type.__name__ == "HTTPException"
    assert exc_info.value.status_code == 404
    assert "not found" in exc_info.value.detail


def test_create_doctor_controller_invalid_team_or_department(session):
    """Test that creating a doctor with invalid team or department returns error"""
    doctor_data = DoctorCreate(
        name="Dr Panos",
        email="drpanos@gmail.com",
        department_id=uuid.uuid4(),
        team_id=uuid.uuid4()
    )

    with pytest.raises(Exception) as exc_info:
        doctor_controllers.create_doctor_controller(doctor_data, session)

    assert exc_info.type.__name__ == "HTTPException"
    assert exc_info.value.status_code == 422
    assert "does not exist" in exc_info.value.detail


def test_create_doctor_pre_assignment_controller_duplicate_date(session, new_doctor, pre_assignment):
    """Tests that creating a duplicate pre assignment returns error"""
    # pre_assignment is needed to crate the first (duplicate) pre assignment
    pre_assignment_data = DoctorPreAssignmentCreate(
        date=pre_assignment.date,
        shift_id=pre_assignment.shift_id
    )

    with pytest.raises(Exception) as exc_info:
        doctor_controllers.create_doctor_pre_assignment_controller(
            session=session,
            doctor_id=new_doctor.id,
            pre_assignment_data=pre_assignment_data
        )

    assert exc_info.type.__name__ == "HTTPException"
    assert exc_info.value.status_code == 400
    assert "already exists" in exc_info.value.detail


def test_create_doctor_pre_assignment_controller_nonexistent_doctor(session, shift):
    """Tests that creating a pre assignment with a nonexistent doctor returns error"""
    pre_assignment_data = DoctorPreAssignmentCreate(
        date=datetime.date(2026, 8, 12),
        shift_id=shift.id
    )

    with pytest.raises(Exception) as exc_info:
        doctor_controllers.create_doctor_pre_assignment_controller(
            session=session,
            doctor_id=uuid.uuid4(),
            pre_assignment_data=pre_assignment_data
        )

    assert exc_info.type.__name__ == "HTTPException"
    assert exc_info.value.status_code == 422
    assert "Doctor or shift does not exist" in exc_info.value.detail


def test_create_doctor_unavailability_controller_duplicate_date(session, new_doctor, unavailability):
    """Test that creating a duplicate unavailability returns error"""
    unavailability_data = DoctorUnavailabilityCreate(
        date=unavailability.date,
    )

    with pytest.raises(Exception) as exc_info:
        doctor_controllers.create_doctor_unavailabilty_controller(
            session=session,
            doctor_id=new_doctor.id,
            unavailability_data=unavailability_data
        )

    assert exc_info.type.__name__ == "HTTPException"
    assert exc_info.value.status_code == 400
    assert "already exists" in exc_info.value.detail


def test_create_doctor_unavailability_controller_nonexistent_doctor(session):
    """Test that creating an unavailability for a non-existent doctor returns error"""
    unavailability_data = DoctorUnavailabilityCreate(
        date=datetime.date(2026, 8, 12),
    )

    with pytest.raises(Exception) as exc_info:
        doctor_controllers.create_doctor_unavailabilty_controller(
            session=session,
            doctor_id=uuid.uuid4(),
            unavailability_data=unavailability_data
        )

    assert exc_info.type.__name__ == "HTTPException"
    assert exc_info.value.status_code == 422
    assert "Doctor does not exist" in exc_info.value.detail


def test_create_doctor_position_controller_duplicate_assignment(session, new_doctor, position):
    """Tests that creating a duplicate doctor-position returns error"""
    create_test_doctor_position(session, new_doctor, position)

    doctor_position_data = DoctorPositionCreate(
        position_id=position.id
    )

    with pytest.raises(Exception) as exc_info:
        doctor_controllers.create_doctor_position_controller(
            session=session,
            doctor_id=new_doctor.id,
            doctor_pos_data=doctor_position_data
        )

    assert exc_info.type.__name__ == "HTTPException"
    assert exc_info.value.status_code == 400
    assert "already assigned" in exc_info.value.detail


def test_create_doctor_position_controller_nonexistent_doctor(session, position):
    """Tests that creating a doctor-position with a nonexistent doctor returns error"""
    doctor_position_data = DoctorPositionCreate(
        position_id=position.id
    )

    with pytest.raises(Exception) as exc_info:
        doctor_controllers.create_doctor_position_controller(
            session=session,
            doctor_id=uuid.uuid4(),
            doctor_pos_data=doctor_position_data
        )

    assert exc_info.type.__name__ == "HTTPException"
    assert exc_info.value.status_code == 422
    assert "Doctor does not exist" in exc_info.value.detail


def test_create_doctor_mismatched_team_and_department(session, team, department):
    """Tests that doctor's team belong to their department"""
    # department input is needed to add the first department
    department_data = DepartmentCreate(name="dep2", code="T")
    new_department = department_repository.create_department(
        session, department_data)

    doctor_data = DoctorCreate(
        name="Dr Panos",
        email="drpanos@gmail.com",
        department_id=new_department.id,
        team_id=team.id
    )

    with pytest.raises(Exception) as exc_info:
        doctor_controllers.create_doctor_controller(
            doctor_data=doctor_data,
            session=session
        )

    assert exc_info.type.__name__ == "HTTPException"
    assert exc_info.value.status_code == 422
    assert "needs to match" in exc_info.value.detail


def test_create_doctor_position_mismatched_department(session, department, team,  position):
    """Tests that doctor's position belongs to their department"""
    department_data = DepartmentCreate(name="dep2", code="T")
    new_department = department_repository.create_department(
        session, department_data)

    new_doctor = create_new_doctor(
        session=session,
        name="Dr Panos",
        email="drpanos@gmail.com",
        department_id=new_department.id,
        team_id=team.id
    )

    doctor_pos_data = DoctorPositionCreate(
        position_id=position.id
    )

    with pytest.raises(Exception) as exc_info:
        doctor_controllers.create_doctor_position_controller(
            session=session,
            doctor_id=new_doctor.id,
            doctor_pos_data=doctor_pos_data
        )

    assert exc_info.type.__name__ == "HTTPException"
    assert exc_info.value.status_code == 422
    assert "needs to match" in exc_info.value.detail


def test_create_pre_assignment_unavailability_conflict(session, new_doctor, unavailability, shift):
    """Tests that a pre assignment cannot be assigned on an unavailable day"""
    pre_assignment_data = DoctorPreAssignmentCreate(
        date=unavailability.date,
        shift_id=shift.id,
    )

    with pytest.raises(Exception) as exc_info:
        doctor_controllers.create_doctor_pre_assignment_controller(
            session=session,
            doctor_id=new_doctor.id,
            pre_assignment_data=pre_assignment_data
        )

    assert exc_info.type.__name__ == "HTTPException"
    assert exc_info.value.status_code == 422
    assert "cannot be assigned to an unavailable day" in exc_info.value.detail

#######################
# Route tests
#######################


def test_create_doctor_route(client, department, team, department_admin_headers):
    """Test the POST /doctors/ route for creating a doctor"""

    response = client.post(
        "api/v1/doctors",
        json={
            "name": "Dr Panos",
            "email": "drpanos@gmail.com",
            "department_id": str(department.id),
            "team_id": str(team.id)
        },
        headers=department_admin_headers,
    )

    assert response.status_code == 201
    data = response.json()
    assert data["name"] == "Dr Panos"
    assert data["email"] == "drpanos@gmail.com"
    assert data["department_id"] == str(department.id)
    assert data["team_id"] == str(team.id)
    assert "id" in data
    assert "created_at" in data
    assert "updated_at" in data


def test_create_doctor_route_invalid_payload(client, department, team, department_admin_headers):
    """Test the POST /doctors/ route rejects invalid payload"""

    response = client.post(
        "api/v1/doctors",
        json={
            "name": "Dr Panos",
            "department_id": str(department.id),
            "team_id": str(team.id),
        },
        headers=department_admin_headers,
    )
    assert response.status_code == 422


def test_get_doctor_by_id_route(client, department, team, new_doctor, viewer_headers):
    """Tests that the GET /doctors/{doctor_id} route returns a doctor"""
    response = client.get(
        f"/api/v1/doctors/{new_doctor.id}",
        headers=viewer_headers,
    )

    assert response.status_code == 200
    data = response.json()
    assert data["name"] == "Dr Panos"
    assert data["email"] == "drpanos@gmail.com"
    assert data["department_id"] == str(department.id)
    assert data["team_id"] == str(team.id)
    assert "id" in data
    assert "created_at" in data
    assert "updated_at" in data


def test_get_doctor_by_id_route_nonexistent(client, viewer_headers):
    """Test the GET /doctors/{doctor_id} route returns error when given a nonexistent id"""
    response = client.get(
        f"/api/v1/doctors/{uuid.uuid4()}",
        headers=viewer_headers,
    )
    assert response.status_code == 404


# new doctor is needed to add a doctor in the db
def test_list_doctors_route(client, new_doctor, viewer_headers):
    """Tests the GET /doctors/ route"""
    response = client.get(
        "api/v1/doctors",
        headers=viewer_headers,
    )

    assert response.status_code == 200

    returned_doctor = next(
        (
            item
            for item in response.json()
            if item["id"] == str(new_doctor.id)
        ),
        None
    )

    assert returned_doctor is not None
    assert returned_doctor["id"] == str(new_doctor.id)


def test_create_doctor_pre_assignments_route(client, shift, new_doctor, department_admin_headers):
    """Tests the POST /doctors/{doctor_id}/pre-assignments route"""

    response = client.post(
        f"api/v1/doctors/{new_doctor.id}/pre-assignments",
        json={
            "date": str(datetime.date(2026, 8, 12)),
            "doctor_id": str(new_doctor.id),
            "shift_id": str(shift.id)
        },
        headers=department_admin_headers,
    )

    assert response.status_code == 201
    data = response.json()
    assert data["date"] == str(datetime.date(2026, 8, 12))
    assert data["doctor_id"] == str(new_doctor.id)
    assert data["shift_id"] == str(shift.id)
    assert "id" in data
    assert "created_at" in data
    assert "updated_at" in data


def test_create_doctor_pre_assignments_route_invalid_payload(client,  new_doctor, department_admin_headers):
    """Tests the POST /doctors/{doctor_id}/pre-assignments route rejects invalid payload"""
    response = client.post(
        f"api/v1/doctors/{new_doctor.id}/pre-assignments",
        json={
            "date": str(datetime.date(2026, 8, 12)),
            "doctor_id": str(new_doctor.id),
        },
        headers=department_admin_headers,
    )
    assert response.status_code == 422


def test_get_doctor_pre_assignments_route(
    client,
    new_doctor,
    pre_assignment,
    department_admin_headers,
):
    """Tests the GET /doctors/{doctor_id}/pre-assignments route"""
    response = client.get(
        f"api/v1/doctors/{new_doctor.id}/pre-assignments",
        headers=department_admin_headers,
    )

    assert response.status_code == 200
    data = response.json()
    assert isinstance(data, list)

    returned_pre_assignment = next(
        (
            item
            for item in data
            if item["id"] == str(pre_assignment.id)
        ),
        None
    )

    assert returned_pre_assignment["date"] == str(datetime.date(2026, 8, 12))
    assert returned_pre_assignment["doctor_id"] == str(new_doctor.id)
    assert returned_pre_assignment["id"] == str(pre_assignment.id)
    assert "created_at" in returned_pre_assignment
    assert "updated_at" in returned_pre_assignment


def test_create_doctor_unavailability_route(client, new_doctor, doctor_headers):
    """Tests the POST /doctors/{doctor_id}/unavailability route"""
    response = client.post(
        f"api/v1/doctors/{new_doctor.id}/unavailability",
        json={
            "date": str(datetime.date(2026, 8, 12)),
        },
        headers=doctor_headers,
    )

    assert response.status_code == 201
    data = response.json()
    assert data["date"] == str(datetime.date(2026, 8, 12))
    assert data["doctor_id"] == str(new_doctor.id)
    assert "id" in data
    assert "created_at" in data
    assert "updated_at" in data


def test_create_doctor_unavailability_route_invalid_payload(client, new_doctor, doctor_headers):
    """Tests the POST /doctors/{doctor_id}/unavailability route rejects invalid payload"""
    response = client.post(
        f"api/v1/doctors/{new_doctor.id}/unavailability",
        json={},
        headers=doctor_headers
    )
    assert response.status_code == 422


def test_get_doctor_unavailability_route(client, new_doctor, unavailability, doctor_headers):
    """Tests the GET /doctors/{doctor_id}/unavailability route"""
    response = client.get(
        f"api/v1/doctors/{new_doctor.id}/unavailability",
        headers=doctor_headers
    )

    assert response.status_code == 200
    data = response.json()

    returned_unavailability = next(
        (
            item
            for item in data
            if item["id"] == str(unavailability.id)
        ),
        None
    )

    assert isinstance(data, list)
    assert returned_unavailability is not None
    assert returned_unavailability["id"] == str(unavailability.id)
    assert returned_unavailability["date"] == str(datetime.date(2026, 8, 12))
    assert returned_unavailability["doctor_id"] == str(new_doctor.id)
    assert "created_at" in returned_unavailability
    assert "updated_at" in returned_unavailability


def test_create_doctor_position_route(client, position, new_doctor, department_admin_headers):
    """Tests the POST /doctors/{doctor_id}/position route"""
    response = client.post(
        f"api/v1/doctors/{new_doctor.id}/position",
        json={
            "position_id": str(position.id)
        },
        headers=department_admin_headers,
    )

    assert response.status_code == 201
    data = response.json()
    assert data["doctor_id"] == str(new_doctor.id)
    assert "id" in data
    assert "created_at" in data
    assert "updated_at" in data


def test_create_doctor_position_route_invalid_payload(client, new_doctor, department_admin_headers):
    """Tests the POST /doctors/{doctor_id}/position route rejects invalid payload"""
    response = client.post(
        f"api/v1/doctors/{new_doctor.id}/position",
        json={},
        headers=department_admin_headers,
    )

    assert response.status_code == 422


def test_get_doctor_position_route(client, session, new_doctor, position, viewer_headers):
    """Tests the GET /doctors/{doctor_id}/position route"""
    new_doctor_pos = create_test_doctor_position(
        session=session,
        doctor=new_doctor,
        position=position
    )

    response = client.get(
        f"api/v1/doctors/{new_doctor.id}/position",
        headers=viewer_headers,
    )

    assert response.status_code == 200
    data = response.json()
    assert isinstance(data, list)

    returned_doctor_position = next(
        (
            item
            for item in data
            if item["id"] == str(new_doctor_pos.id)
        ),
        None
    )

    assert returned_doctor_position is not None
    assert returned_doctor_position["id"] == str(new_doctor_pos.id)
    assert returned_doctor_position["doctor_id"] == str(new_doctor.id)
    assert "created_at" in returned_doctor_position
    assert "updated_at" in returned_doctor_position


@pytest.mark.parametrize(
    "role",
    [
        UserRole.VIEWER,
        UserRole.DOCTOR,
        UserRole.SUPER_ADMIN,
    ]
)
def test_non_department_admin_cannot_create_doctor(
    client,
    role,
    user_factory,
    auth_headers_factory,
    department,
    team,
    session,
):
    """Tests post /api/v1/doctors route with non department admin headers"""
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
        "/api/v1/doctors",
        json={
            "name": "Dr Panos",
            "email": "drpanos@gmail.com",
            "department_id": str(department.id),
            "team_id": str(team.id),
        },
        headers=headers,
    )

    assert response.status_code == 403
    assert response.json() == {
        "detail": "Insufficient permissions for this operation"}
    assert response.headers.get("WWW-Authenticate") is None
    assert doctor_repository.get_doctor_by_email(
        session=session,
        email="drpanos@gmail.com",
    ) is None


def test_create_doctor_requires_authentication(client, session, department, team):
    """Tests the post /api/v1/doctors route without auth"""
    response = client.post(
        "/api/v1/doctors",
        json={
            "name": "Dr Panos",
            "email": "drpanos@gmail.com",
            "department_id": str(department.id),
            "team_id": str(team.id),
        },
    )

    assert response.status_code == 401
    assert response.json() == {"detail": "Unauthorized"}
    assert response.headers.get("WWW-Authenticate") == "Bearer"
    assert doctor_repository.get_doctor_by_email(
        session=session,
        email="drpanos@gmail.com",
    ) is None


@pytest.mark.parametrize(
    "path_template",
    [
        pytest.param(
            "/api/v1/doctors/",
            id="list-doctors",
        ),
        pytest.param(
            "/api/v1/doctors/{doctor_id}",
            id="get-doctor",
        ),
    ],
)
def test_doctor_read_routes_require_authentication(
    client,
    new_doctor,
    path_template,
):
    """Tests that the doctor read routes require auth"""
    path = path_template.format(doctor_id=new_doctor.id)

    response = client.get(path)

    assert response.status_code == 401
    assert response.json() == {"detail": "Unauthorized"}
    assert response.headers.get("WWW-Authenticate") == "Bearer"


@pytest.mark.parametrize(
    "role",
    [
        UserRole.VIEWER,
        UserRole.DOCTOR,
        UserRole.SUPER_ADMIN,
    ]
)
def test_non_department_admin_cannot_create_pre_assignment(
    client,
    role,
    user_factory,
    auth_headers_factory,
    department,
    new_doctor,
    shift,
    session,
):
    """Tests post /api/v1/doctors/{doctor_id}/pre-assignments route with non department admin headers"""
    user = user_factory(
        role=role,
        department_id=(
            None
            if role == UserRole.SUPER_ADMIN
            else department.id
        ),
        doctor_id=(
            new_doctor.id
            if role == UserRole.DOCTOR
            else None
        ),
    )
    headers = auth_headers_factory(user)

    response = client.post(
        f"api/v1/doctors/{new_doctor.id}/pre-assignments",
        json={
            "date": str(datetime.date(2026, 8, 12)),
            "doctor_id": str(new_doctor.id),
            "shift_id": str(shift.id)
        },
        headers=headers,
    )

    assert response.status_code == 403
    assert response.json() == {
        "detail": "Insufficient permissions for this operation"}
    assert response.headers.get("WWW-Authenticate") is None
    retrieved_pre_assignment = doctor_repository.get_doctor_pre_assignment_by_date(
        session=session,
        doctor_id=new_doctor.id,
        target_date=datetime.date(2026, 8, 12),
    )

    assert retrieved_pre_assignment is None


def test_create_pre_assignment_requires_authentication(
    client,
    session,
    new_doctor,
    shift,
):
    """Tests the post /api/v1/doctors/{doctor_id}/pre-assignments route without auth"""
    response = client.post(
        f"api/v1/doctors/{new_doctor.id}/pre-assignments",
        json={
            "date": str(datetime.date(2026, 8, 12)),
            "doctor_id": str(new_doctor.id),
            "shift_id": str(shift.id)
        },
    )

    assert response.status_code == 401
    assert response.json() == {"detail": "Unauthorized"}
    assert response.headers.get("WWW-Authenticate") == "Bearer"

    retrieved_pre_assignment = doctor_repository.get_doctor_pre_assignment_by_date(
        session=session,
        doctor_id=new_doctor.id,
        target_date=datetime.date(2026, 8, 12),
    )

    assert retrieved_pre_assignment is None


@pytest.mark.parametrize(
    "role",
    [
        UserRole.DOCTOR,
        UserRole.VIEWER,
        UserRole.SUPER_ADMIN,
    ]
)
def test_non_department_admin_cannot_list_pre_assignments(
    client,
    new_doctor,
    user_factory,
    auth_headers_factory,
    role,
    department,
):
    """Tests that the list pre assignments route does not work with non dept admin headers"""
    user = user_factory(
        role=role,
        department_id=(
            None
            if role == UserRole.SUPER_ADMIN
            else department.id
        ),
        doctor_id=(
            new_doctor.id
            if role == UserRole.DOCTOR
            else None
        ),
    )
    headers = auth_headers_factory(user)

    response = client.get(
        f"api/v1/doctors/{new_doctor.id}/pre-assignments",
        headers=headers,
    )

    assert response.status_code == 403
    assert response.json() == {
        "detail": "Insufficient permissions for this operation"}
    assert response.headers.get("WWW-Authenticate") is None


def test_list_pre_assignments_requires_authentication(client, new_doctor):
    """Tests the get /api/v1/doctors/{doctor_id}/pre-assignments route with no headers"""
    response = client.get(
        f"/api/v1/doctors/{new_doctor.id}/pre-assignments"
    )

    assert response.status_code == 401
    assert response.json() == {"detail": "Unauthorized"}
    assert response.headers.get("WWW-Authenticate") == "Bearer"


@pytest.mark.parametrize(
    "role",
    [
        UserRole.VIEWER,
        UserRole.SUPER_ADMIN,
    ]
)
def test_non_doctor_or_department_admin_cannot_create_unavailability(
    client,
    role,
    user_factory,
    auth_headers_factory,
    department,
    new_doctor,
    session,
):
    """Tests the POST /doctors/{doctor_id}/unavailability route with non allowed headers"""
    user = user_factory(
        role=role,
        department_id=(
            None
            if role == UserRole.SUPER_ADMIN
            else department.id
        ),
        doctor_id=None,
    )
    headers = auth_headers_factory(user)

    response = client.post(
        f"api/v1/doctors/{new_doctor.id}/unavailability",
        json={
            "date": str(datetime.date(2026, 8, 12)),
        },
        headers=headers,
    )

    assert response.status_code == 403
    assert response.json() == {
        "detail": "Insufficient permissions for this operation"}
    assert response.headers.get("WWW-Authenticate") is None
    retrieved_unavailability = doctor_repository.get_doctor_unavailability_by_date(
        session=session,
        doctor_id=new_doctor.id,
        target_date=datetime.date(2026, 8, 12),
    )

    assert retrieved_unavailability is None


def test_create_unavailability_requires_authentication(client, session, new_doctor):
    """Tests the POST /doctors/{doctor_id}/unavailability route with no headers"""
    response = client.post(
        f"api/v1/doctors/{new_doctor.id}/unavailability",
        json={
            "date": str(datetime.date(2026, 8, 12)),
        },
    )

    assert response.status_code == 401
    assert response.json() == {"detail": "Unauthorized"}
    assert response.headers.get("WWW-Authenticate") == "Bearer"
    retrieved_unavailability = doctor_repository.get_doctor_unavailability_by_date(
        session=session,
        doctor_id=new_doctor.id,
        target_date=datetime.date(2026, 8, 12),
    )

    assert retrieved_unavailability is None


def test_department_admin_can_create_unavailability(
    client,
    new_doctor,
    department_admin_headers,
    session,
):
    """Tests the POST /doctors/{doctor_id}/unavailability route"""
    response = client.post(
        f"api/v1/doctors/{new_doctor.id}/unavailability",
        json={
            "date": str(datetime.date(2026, 8, 12)),
        },
        headers=department_admin_headers,
    )

    assert response.status_code == 201
    data = response.json()
    assert data["date"] == str(datetime.date(2026, 8, 12))
    assert data["doctor_id"] == str(new_doctor.id)
    assert "id" in data
    assert "created_at" in data
    assert "updated_at" in data
    retrieved_unavailability = doctor_repository.get_doctor_unavailability_by_date(
        session=session,
        doctor_id=new_doctor.id,
        target_date=datetime.date(2026, 8, 12)
    )
    assert retrieved_unavailability is not None
    assert str(retrieved_unavailability.id) == data["id"]


@pytest.mark.parametrize(
    "role",
    [
        UserRole.VIEWER,
        UserRole.SUPER_ADMIN,
    ]
)
def test_non_doctor_or_department_admin_cannot_list_unavailability(
    client,
    new_doctor,
    role,
    user_factory,
    auth_headers_factory,
    department,
):
    """Tests the get /doctors/{doctor_id}/unavailability route with non allowed headers"""
    user = user_factory(
        role=role,
        department_id=(
            None
            if role == UserRole.SUPER_ADMIN
            else department.id
        ),
        doctor_id=None,
    )
    headers = auth_headers_factory(user)

    response = client.get(
        f"api/v1/doctors/{new_doctor.id}/unavailability",
        headers=headers,
    )

    assert response.status_code == 403
    assert response.json() == {
        "detail": "Insufficient permissions for this operation"}
    assert response.headers.get("WWW-Authenticate") is None


def test_list_unavailability_requires_authentication(client, new_doctor):
    """Tests the GET /doctors/{doctor_id}/unavailability route with no headers"""
    response = client.get(
        f"api/v1/doctors/{new_doctor.id}/unavailability",
    )

    assert response.status_code == 401
    assert response.json() == {"detail": "Unauthorized"}
    assert response.headers.get("WWW-Authenticate") == "Bearer"


@pytest.mark.parametrize(
    "role",
    [
        UserRole.SUPER_ADMIN,
        UserRole.VIEWER,
        UserRole.DOCTOR,
    ]
)
def test_non_department_admin_cannot_create_doctor_position(
    client,
    user_factory,
    auth_headers_factory,
    new_doctor,
    department,
    role,
    session,
    position,
):
    """Tests the POST /doctors/{doctor_id}/position route with non allowed headers"""
    user = user_factory(
        role=role,
        department_id=(
            None if role == UserRole.SUPER_ADMIN
            else department.id
        ),
        doctor_id=(
            new_doctor.id if role == UserRole.DOCTOR
            else None
        )
    )

    headers = auth_headers_factory(user)

    response = client.post(
        f"api/v1/doctors/{new_doctor.id}/position",
        json={
            "position_id": str(position.id),
        },
        headers=headers,
    )

    assert response.status_code == 403
    assert response.json() == {
        "detail": "Insufficient permissions for this operation"
    }
    assert response.headers.get("WWW-Authenticate") is None
    assert doctor_repository.get_doctor_position_by_id(
        doctor_id=new_doctor.id,
        position_id=position.id,
        session=session,
    ) is None


def test_create_doctor_position_requires_authentication(client, new_doctor, position, session):
    """Tests the POST /doctors/{doctor_id}/position route with no headers"""
    response = client.post(
        f"api/v1/doctors/{new_doctor.id}/position",
        json={
            "position_id": str(position.id),
        },
    )

    assert response.status_code == 401
    assert response.json() == {"detail": "Unauthorized"}
    assert response.headers.get("WWW-Authenticate") == "Bearer"
    assert doctor_repository.get_doctor_position_by_id(
        doctor_id=new_doctor.id,
        position_id=position.id,
        session=session,
    ) is None


def test_list_doctor_positions_requires_authentication(client, new_doctor):
    """Tests the GET /api/v1/doctors/{doctor_id}/position with no headers"""
    response = client.get(
        f"api/v1/doctors/{new_doctor.id}/position"
    )

    assert response.status_code == 401
    assert response.json() == {"detail": "Unauthorized"}
    assert response.headers.get("WWW-Authenticate") == "Bearer"
