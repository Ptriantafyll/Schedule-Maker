"""
Tests for the doctor module
"""

import uuid
import datetime
import pytest
from fastapi.testclient import TestClient
# Session fixture for database tests
from sqlalchemy.pool import StaticPool
from sqlmodel import SQLModel, create_engine, Session


from src.main import app
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
from src.doctor.schemas import DoctorCreate, DoctorPreAssignmentCreate, DoctorUnavailabilityCreate, DoctorPositionCreate
from src.doctor import repository as doctor_repository
from src.doctor import controllers as doctor_controllers


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
    doctor: DoctorModel
):
    """Helper that creates a new doctor pre assignment in the db"""
    pre_assignment_data = DoctorPreAssignmentCreate(
        date=datetime.date(2026, 8, 12),
        shift_id=shift_id
    )

    return doctor_repository.create_doctor_pre_assignment(
        session, doctor.id, pre_assignment_data
    )


def create_test_unavailability(
    session: Session,
    doctor: DoctorModel
):
    """Helper that creates a new doctor pre assignment in the db"""
    unavailability_data = DoctorUnavailabilityCreate(
        date=datetime.date(2026, 8, 12)
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


@pytest.fixture(name="department")
def department_fixture(session):
    """Creates a reusable department for tests"""
    dept_data = DepartmentCreate(name="Cardiology", code="CARD")
    return department_repository.create_department(session, dept_data)


@pytest.fixture(name="team")
def team_fixture(session, department):
    """Creates a reusable team for tests"""
    team_data = TeamCreate(name="ER Team A", department_id=department.id)
    return create_team(session, team_data)


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
    return create_test_pre_assignment(session, shift.id, new_doctor)


@pytest.fixture(name="unavailability")
def unavailability_fixture(session, new_doctor):
    """Creates a reusable pre-assignment for tests"""
    return create_test_unavailability(session, new_doctor)

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
    """Test retrieving a doctor by name"""
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


def test_doctor_has_department_foreign_key(session):
    """Test that doctors can be associated with a department and retrieved correctly."""
    # """Test that doctors can be associated with a department and retrieved correctly."""
    # dept = Department(name="Emergency", code="ER")
    # session.add(dept)
    # session.commit()
    # session.refresh(dept)

    # team = Team(name="ER Team A", department_id=dept.id)
    # session.add(team)
    # session.commit()
    # session.refresh(team)

    # doctor = Doctor(name="Dr. Smith", email="smith@test.com",
    #                 department_id=dept.id, team_id=team.id)
    # session.add(doctor)
    # session.commit()
    # session.refresh(doctor)

    # assert isinstance(doctor.id, uuid.UUID)
    # assert doctor.department_id == dept.id
    # assert doctor.name == "Dr. Smith"
    # assert doctor.team_id == team.id
    # assert doctor.email == "smith@test.com"

#######################
# Controller tests
#######################


def test_create_doctor_controller_duplicate_name(session, department, team, new_doctor):
    """Test that creating a doctor with a duplicate email raises an error."""
    # new_doctor is needed in the input to create the first (duplicate) doctor
    doctor_data = DoctorCreate(
        name="Dr Panos",
        email="drpanos@gmail.com",
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


def test_get_doctor_controller_deleted(session):
    """Test that retrieving a deleted team raises a 404 error"""
    pass


#######################
# Route tests
#######################
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


def test_create_doctor_route(client, department, team):
    """Test the POST /doctors/ route for creating a doctor"""

    response = client.post(
        "api/v1/doctors",
        json={
            "name": "Dr Panos",
            "email": "drpanos@gmail.com",
            "department_id": str(department.id),
            "team_id": str(team.id)
        }
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


def test_create_doctor_route_invalid_payload(client, department, team):
    """Test the POST /doctors/ route rejects invalid payload"""

    response = client.post(
        "api/v1/doctors",
        json={
            "name": "Dr Panos",
            "department_id": str(department.id),
            "team_id": str(team.id),
        }
    )
    assert response.status_code == 422


def test_get_doctor_by_id_route(client, department, team, new_doctor):
    """Tests that the GET /doctors/{doctor_id} route returns a doctor"""
    response = client.get(
        f"/api/v1/doctors/{new_doctor.id}"
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


# def test_get_doctor_by_id_route_nonexistent(client, team, department):
#     """Test the GET /doctors/{doctor_id} route returns error when given a nonexistent id"""
#     pass


# new doctor is needed to add a doctor in the db
def test_list_doctors_route(client, new_doctor):
    """Tests the GET /doctors/ route"""
    response = client.get("api/v1/doctors")

    assert response.status_code == 200


def test_create_doctor_pre_assignments_route(client, shift, new_doctor):
    """Tests the POST /doctors/{doctor_id}/pre-assignments route"""

    response = client.post(
        f"api/v1/doctors/{new_doctor.id}/pre-assignments",
        json={
            "date": str(datetime.date(2026, 8, 12)),
            "doctor_id": str(new_doctor.id),
            "shift_id": str(shift.id)
        }
    )

    assert response.status_code == 201
    data = response.json()
    assert data["date"] == str(datetime.date(2026, 8, 12))
    assert data["doctor_id"] == str(new_doctor.id)
    assert data["shift_id"] == str(shift.id)
    assert "id" in data
    assert "created_at" in data
    assert "updated_at" in data


def test_create_doctor_pre_assignments_route_invalid_payload(client,  new_doctor):
    """Tests the POST /doctors/{doctor_id}/pre-assignments route rejects invalid payload"""
    response = client.post(
        f"api/v1/doctors/{new_doctor.id}/pre-assignments",
        json={
            "date": str(datetime.date(2026, 8, 12)),
            "doctor_id": str(new_doctor.id),
        }
    )
    assert response.status_code == 422


def test_get_doctor_pre_assigments_route():
    """Tests the GET /doctors/{doctor_id}/pre-assignments route"""
    # TODO


def test_create_doctor_unavailability_route(client, new_doctor):
    """Tests the POST /doctors/{doctor_id}/unavailability route"""
    response = client.post(
        f"api/v1/doctors/{new_doctor.id}/unavailability",
        json={
            "date": str(datetime.date(2026, 8, 12)),
        }
    )

    assert response.status_code == 201
    data = response.json()
    assert data["date"] == str(datetime.date(2026, 8, 12))
    assert data["doctor_id"] == str(new_doctor.id)
    assert "id" in data
    assert "created_at" in data
    assert "updated_at" in data


def test_create_doctor_unavailability_route_invalid_payload(client, new_doctor):
    """Tests the POST /doctors/{doctor_id}/unavailability route rejects invalid payload"""
    response = client.post(
        f"api/v1/doctors/{new_doctor.id}/unavailability",
        json={}
    )
    assert response.status_code == 422


def test_get_doctor_unavailability_route():
    """Tests the GET /doctors/{doctor_id}/unavailability route"""
    # TODO


def test_create_doctor_position_route():
    """Tests the POST /doctors/{doctor_id}/position route"""
    # TODO


def test_create_doctor_position_route_invalid_payload():
    """Tests the POST /doctors/{doctor_id}/position route rejects invalid payload"""
    # TODO


def test_get_doctor_position_route():
    """Tests the GET /doctors/{doctor_id}/position route"""
    # TODO
