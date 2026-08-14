"""
Tests for the shift module
"""

import uuid
import datetime
import pytest
from fastapi.testclient import TestClient
from sqlalchemy.pool import StaticPool
from sqlmodel import SQLModel, create_engine, Session

from src.main import app
from src.shift.schemas import ShiftCreate, ShiftAssignmentCreate
from src.shift.models import Shift as ShiftModel
from src.shift.models import ShiftAssignment as ShiftAssignmentModel
from src.shift import repository as shift_repository
from src.shift import controllers as shift_controllers
from src.position.schemas import PositionCreate
from src.position import repository as position_repository
from src.doctor.models import Doctor as DoctorModel
from src.doctor.schemas import DoctorCreate
from src.doctor import repository as doctor_repository
from src.department.schemas import DepartmentCreate
from src.department import repository as department_repository
from src.team.schemas import TeamCreate
from src.team import repository as team_repository

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


def test_get_shift_assignment_by_date(session, shift_assignment):
    """Tests retrieving a shift assignemtnt by its date"""
    retrieved_shift_assignment = shift_repository.get_shift_assignment_by_date(
        session=session,
        shift_id=shift_assignment.shift_id,
        target_date=datetime.date(2026, 8, 12)
    )

    assert isinstance(retrieved_shift_assignment.id, uuid.UUID)
    assert shift_assignment.id == retrieved_shift_assignment.id


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


#####################
# Route Tests
#####################


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


def test_create_shift_route(client, position):
    """Tests post /api/v1/shifts route"""
    response = client.post(
        "api/v1/shifts",
        json={
            "name": "ER 1",
            "doctors_per_shift": 1,
            "grants_day_off": False,
            "position_id": str(position.id)
        }
    )

    assert response.status_code == 201
    data = response.json()
    assert data["name"] == "ER 1"
    assert data["position_id"] == str(position.id)
    assert "id" in data
    assert "created_at" in data
    assert "updated_at" in data


def test_create_shift_route_invalid_payload(client, position):
    """Tests post /api/v1/shifts route with invalid payload"""
    response = client.post(
        "api/v1/shifts",
        json={
            "doctors_per_shift": 1,
            "grants_day_off": False,
            "position_id": str(position.id)
        }
    )

    assert response.status_code == 422


def test_list_shifts_route(client, shift):
    """Tests get /api/v1/shifts route"""
    response = client.get("/api/v1/shifts")

    assert response.status_code == 200
    data = response.json()
    assert isinstance(data, list)
    assert data[0]["id"] == str(shift.id)


def test_get_shift_route(client, shift):
    """Tests get /api/v1/shifts/{shift_id} route"""
    response = client.get(f"/api/v1/shifts/{shift.id}")

    assert response.status_code == 200
    data = response.json()
    assert data["id"] == str(shift.id)


def test_get_shift_route_nonexistent_id(client):
    """Tests get /api/v1/shifts/{shift_id} route with invalid payload"""
    response = client.get(f"/api/v1/shifts/{uuid.uuid4()}")

    assert response.status_code == 404


def test_create_shift_assignment_route(client, shift, new_doctor):
    """Tests post /api/v1/shifts/{shift_id}/assignments"""
    response = client.post(
        f"/api/v1/shifts/{shift.id}/assignments",
        json={
            "doctor_id": str(new_doctor.id),
            "date": str(datetime.date(2026, 8, 12))
        }
    )

    assert response.status_code == 201
    data = response.json()
    assert data["doctor_id"] == str(new_doctor.id)
    assert "id" in data
    assert "created_at" in data
    assert "updated_at" in data


def test_create_shift_assignment_route_invalid_payload(client, shift, new_doctor):
    """Tests post /api/v1/shifts/{shift_id}/assignments with invalid payload"""
    response = client.post(
        f"/api/v1/shifts/{shift.id}/assignments",
        json={
            "doctor_id": str(new_doctor.id),
        }
    )

    assert response.status_code == 422


def test_create_shift_assignment_route_duplicate(client,shift, new_doctor):
    """Tests post /api/v1/shifts/{shift_id}/assignments with invalid payload"""
    client.post(
        f"/api/v1/shifts/{shift.id}/assignments",
        json={
            "doctor_id": str(new_doctor.id),
            "date": str(datetime.date(2026, 8, 12))
        }
    )

    response = client.post(
        f"/api/v1/shifts/{shift.id}/assignments",
        json={
            "doctor_id": str(new_doctor.id),
            "date": str(datetime.date(2026, 8, 12))
        }
    )

    assert response.status_code == 400


def tests_list_shift_assignments_route(client, shift_assignment):
    """Tests get /api/v1/shifts/assignments"""
    response = client.get("api/v1/shifts/assignments")

    assert response.status_code == 200
    data = response.json()
    assert isinstance(data, list)
    assert data[0]["id"] == str(shift_assignment.id)
