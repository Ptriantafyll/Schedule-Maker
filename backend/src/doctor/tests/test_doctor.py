"""
Tests for the doctor module
"""

from fastapi.testclient import TestClient
import uuid
import datetime
import pytest
# Session fixture for database tests
from sqlalchemy.pool import StaticPool
from sqlmodel import SQLModel, create_engine, Session


from src.main import app
from src.department.schemas import DepartmentCreate
from src.department import repository as department_repository
from src.team.schemas import TeamCreate
from src.team.repository import create_team
from src.doctor.schemas import DoctorCreate
from src.doctor import repository as doctor_repository


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

#####################
# Repository tests
#####################


def test_create_doctor(session, department, team):
    """Test creating a doctor and verifying their fields"""
    doctor_data = DoctorCreate(
        name="Dr Panos",
        email="drpanos@gmail.com",
        department_id=department.id,
        team_id=team.id
    )
    new_doctor = doctor_repository.create_doctor(session, doctor_data)

    assert isinstance(new_doctor.id, uuid.UUID)
    assert new_doctor.name == "Dr Panos"
    assert new_doctor.email == "drpanos@gmail.com"
    assert new_doctor.department_id == department.id
    assert new_doctor.team_id == team.id
    assert new_doctor.is_deleted is False
    assert new_doctor.sync_status is False
    assert isinstance(new_doctor.created_at, datetime.datetime)
    assert isinstance(new_doctor.updated_at, datetime.datetime)


def test_get_doctor_by_email(session, department, team):
    """Test retrieving a doctor by name"""

    doctor_data = DoctorCreate(
        name="Dr Panos",
        email="drpanos@gmail.com",
        department_id=department.id,
        team_id=team.id
    )
    new_doctor = doctor_repository.create_doctor(session, doctor_data)

    retrieved_doctor = doctor_repository.get_doctor_by_email(session, "drpanos@gmail.com")

    assert retrieved_doctor is not None
    assert retrieved_doctor.id == new_doctor.id

def test_get_doctor_by_id(session, department, team):
    """Test retrieving a doctor by id"""

    doctor_data = DoctorCreate(
        name="Dr Panos",
        email="drpanos@gmail.com",
        department_id=department.id,
        team_id=team.id
    )
    new_doctor = doctor_repository.create_doctor(session, doctor_data)

    retrieved_doctor = doctor_repository.get_doctor_by_id(session, new_doctor.id)

    assert retrieved_doctor is not None
    assert retrieved_doctor.id == new_doctor.id


def test_get_active_doctors(session):
    """Test retrieving all active doctors"""


def test_create_doctor_pre_assignment(session):
    """Test creating pre assignments for a doctor"""


def test_get_doctor_pre_assignment(session):
    """Test retrieving a doctor's pre assignments"""


def test_create_doctor_unavailability(session):
    """Test creating unavailability dates for a doctor"""


def test_get_doctor_unavailability(session):
    """Test retrieving unavailability dates for a doctor"""


def test_create_doctor_position(session):
    """Test creating a position for a doctor"""


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

    # response = client.post(
    #     "api/v1/doctors",
    #     json={
    #         "name": "Dr Panos",
    #         "email": "drpanos@gmeil.com",
    #         "department_id": department.id,
    #         "team_id": team.id
    #     }
    # )

    # assert response.status_code == 201
    # data = response.json()
    # assert data["name"] == "Dr Panos"
    # assert data["email"] == "drpanos@gmeil.com"
    # assert data["department_id"] == department.id
    # assert data["team_id"] == team.id
    # assert "id" in data
    # assert "created_at" in data
    # assert "updated_at" in data


def test_create_doctor_route_invalid_payload(client, department, team):
    """Test the POST /doctors/ route rejects invalid payload"""

    # response = client.post(
    #     "api/v1/doctors",
    #     json={
    #         "name": "Dr Panos",
    #         "department_id": department.id,
    #         "team_id": team.id,
    #     }
    # )
    # assert response.status_code == 422


def test_get_doctor_by_id_route(client, department, team, session):
    """Tests that the GET /doctors/{doctor_id} route returns a doctor"""

    # doctor_data = DoctorCreate(
    #     name="Dr Panos",
    #     email="drpanos@gmail.com",
    #     department_id=department.id,
    #     team_id=team.id
    # )

    # new_doctor = create_doctor(session, doctor_data)

    # response = client.get(
    #     f"/api/v1/doctors/{new_doctor.id}"
    # )

    # assert response.status_code == 200
    # data = response.json()
    # assert data["name"] == "Dr Panos"
    # assert data["email"] == "drpanos@gmeil.com"
    # assert data["department_id"] == department.id
    # assert data["team_id"] == team.id
    # assert "id" in data
    # assert "created_at" in data
    # assert "updated_at" in data


def test_get_doctor_by_id_route_nonexistent(client):
    """Test the GET /doctors/{doctor_id} route returns error when given a nonexistent id"""
    # TODO


def test_list_doctors_route():
    """Tests the GET /doctors/ route"""
    # TODO


def test_create_doctor_pre_assignments_route():
    """Tests the POST /doctors/{doctor_id}/pre-assignments route"""
    # TODO


def test_create_doctor_pre_assignments_route_invalid_payload():
    """Tests the POST /doctors/{doctor_id}/pre-assignments route rejects invalid payload"""
    # TODO


def test_get_doctor_pre_assigments_route():
    """Tests the GET /doctors/{doctor_id}/pre-assignments route"""
    # TODO


def test_create_doctor_unavailability_route():
    """Tests the POST /doctors/{doctor_id}/unavailability route"""
    # TODO


def test_create_doctor_unavailability_route_invalid_payload():
    """Tests the POST /doctors/{doctor_id}/unavailability route rejects invalid payload"""
    # TODO


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
