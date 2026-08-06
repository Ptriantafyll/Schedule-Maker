"""
Tests for the doctor module
"""

from fastapi.testclient import TestClient
import uuid
import datetime
from src.main import app

from src.department.schemas import DepartmentCreate
from src.department.repository import create_department
from src.team.schemas import TeamCreate
from src.team.repository import create_team
from src.doctor.schemas import DoctorCreate
from src.db.schemas import Doctor
from src.department.models import Department
from src.team.models import Team

# Session fixture for database tests
import pytest
from sqlalchemy.pool import StaticPool
from sqlmodel import SQLModel, create_engine, Session


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
    return create_department(session, dept_data)


@pytest.fixture(name="team")
def team_fixture(session, department):
    """Creates a reusable team for tests"""
    team_data = TeamCreate(name="ER Team A", department_id=department.id)
    return create_team(session, team_data)

#####################
# Repository tests
#####################


def test_create_doctor(session):
    """Test creating a doctor and verifying their fields"""
    doctor_data = DoctorCreate()

def test_get_doctor_by_name(session):
    """Test retrieving a doctor by name"""

def test_get_doctor_by_id(session):
    """Test retrieving a doctor by id"""

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
    