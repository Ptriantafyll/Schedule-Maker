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
    