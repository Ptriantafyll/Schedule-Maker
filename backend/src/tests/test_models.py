"""
Unit tests for db models defined in db_models.py. These tests verify that the SQLModel models correctly implement the expected fields, relationships, and behaviors. The tests use an in-memory SQLite database to ensure isolation and repeatability. 
"""
import pytest
import uuid
import datetime
from sqlmodel import SQLModel, create_engine, Session
from src.department.models import Department
from src.db.schemas import Doctor, Team


@pytest.fixture(name="session")
def session_fixture():
    """Creates a fresh in-memory database session for each test."""
    engine = create_engine("sqlite:///:memory:")
    SQLModel.metadata.create_all(engine)
    with Session(engine) as session:
        yield session


def test_sync_base_fields(session):
    """Verify that models inheriting from SyncBase automatically get UUID and sync metadata."""
    dept = Department(name="Cardiology", code="CARD")
    session.add(dept)
    session.commit()
    session.refresh(dept)

    assert isinstance(dept.id, uuid.UUID)
    assert dept.is_deleted is False
    assert dept.sync_status is False
    assert isinstance(dept.created_at, datetime.datetime)
    assert isinstance(dept.updated_at, datetime.datetime)


def test_doctor_has_department_foreign_key(session):
    """Test that doctors can be associated with a department and retrieved correctly."""
    dept = Department(name="Emergency", code="ER")
    session.add(dept)
    session.commit()
    session.refresh(dept)

    team = Team(name="ER Team A", department_id=dept.id)
    session.add(team)
    session.commit()
    session.refresh(team)

    doctor = Doctor(name="Dr. Smith", email="smith@test.com",
                    department_id=dept.id, team_id=team.id)
    session.add(doctor)
    session.commit()
    session.refresh(doctor)

    assert isinstance(doctor.id, uuid.UUID)
    assert doctor.department_id == dept.id
    assert doctor.name == "Dr. Smith"
    assert doctor.email == "smith@test.com"
