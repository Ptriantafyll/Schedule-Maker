"""
Unit tests for db models defined in db_models.py. These tests verify that the SQLModel models correctly implement the expected fields, relationships, and behaviors. The tests use an in-memory SQLite database to ensure isolation and repeatability.
"""

import datetime
import pytest
import uuid
from sqlmodel import SQLModel, create_engine, Session
from src.department.models import Department


@pytest.fixture(name="session")
def session_fixture():
    """Creates a fresh in-memory database session for each test."""
    # Import all models to ensure SQLModel has registered them before create_all
    from src.db.schemas import ScheduleConfig 

    from src.department.models import Department
    from src.team.models import Team
    from src.doctor.models import Doctor, DoctorUnavailability, DoctorPreAssignment, DoctorPosition
    from src.shift.models import Shift, ShiftAssignment
    from src.position.models import Position

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

    # Verify SyncBase properties
    assert isinstance(dept.id, uuid.UUID)
    assert dept.is_deleted is False
    assert dept.sync_status is False
    assert isinstance(dept.created_at, datetime.datetime)
    assert isinstance(dept.updated_at, datetime.datetime)
