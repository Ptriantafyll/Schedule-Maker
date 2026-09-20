"""
Unit tests for db models defined in db_models.py. These tests verify that the SQLModel models correctly implement the expected fields, relationships, and behaviors. The tests use an in-memory SQLite database to ensure isolation and repeatability.
"""

import datetime
import uuid
import pytest
from sqlalchemy import text
from sqlalchemy.exc import IntegrityError
from sqlmodel import SQLModel, create_engine, Session

import src.db.connection  # noqa: F401 - ensures SQLite PRAGMA event listener is registered
from src.department.models import Department
from src.position.models import Position
from src.team.models import Team
from src.doctor.models import (
    Doctor,
    DoctorPosition,
    DoctorPreAssignment,
    DoctorUnavailability,
)
from src.shift.models import Shift, ShiftAssignment
from src.user.models import User, UserRole


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

    # Verify SyncBase properties
    assert isinstance(dept.id, uuid.UUID)
    assert dept.is_deleted is False
    assert dept.sync_status is False
    assert isinstance(dept.created_at, datetime.datetime)
    assert isinstance(dept.updated_at, datetime.datetime)


def test_doctor_unavailability_unique_constraint(session):
    """Verify that duplicate (doctor_id, date) in DoctorUnavailability raises IntegrityError."""
    dept = Department(name="Cardiology", code="CARD")
    session.add(dept)
    session.commit()

    doc = Doctor(name="Dr. Smith", email="smith@example.com", department_id=dept.id)
    session.add(doc)
    session.commit()

    target_date = datetime.date(2026, 10, 1)
    unavail1 = DoctorUnavailability(doctor_id=doc.id, date=target_date)
    unavail2 = DoctorUnavailability(doctor_id=doc.id, date=target_date)

    session.add(unavail1)
    session.commit()

    session.add(unavail2)
    with pytest.raises(IntegrityError):
        session.commit()


def test_doctor_pre_assignment_unique_constraint(session):
    """Verify that duplicate (doctor_id, date) in DoctorPreAssignment raises IntegrityError."""
    dept = Department(name="Cardiology", code="CARD")
    session.add(dept)
    session.commit()

    doc = Doctor(name="Dr. Smith", email="smith@example.com", department_id=dept.id)
    pos = Position(name="Attending", department_id=dept.id)
    session.add_all([doc, pos])
    session.commit()

    shift1 = Shift(name="Morning", position_id=pos.id)
    shift2 = Shift(name="Night", position_id=pos.id)
    session.add_all([shift1, shift2])
    session.commit()

    target_date = datetime.date(2026, 10, 1)
    pre1 = DoctorPreAssignment(doctor_id=doc.id, shift_id=shift1.id, date=target_date)
    pre2 = DoctorPreAssignment(doctor_id=doc.id, shift_id=shift2.id, date=target_date)

    session.add(pre1)
    session.commit()

    session.add(pre2)
    with pytest.raises(IntegrityError):
        session.commit()


def test_doctor_position_unique_constraint(session):
    """Verify that duplicate (doctor_id, position_id) in DoctorPosition raises IntegrityError."""
    dept = Department(name="Cardiology", code="CARD")
    session.add(dept)
    session.commit()

    doc = Doctor(name="Dr. Smith", email="smith@example.com", department_id=dept.id)
    pos = Position(name="Attending", department_id=dept.id)
    session.add_all([doc, pos])
    session.commit()

    doc_pos1 = DoctorPosition(doctor_id=doc.id, position_id=pos.id)
    doc_pos2 = DoctorPosition(doctor_id=doc.id, position_id=pos.id)

    session.add(doc_pos1)
    session.commit()

    session.add(doc_pos2)
    with pytest.raises(IntegrityError):
        session.commit()


def test_shift_assignment_unique_constraint(session):
    """Verify that duplicate (doctor_id, date) in ShiftAssignment raises IntegrityError."""
    dept = Department(name="Cardiology", code="CARD")
    session.add(dept)
    session.commit()

    doc = Doctor(name="Dr. Smith", email="smith@example.com", department_id=dept.id)
    pos = Position(name="Attending", department_id=dept.id)
    session.add_all([doc, pos])
    session.commit()

    shift1 = Shift(name="Morning", position_id=pos.id)
    shift2 = Shift(name="Night", position_id=pos.id)
    session.add_all([shift1, shift2])
    session.commit()

    target_date = datetime.date(2026, 10, 1)
    assign1 = ShiftAssignment(doctor_id=doc.id, shift_id=shift1.id, date=target_date)
    assign2 = ShiftAssignment(doctor_id=doc.id, shift_id=shift2.id, date=target_date)

    session.add(assign1)
    session.commit()

    session.add(assign2)
    with pytest.raises(IntegrityError):
        session.commit()


def test_sqlite_foreign_keys_pragma_enabled(session):
    """Verify that every SQLite connection has foreign-key enforcement enabled."""
    result = session.exec(text("PRAGMA foreign_keys;")).scalar()
    assert result == 1


def test_sqlite_foreign_key_rejects_nonexistent_parent(session):
    """Verify that SQLite rejects records that reference nonexistent parent records."""
    nonexistent_dept_id = uuid.uuid4()
    doc = Doctor(
        name="Dr. Orphan",
        email="orphan@example.com",
        department_id=nonexistent_dept_id,
    )
    session.add(doc)
    with pytest.raises(IntegrityError):
        session.commit()


def test_nullable_foreign_key_accepts_null(session):
    """Verify that nullable foreign keys accept NULL, including tenantless super-admin."""
    super_admin = User(
        email="superadmin@example.com",
        hashed_password="hashed_secret_password",
        full_name="Super Admin",
        role=UserRole.SUPER_ADMIN,
        department_id=None,
        doctor_id=None,
    )
    session.add(super_admin)
    session.commit()
    session.refresh(super_admin)

    assert super_admin.id is not None
    assert super_admin.department_id is None
    assert super_admin.doctor_id is None
