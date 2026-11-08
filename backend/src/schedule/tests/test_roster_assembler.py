"""
Tests for assembling solver Department domain graphs from database roster entities.
"""

import datetime
import uuid
import pytest

from src.department.schemas import DepartmentCreate
from src.department import repository as department_repository
from src.team import repository as team_repository
from src.position import repository as position_repository
from src.shift import repository as shift_repository
from src.shift.schemas import ShiftCreate
from src.doctor import repository as doctor_repository
from src.doctor.schemas import (
    DoctorPositionCreate,
    DoctorUnavailabilityCreate,
    DoctorPreAssignmentCreate,
)
from src.user.models import UserRole
from src.schedule.roster_assembler import assemble_department_from_db


@pytest.fixture(name="roster_setup")
def roster_setup_fixture(session, user_factory):
    """Sets up a realistic department roster with teams, teamless doctors, multiple positions, and shifts."""
    dept = department_repository.create_department(
        session, DepartmentCreate(name="Emergency Dept", code="ED")
    )
    team = team_repository.create_team(
        session, name="Team Red", department_id=dept.id
    )

    # 2 Positions: Emergency has Day and Night; ICU has Duty
    pos_er = position_repository.create_position(
        session, position_name="Emergency", department_id=dept.id, duty_days=[0, 1, 2, 3, 4, 5, 6]
    )
    pos_icu = position_repository.create_position(
        session, position_name="ICU", department_id=dept.id, duty_days=[0, 1, 2, 3, 4, 5, 6]
    )

    shift_er_day = shift_repository.create_shift(
        session,
        ShiftCreate(name="Day", position_id=pos_er.id, doctors_per_shift=1, grants_day_off=True),
    )
    shift_er_night = shift_repository.create_shift(
        session,
        ShiftCreate(name="Night", position_id=pos_er.id, doctors_per_shift=1, grants_day_off=True),
    )
    shift_icu_duty = shift_repository.create_shift(
        session,
        ShiftCreate(name="Duty", position_id=pos_icu.id, doctors_per_shift=1, grants_day_off=False),
    )

    # Doctors: Alice (Team Red, ER + ICU), Bob (Teamless, ER)
    doc_alice = doctor_repository.create_doctor(
        session, name="Dr Alice", department_id=dept.id, team_id=team.id
    )
    doc_bob = doctor_repository.create_doctor(
        session, name="Dr Bob", department_id=dept.id, team_id=None
    )

    # User account for Alice, none for Bob
    user_factory(
        email="alice@hospital.com",
        role=UserRole.DOCTOR,
        department_id=dept.id,
        doctor_id=doc_alice.id,
    )

    # Position assignments
    doctor_repository.create_doctor_position(
        session, doc_alice.id, DoctorPositionCreate(position_id=pos_er.id)
    )
    doctor_repository.create_doctor_position(
        session, doc_alice.id, DoctorPositionCreate(position_id=pos_icu.id)
    )
    doctor_repository.create_doctor_position(
        session, doc_bob.id, DoctorPositionCreate(position_id=pos_er.id)
    )

    # Alice unavailabilities and pre-assignments for 2026-11
    doctor_repository.create_doctor_unavailability(
        session,
        doc_alice.id,
        DoctorUnavailabilityCreate(date=datetime.date(2026, 11, 10)),
    )
    doctor_repository.create_doctor_pre_assignment(
        session,
        doc_alice.id,
        DoctorPreAssignmentCreate(shift_id=shift_er_day.id, date=datetime.date(2026, 11, 15)),
    )

    return {
        "dept": dept,
        "team": team,
        "pos_er": pos_er,
        "pos_icu": pos_icu,
        "shift_er_day": shift_er_day,
        "shift_er_night": shift_er_night,
        "shift_icu_duty": shift_icu_duty,
        "doc_alice": doc_alice,
        "doc_bob": doc_bob,
    }


def test_assemble_department_from_db_success(session, roster_setup):
    """Tests assembling a complete domain graph from the database roster."""
    dept = roster_setup["dept"]
    domain_dept = assemble_department_from_db(session, dept.id, "2026-11")

    assert domain_dept.name == "Emergency Dept"
    assert len(domain_dept.positions) == 2

    # Check positions and multi-shift
    positions_by_name = {p.name: p for p in domain_dept.positions}
    er_pos = positions_by_name["Emergency"]
    icu_pos = positions_by_name["ICU"]

    assert len(er_pos.shifts) == 2
    assert {s.name for s in er_pos.shifts} == {"Day", "Night"}
    assert len(icu_pos.shifts) == 1
    assert icu_pos.shifts[0].name == "Duty"

    # Check position doctor eligibility (Alice in both, Bob only in ER)
    assert {d.name for d in er_pos.eligible_doctors} == {"Dr Alice", "Dr Bob"}
    assert {d.name for d in icu_pos.eligible_doctors} == {"Dr Alice"}

    # Check teams vs teamless doctors
    assert len(domain_dept.teams) == 1
    assert domain_dept.teams[0].name == "Team Red"
    assert [d.name for d in domain_dept.teams[0].doctors] == ["Dr Alice"]
    assert [d.name for d in domain_dept.teamless_doctors] == ["Dr Bob"]
    assert len(domain_dept.doctors) == 2

    # Check email resolution & fallback
    doctors_by_name = {d.name: d for d in domain_dept.doctors}
    assert doctors_by_name["Dr Alice"].email == "alice@hospital.com"
    assert doctors_by_name["Dr Bob"].email == ""

    # Check unavailability & pre-assignment mapping
    assert doctors_by_name["Dr Alice"].unavailability == {datetime.date(2026, 11, 10)}
    assert len(doctors_by_name["Dr Alice"].pre_assignments) == 1
    date, shift = doctors_by_name["Dr Alice"].pre_assignments[0]
    assert date == datetime.date(2026, 11, 15)
    assert shift.name == "Day"


def test_assemble_department_from_db_missing_department_raises(session):
    """Tests that querying a nonexistent department raises ValueError."""
    fake_id = uuid.uuid4()
    with pytest.raises(ValueError, match=f"Department {fake_id} not found"):
        assemble_department_from_db(session, fake_id, "2026-11")


def test_assemble_department_from_db_invalid_month_format_raises(session, roster_setup):
    """Tests that malformed target month strings raise ValueError."""
    dept = roster_setup["dept"]
    with pytest.raises(ValueError, match="Invalid target month format"):
        assemble_department_from_db(session, dept.id, "2026/11")


def test_assemble_department_from_db_no_positions_raises(session):
    """Tests that a department without active positions raises ValueError."""
    dept = department_repository.create_department(
        session, DepartmentCreate(name="Empty Dept", code="EMPTY")
    )
    with pytest.raises(ValueError, match="No active positions found"):
        assemble_department_from_db(session, dept.id, "2026-11")


def test_assemble_department_from_db_position_without_shifts_raises(session):
    """Tests that an active position with 0 shifts raises ValueError."""
    dept = department_repository.create_department(
        session, DepartmentCreate(name="No Shifts Dept", code="NOSHIFT")
    )
    position_repository.create_position(
        session, position_name="Clinic", department_id=dept.id, duty_days=[0, 1]
    )
    with pytest.raises(ValueError, match="Position 'Clinic' has no active shifts configured"):
        assemble_department_from_db(session, dept.id, "2026-11")


def test_assemble_department_from_db_no_doctors_raises(session):
    """Tests that a department with positions but no doctors raises ValueError."""
    dept = department_repository.create_department(
        session, DepartmentCreate(name="No Docs Dept", code="NODOC")
    )
    pos = position_repository.create_position(
        session, position_name="Clinic", department_id=dept.id, duty_days=[0, 1]
    )
    shift_repository.create_shift(
        session, ShiftCreate(name="Day", position_id=pos.id, doctors_per_shift=1)
    )
    with pytest.raises(ValueError, match="No active doctors found"):
        assemble_department_from_db(session, dept.id, "2026-11")


def test_assemble_department_from_db_position_without_eligible_doctors_raises(session):
    """Tests that a position with 0 eligible doctors raises ValueError."""
    dept = department_repository.create_department(
        session, DepartmentCreate(name="Unstaffed Dept", code="UNSTAFF")
    )
    pos = position_repository.create_position(
        session, position_name="Solo Position", department_id=dept.id, duty_days=[0, 1]
    )
    shift_repository.create_shift(
        session, ShiftCreate(name="Day", position_id=pos.id, doctors_per_shift=1)
    )
    # Create a doctor but don't assign them to Solo Position
    doctor_repository.create_doctor(session, name="Dr Free", department_id=dept.id)

    with pytest.raises(ValueError, match="Position 'Solo Position' has no eligible doctors assigned"):
        assemble_department_from_db(session, dept.id, "2026-11")


def test_assemble_department_from_db_unavailability_preassignment_conflict_raises(
    session, roster_setup
):
    """Tests that having an unavailability and pre-assignment on the same date raises ValueError."""
    dept = roster_setup["dept"]
    doc_bob = roster_setup["doc_bob"]
    shift_er_day = roster_setup["shift_er_day"]

    # Bob has unavailability on Nov 12
    doctor_repository.create_doctor_unavailability(
        session, doc_bob.id, DoctorUnavailabilityCreate(date=datetime.date(2026, 11, 12))
    )
    # Bob also has pre-assignment on Nov 12
    doctor_repository.create_doctor_pre_assignment(
        session,
        doc_bob.id,
        DoctorPreAssignmentCreate(shift_id=shift_er_day.id, date=datetime.date(2026, 11, 12)),
    )

    with pytest.raises(ValueError, match="cannot be pre-assigned on unavailable day"):
        assemble_department_from_db(session, dept.id, "2026-11")
