"""
Tests for the doctor module
"""

from typing import Optional
import uuid
import datetime
import pytest
from fastapi import HTTPException
from pydantic import ValidationError
from sqlmodel import Session


from src.department.schemas import DepartmentCreate
from src.department import repository as department_repository
from src.team.repository import create_team
from src.shift import repository as shift_repository
from src.shift.schemas import ShiftCreate
from src.position import repository as position_repository
from src.position.models import Position as PositionModel
from src.doctor.models import Doctor as DoctorModel
from src.doctor.schemas import (
    DoctorCreate,
    DoctorRead,
    DoctorUpdate,
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
    email: Optional[str] = None,
    department_id: Optional[uuid.UUID] = None,
    team_id: Optional[uuid.UUID] = None,
    **kwargs,
) -> DoctorModel:
    """Helper that creates a new doctor in the db"""
    return doctor_repository.create_doctor(
        session=session,
        department_id=department_id,
        name=name,
        team_id=team_id,
    )


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


def create_test_doctor_position(
    session: Session,
    doctor: DoctorModel,
    position: PositionModel,
):
    """Helper that creates a doctor-position"""
    doctor_pos_data = DoctorPositionCreate(
        position_id=position.id
    )
    return doctor_repository.create_doctor_position(
        session=session,
        doctor_id=doctor.id,
        doctor_pos_data=doctor_pos_data,
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
    return create_team(
        session=session,
        name="ER Team A",
        department_id=department.id,
    )


@pytest.fixture(name="position")
def position_fixture(session, department):
    """Creates a reusable position for tests"""
    return position_repository.create_position(
        session=session,
        position_name="ER",
        department_id=department.id,
        duty_days=[1, 3, 5],
    )


@pytest.fixture(name="department_b")
def department_b_fixture(session):
    """Creates a second department for tenant-isolation tests."""
    dept_data = DepartmentCreate(name="Radiology", code="RAD")
    return department_repository.create_department(session, dept_data)


@pytest.fixture(name="team_b")
def team_b_fixture(session, department_b):
    """Creates a reusable team in Department B."""
    return create_team(
        session=session,
        name="ER Team A",
        department_id=department_b.id,
    )


@pytest.fixture(name="position_b")
def position_b_fixture(session, department_b):
    """Creates a reusable Position in Department B."""
    return position_repository.create_position(
        session=session,
        position_name="ER",
        department_id=department_b.id,
        duty_days=[1, 3, 5],
    )


@pytest.fixture(name="doctor_b")
def doctor_b_fixture(session, department_b, team_b):
    """Creates a reusable doctor in Department B."""
    return create_new_doctor(
        session=session,
        name="Dr Radiology",
        email="drradiology@gmail.com",
        department_id=department_b.id,
        team_id=team_b.id,
    )


@pytest.fixture(name="shift_b")
def shift_b_fixture(session, position_b):
    """Creates a reusable shift in Department B."""
    shift_data = ShiftCreate(
        name="ER 1",
        position_id=position_b.id,
        grants_day_off=False,
        doctors_per_shift=2
    )

    return shift_repository.create_shift(session, shift_data)


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


@pytest.fixture(name="other_doctor")
def other_doctor_fixture(session, department, team):
    """Creates a second doctor in Department A for ownership-scope tests."""
    return create_new_doctor(session, "Dr Other", "drother@gmail.com", department.id, team.id)


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
    assert new_doctor.department_id == department.id
    assert new_doctor.team_id == team.id
    assert new_doctor.is_deleted is False
    assert new_doctor.sync_status is False
    assert isinstance(new_doctor.created_at, datetime.datetime)
    assert isinstance(new_doctor.updated_at, datetime.datetime)


def test_stage_doctor_flushes_without_committing(session, department):
    """Verify that stage_doctor flushes in-session but rolls back cleanly."""
    staged = doctor_repository.stage_doctor(
        session=session,
        name="Dr. Staged",
        department_id=department.id,
    )

    assert staged.id is not None
    assert staged.name == "Dr. Staged"
    assert staged.department_id == department.id

    fetched = session.get(DoctorModel, staged.id)
    assert fetched is not None
    assert fetched.name == "Dr. Staged"

    session.rollback()
    assert session.get(DoctorModel, staged.id) is None


def test_get_doctor_by_email(session, new_doctor):
    """Test retrieving a doctor by linked user email"""
    from src.user.models import User as UserModel
    from src.auth.security import hash_password

    # Before linking user
    assert doctor_repository.get_doctor_by_email(session, "drpanos@gmail.com") is None

    # Link user to doctor
    user = UserModel(
        email="drpanos@gmail.com",
        hashed_password=hash_password("Pass123!"),
        full_name="Dr Panos",
        role=UserRole.DOCTOR,
        department_id=new_doctor.department_id,
        doctor_id=new_doctor.id,
    )
    session.add(user)
    session.commit()

    retrieved_doctor = doctor_repository.get_doctor_by_email(
        session, "drpanos@gmail.com")

    assert retrieved_doctor is not None
    assert retrieved_doctor.id == new_doctor.id


def test_get_doctor_by_id_for_department_returns_own_active_doctor(
    session,
    new_doctor,
):
    """Tests retrieving an active doctor within department scope."""
    retrieved_doctor = doctor_repository.get_doctor_by_id_for_department(
        session=session,
        doctor_id=new_doctor.id,
        department_id=new_doctor.department_id,
    )

    assert retrieved_doctor is not None
    assert retrieved_doctor.id == new_doctor.id


def test_get_doctor_by_id_for_department_hides_foreign_doctor(
    session,
    department,
    doctor_b,
):
    """Tests that scoped lookup hides another department's doctor."""
    retrieved_doctor = doctor_repository.get_doctor_by_id_for_department(
        session=session,
        doctor_id=doctor_b.id,
        department_id=department.id,
    )

    assert retrieved_doctor is None


def test_get_doctor_by_id_for_department_hides_deleted_doctor(
    session,
    new_doctor,
):
    """Tests that scoped lookup hides deleted doctors."""
    new_doctor.is_deleted = True
    session.add(new_doctor)
    session.commit()

    retrieved_doctor = doctor_repository.get_doctor_by_id_for_department(
        session=session,
        doctor_id=new_doctor.id,
        department_id=new_doctor.department_id,
    )

    assert retrieved_doctor is None


def test_get_active_doctors_for_department(
    session,
    team,
    department,
    new_doctor,
    doctor_b,
):
    """Tests that active-doctor listing is scoped to a single department."""
    deleted_doctor = create_new_doctor(
        session, "Dr Panagiotis", "drpanagiotis@gmail.com", department.id, team.id)

    deleted_doctor.is_deleted = True
    session.add(deleted_doctor)
    session.commit()

    active_doctors = doctor_repository.get_active_doctors_for_department(
        session=session,
        department_id=department.id,
    )

    assert {doctor.id for doctor in active_doctors} == {new_doctor.id}
    assert all(
        doctor.department_id == department.id for doctor in active_doctors
    )


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


def test_create_doctor_unavailability(new_doctor, unavailability):
    """Test creating unavailability dates for a doctor"""
    assert isinstance(unavailability.id, uuid.UUID)
    assert unavailability.date == datetime.date(2026, 8, 12)
    assert unavailability.doctor_id == new_doctor.id


def test_get_doctor_unavailability_returns_only_active_records(
    session,
    new_doctor,
    unavailability,
):
    """Tests that only active doctor unavailability records are returned."""
    unavailability_b = create_test_unavailability(
        session=session,
        doctor=new_doctor,
        date=datetime.date(2026, 8, 13),
    )

    unavailability.is_deleted = True
    session.add(unavailability)
    session.commit()

    retrieved_unavailabilities = doctor_repository.get_doctor_unavailability(
        session=session,
        doctor_id=new_doctor.id,
    )

    assert isinstance(retrieved_unavailabilities, list)
    assert unavailability_b in retrieved_unavailabilities
    assert unavailability not in retrieved_unavailabilities


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

    doctor_pos = doctor_repository.get_doctor_positions_for_department(
        session=session,
        doctor_id=new_doctor.id,
        department_id=new_doctor.department_id
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


def test_doctor_has_department_foreign_key(new_doctor, department, team):
    """Test that doctors can be associated with a department and retrieved correctly."""

    assert isinstance(new_doctor.id, uuid.UUID)
    assert new_doctor.department_id == department.id
    assert new_doctor.team_id == team.id


#######################
# Controller tests
#######################


def test_create_doctor_controller_optional_team(session, department):
    """Test that creating a doctor with team_id=None succeeds."""
    doctor_data = DoctorCreate(
        name="Dr Teamless",
        team_id=None,
    )
    doctor = doctor_controllers.create_doctor_controller(
        session=session,
        department_id=department.id,
        doctor_data=doctor_data,
    )
    assert doctor.name == "Dr Teamless"
    assert doctor.team_id is None
    assert doctor.department_id == department.id


def test_get_doctor_controller_nonexistent(session, department):
    """Test that retrieving a non-existent doctor raises a 404 error"""
    non_existent_id = uuid.uuid4()
    with pytest.raises(HTTPException) as exc_info:
        doctor_controllers.get_doctor_controller(
            session=session,
            department_id=department.id,
            doctor_id=non_existent_id,
        )

    assert exc_info.value.status_code == 404
    assert "not found" in exc_info.value.detail


def test_get_doctor_controller_deleted(session, new_doctor):
    """Test that retrieving a deleted team raises a 404 error"""
    new_doctor.is_deleted = True
    session.add(new_doctor)
    session.commit()

    with pytest.raises(HTTPException) as exc_info:
        doctor_controllers.get_doctor_controller(
            session=session,
            doctor_id=new_doctor.id,
            department_id=new_doctor.department_id,
        )

    assert exc_info.value.status_code == 404
    assert "not found" in exc_info.value.detail


def test_create_doctor_pre_assignment_controller_duplicate_date(
    session,
    new_doctor,
    pre_assignment,
):
    """Tests that creating a duplicate pre assignment returns error"""
    # pre_assignment is needed to crate the first (duplicate) pre assignment
    pre_assignment_data = DoctorPreAssignmentCreate(
        date=pre_assignment.date,
        shift_id=pre_assignment.shift_id
    )

    with pytest.raises(HTTPException) as exc_info:
        doctor_controllers.create_doctor_pre_assignment_controller(
            session=session,
            doctor_id=new_doctor.id,
            department_id=new_doctor.department_id,
            pre_assignment_data=pre_assignment_data,
        )

    assert exc_info.value.status_code == 400
    assert "already exists" in exc_info.value.detail


def test_create_doctor_pre_assignment_controller_nonexistent_doctor(session, department, shift):
    """Tests that creating a pre assignment with a nonexistent doctor returns a scoped 404."""
    missing_doctor_id = uuid.uuid4()
    pre_assignment_data = DoctorPreAssignmentCreate(
        date=datetime.date(2026, 8, 12),
        shift_id=shift.id
    )

    with pytest.raises(HTTPException) as exc_info:
        doctor_controllers.create_doctor_pre_assignment_controller(
            session=session,
            doctor_id=missing_doctor_id,
            department_id=department.id,
            pre_assignment_data=pre_assignment_data
        )

    assert exc_info.value.status_code == 404
    assert exc_info.value.detail == "Doctor or shift not found."
    assert doctor_repository.get_doctor_pre_assignment_by_date(
        session=session,
        doctor_id=missing_doctor_id,
        target_date=datetime.date(2026, 8, 12),
    ) is None


def test_create_doctor_pre_assignment_controller_hides_foreign_doctor(
    session,
    department,
    doctor_b,
    shift,
):
    """Tests that a department admin cannot create a pre-assignment for a foreign doctor."""
    pre_assignment_data = DoctorPreAssignmentCreate(
        date=datetime.date(2026, 8, 12),
        shift_id=shift.id,
    )

    with pytest.raises(HTTPException) as exc_info:
        doctor_controllers.create_doctor_pre_assignment_controller(
            session=session,
            doctor_id=doctor_b.id,
            department_id=department.id,
            pre_assignment_data=pre_assignment_data,
        )

    assert exc_info.value.status_code == 404
    assert exc_info.value.detail == "Doctor or shift not found."
    assert doctor_repository.get_doctor_pre_assignment_by_date(
        session=session,
        doctor_id=doctor_b.id,
        target_date=datetime.date(2026, 8, 12),
    ) is None


def test_create_doctor_pre_assignment_controller_hides_foreign_shift(
    session,
    new_doctor,
    shift_b,
):
    """Tests that a department admin cannot create a pre-assignment using a foreign shift."""
    pre_assignment_data = DoctorPreAssignmentCreate(
        date=datetime.date(2026, 8, 12),
        shift_id=shift_b.id,
    )

    with pytest.raises(HTTPException) as exc_info:
        doctor_controllers.create_doctor_pre_assignment_controller(
            session=session,
            doctor_id=new_doctor.id,
            department_id=new_doctor.department_id,
            pre_assignment_data=pre_assignment_data,
        )

    assert exc_info.value.status_code == 404
    assert exc_info.value.detail == "Doctor or shift not found."
    assert doctor_repository.get_doctor_pre_assignment_by_date(
        session=session,
        doctor_id=new_doctor.id,
        target_date=datetime.date(2026, 8, 12),
    ) is None


def test_create_doctor_pre_assignment_controller_handles_missing_shift(
    session,
    new_doctor,
):
    """Tests that a missing Shift returns 404 instead of causing a 500."""
    missing_shift_id = uuid.uuid4()
    pre_assignment_data = DoctorPreAssignmentCreate(
        date=datetime.date(2026, 8, 12),
        shift_id=missing_shift_id,
    )

    with pytest.raises(HTTPException) as exc_info:
        doctor_controllers.create_doctor_pre_assignment_controller(
            session=session,
            doctor_id=new_doctor.id,
            department_id=new_doctor.department_id,
            pre_assignment_data=pre_assignment_data,
        )

    assert exc_info.value.status_code == 404
    assert exc_info.value.detail == "Doctor or shift not found."
    assert doctor_repository.get_doctor_pre_assignment_by_date(
        session=session,
        doctor_id=new_doctor.id,
        target_date=datetime.date(2026, 8, 12),
    ) is None


def test_create_doctor_pre_assignment_checks_scope_before_duplicate(
    session,
    department,
    doctor_b,
    shift_b,
):
    """Tests that foreign Doctor/Shift resources are hidden before duplicate-date checks."""
    existing_pre_assignment = create_test_pre_assignment(
        session, shift_b.id, doctor_b, datetime.date(2026, 8, 12)
    )

    pre_assignment_data = DoctorPreAssignmentCreate(
        date=datetime.date(2026, 8, 12),
        shift_id=shift_b.id,
    )

    with pytest.raises(HTTPException) as exc_info:
        doctor_controllers.create_doctor_pre_assignment_controller(
            session=session,
            doctor_id=doctor_b.id,
            department_id=department.id,
            pre_assignment_data=pre_assignment_data,
        )

    assert exc_info.value.status_code == 404
    assert exc_info.value.detail == "Doctor or shift not found."

    stored_pre_assignments = doctor_repository.get_doctor_pre_assignments(
        session=session,
        doctor_id=doctor_b.id,
    )
    assert len(stored_pre_assignments) == 1
    assert stored_pre_assignments[0].id == existing_pre_assignment.id


def test_list_doctor_pre_assignments_controller_hides_foreign_doctor(
    session,
    department,
    doctor_b,
):
    """Tests that listing pre-assignments for a foreign doctor returns 404, not an empty list."""
    with pytest.raises(HTTPException) as exc_info:
        doctor_controllers.list_doctor_pre_assignments_controller(
            session=session,
            doctor_id=doctor_b.id,
            department_id=department.id,
        )

    assert exc_info.value.status_code == 404
    assert exc_info.value.detail == "Doctor not found."


def test_create_doctor_unavailability_controller_duplicate_date(
    session,
    new_doctor,
    unavailability,
    department_admin_user,
):
    """Test that creating a duplicate unavailability returns error after authorization passes"""
    unavailability_data = DoctorUnavailabilityCreate(
        date=unavailability.date,
    )

    with pytest.raises(HTTPException) as exc_info:
        doctor_controllers.create_doctor_unavailability_controller(
            session=session,
            doctor_id=new_doctor.id,
            department_id=new_doctor.department_id,
            current_user=department_admin_user,
            unavailability_data=unavailability_data
        )

    assert exc_info.value.status_code == 400
    assert "already exists" in exc_info.value.detail


def test_create_doctor_unavailability_controller_nonexistent_doctor(
    session,
    department,
    department_admin_user,
):
    """Tests that creating an unavailability for a nonexistent doctor returns a scoped 404."""
    missing_doctor_id = uuid.uuid4()
    unavailability_data = DoctorUnavailabilityCreate(
        date=datetime.date(2026, 8, 12),
    )

    with pytest.raises(HTTPException) as exc_info:
        doctor_controllers.create_doctor_unavailability_controller(
            session=session,
            doctor_id=missing_doctor_id,
            department_id=department.id,
            current_user=department_admin_user,
            unavailability_data=unavailability_data,
        )

    assert exc_info.value.status_code == 404
    assert exc_info.value.detail == "Doctor not found."
    assert doctor_repository.get_doctor_unavailability_by_date(
        session=session,
        doctor_id=missing_doctor_id,
        target_date=datetime.date(2026, 8, 12),
    ) is None


def test_create_doctor_unavailability_controller_hides_foreign_doctor(
    session,
    department,
    department_admin_user,
    doctor_b,
):
    """Tests that a department admin cannot create unavailability for a foreign doctor."""
    unavailability_data = DoctorUnavailabilityCreate(
        date=datetime.date(2026, 9, 1))

    with pytest.raises(HTTPException) as exc_info:
        doctor_controllers.create_doctor_unavailability_controller(
            session=session,
            doctor_id=doctor_b.id,
            department_id=department.id,
            current_user=department_admin_user,
            unavailability_data=unavailability_data,
        )

    assert exc_info.value.status_code == 404
    assert exc_info.value.detail == "Doctor not found."
    assert doctor_repository.get_doctor_unavailability_by_date(
        session=session,
        doctor_id=doctor_b.id,
        target_date=datetime.date(2026, 9, 1),
    ) is None


def test_create_doctor_unavailability_controller_rejects_doctor_targeting_other_doctor(
    session,
    department,
    doctor_user,
    other_doctor,
):
    """Tests that a doctor cannot create unavailability for another same-department doctor."""
    unavailability_data = DoctorUnavailabilityCreate(
        date=datetime.date(2026, 9, 2))

    with pytest.raises(HTTPException) as exc_info:
        doctor_controllers.create_doctor_unavailability_controller(
            session=session,
            doctor_id=other_doctor.id,
            department_id=department.id,
            current_user=doctor_user,
            unavailability_data=unavailability_data,
        )

    assert exc_info.value.status_code == 403
    assert exc_info.value.detail == "Cannot access another doctor's unavailability."
    assert doctor_repository.get_doctor_unavailability_by_date(
        session=session,
        doctor_id=other_doctor.id,
        target_date=datetime.date(2026, 9, 2),
    ) is None


def test_create_doctor_unavailability_controller_rejects_doctor_targeting_foreign_doctor(
    session,
    department,
    doctor_user,
    doctor_b,
):
    """Tests that a doctor cannot create unavailability for a foreign-department doctor."""
    unavailability_data = DoctorUnavailabilityCreate(
        date=datetime.date(2026, 9, 3))

    with pytest.raises(HTTPException) as exc_info:
        doctor_controllers.create_doctor_unavailability_controller(
            session=session,
            doctor_id=doctor_b.id,
            department_id=department.id,
            current_user=doctor_user,
            unavailability_data=unavailability_data,
        )

    assert exc_info.value.status_code == 403
    assert exc_info.value.detail == "Cannot access another doctor's unavailability."
    assert doctor_repository.get_doctor_unavailability_by_date(
        session=session,
        doctor_id=doctor_b.id,
        target_date=datetime.date(2026, 9, 3),
    ) is None


def test_create_doctor_unavailability_checks_ownership_before_duplicate(
    session,
    department,
    doctor_user,
    other_doctor,
):
    """Tests that doctor ownership is enforced before the duplicate-date check runs."""
    existing_unavailability = create_test_unavailability(
        session, other_doctor, datetime.date(2026, 9, 5)
    )

    unavailability_data = DoctorUnavailabilityCreate(
        date=datetime.date(2026, 9, 5))

    with pytest.raises(HTTPException) as exc_info:
        doctor_controllers.create_doctor_unavailability_controller(
            session=session,
            doctor_id=other_doctor.id,
            department_id=department.id,
            current_user=doctor_user,
            unavailability_data=unavailability_data,
        )

    assert exc_info.value.status_code == 403
    assert exc_info.value.detail == "Cannot access another doctor's unavailability."

    stored_unavailabilities = doctor_repository.get_doctor_unavailability(
        session=session,
        doctor_id=other_doctor.id,
    )
    assert len(stored_unavailabilities) == 1
    assert stored_unavailabilities[0].id == existing_unavailability.id


def test_create_doctor_unavailability_checks_admin_scope_before_duplicate(
    session,
    department,
    department_admin_user,
    doctor_b,
):
    """Tests that admin department scope is enforced before the duplicate-date check runs."""
    existing_unavailability = create_test_unavailability(
        session, doctor_b, datetime.date(2026, 9, 6)
    )

    unavailability_data = DoctorUnavailabilityCreate(
        date=datetime.date(2026, 9, 6))

    with pytest.raises(HTTPException) as exc_info:
        doctor_controllers.create_doctor_unavailability_controller(
            session=session,
            doctor_id=doctor_b.id,
            department_id=department.id,
            current_user=department_admin_user,
            unavailability_data=unavailability_data,
        )

    assert exc_info.value.status_code == 404
    assert exc_info.value.detail == "Doctor not found."

    stored_unavailabilities = doctor_repository.get_doctor_unavailability(
        session=session,
        doctor_id=doctor_b.id,
    )
    assert len(stored_unavailabilities) == 1
    assert stored_unavailabilities[0].id == existing_unavailability.id


def test_list_doctor_unavailability_controller_admin_own_succeeds(
    session,
    department,
    department_admin_user,
    new_doctor,
    unavailability,
):
    """Tests that a department admin can list unavailability for their own department's doctor."""
    result = doctor_controllers.list_doctor_unavailability_controller(
        session=session,
        doctor_id=new_doctor.id,
        department_id=department.id,
        current_user=department_admin_user,
    )

    assert unavailability in result


def test_list_doctor_unavailability_controller_hides_foreign_doctor(
    session,
    department,
    department_admin_user,
    doctor_b,
):
    """Tests that a department admin cannot list unavailability for a foreign doctor."""
    with pytest.raises(HTTPException) as exc_info:
        doctor_controllers.list_doctor_unavailability_controller(
            session=session,
            doctor_id=doctor_b.id,
            department_id=department.id,
            current_user=department_admin_user,
        )

    assert exc_info.value.status_code == 404
    assert exc_info.value.detail == "Doctor not found."


def test_list_doctor_unavailability_controller_doctor_own_succeeds(
    session,
    department,
    doctor_user,
    new_doctor,
    unavailability,
):
    """Tests that a doctor can list their own unavailability."""
    result = doctor_controllers.list_doctor_unavailability_controller(
        session=session,
        doctor_id=new_doctor.id,
        department_id=department.id,
        current_user=doctor_user,
    )

    assert unavailability in result


def test_list_doctor_unavailability_controller_rejects_doctor_targeting_other_doctor(
    session,
    department,
    doctor_user,
    other_doctor,
):
    """Tests that a doctor cannot list another same-department doctor's unavailability."""
    with pytest.raises(HTTPException) as exc_info:
        doctor_controllers.list_doctor_unavailability_controller(
            session=session,
            doctor_id=other_doctor.id,
            department_id=department.id,
            current_user=doctor_user,
        )

    assert exc_info.value.status_code == 403
    assert exc_info.value.detail == "Cannot access another doctor's unavailability."


def test_list_doctor_unavailability_controller_rejects_doctor_targeting_foreign_doctor(
    session,
    department,
    doctor_user,
    doctor_b,
):
    """Tests that a doctor cannot list a foreign-department doctor's unavailability."""
    with pytest.raises(HTTPException) as exc_info:
        doctor_controllers.list_doctor_unavailability_controller(
            session=session,
            doctor_id=doctor_b.id,
            department_id=department.id,
            current_user=doctor_user,
        )

    assert exc_info.value.status_code == 403
    assert exc_info.value.detail == "Cannot access another doctor's unavailability."


def test_create_doctor_position_controller_duplicate_assignment(session, new_doctor, position):
    """Tests that creating a duplicate doctor-position returns error"""
    create_test_doctor_position(session, new_doctor, position)

    doctor_position_data = DoctorPositionCreate(
        position_id=position.id
    )

    with pytest.raises(HTTPException) as exc_info:
        doctor_controllers.create_doctor_position_controller(
            session=session,
            doctor_id=new_doctor.id,
            doctor_pos_data=doctor_position_data,
            department_id=position.department_id,
        )

    assert exc_info.value.status_code == 400
    assert "already assigned" in exc_info.value.detail


def test_create_doctor_position_controller_nonexistent_doctor(session, position):
    """Tests that a missing doctor returns the normal scoped 404 response."""
    missing_doctor_id = uuid.uuid4()
    doctor_position_data = DoctorPositionCreate(
        position_id=position.id
    )

    with pytest.raises(HTTPException) as exc_info:
        doctor_controllers.create_doctor_position_controller(
            session=session,
            doctor_id=missing_doctor_id,
            doctor_pos_data=doctor_position_data,
            department_id=position.department_id,
        )

    assert exc_info.value.status_code == 404
    assert "not found." in exc_info.value.detail
    assert doctor_repository.get_doctor_position_by_id(
        session=session,
        doctor_id=missing_doctor_id,
        position_id=position.id,
    ) is None


def test_create_doctor_position_controller_hides_foreign_doctor(
    session,
    doctor_b,
    position,
):
    """Tests that an admin cannot assign a Position to a foreign doctor."""
    doctor_pos_data = DoctorPositionCreate(
        position_id=position.id
    )

    with pytest.raises(HTTPException) as exc_info:
        doctor_controllers.create_doctor_position_controller(
            session=session,
            doctor_id=doctor_b.id,
            doctor_pos_data=doctor_pos_data,
            department_id=position.department_id,
        )

    assert exc_info.value.status_code == 404
    assert "not found." in exc_info.value.detail
    assert doctor_repository.get_doctor_position_by_id(
        session=session,
        doctor_id=doctor_b.id,
        position_id=position.id,
    ) is None


def test_create_doctor_position_controller_hides_foreign_position(
    session,
    new_doctor,
    position_b,
):
    """Tests that an admin cannot assign a foreign Position to a doctor."""
    doctor_pos_data = DoctorPositionCreate(
        position_id=position_b.id
    )

    with pytest.raises(HTTPException) as exc_info:
        doctor_controllers.create_doctor_position_controller(
            session=session,
            doctor_id=new_doctor.id,
            doctor_pos_data=doctor_pos_data,
            department_id=new_doctor.department_id,
        )

    assert exc_info.value.status_code == 404
    assert "not found." in exc_info.value.detail
    assert doctor_repository.get_doctor_position_by_id(
        session=session,
        doctor_id=new_doctor.id,
        position_id=position_b.id,
    ) is None


def test_create_doctor_position_controller_handles_missing_position(
    session,
    new_doctor,
):
    """Tests that a missing Position returns 404 instead of causing a 500."""
    missing_position_id = uuid.uuid4()
    doctor_pos_data = DoctorPositionCreate(
        position_id=missing_position_id
    )

    with pytest.raises(HTTPException) as exc_info:
        doctor_controllers.create_doctor_position_controller(
            session=session,
            doctor_id=new_doctor.id,
            doctor_pos_data=doctor_pos_data,
            department_id=new_doctor.department_id,
        )

    assert exc_info.value.status_code == 404
    assert "not found." in exc_info.value.detail
    assert doctor_repository.get_doctor_position_by_id(
        session=session,
        doctor_id=new_doctor.id,
        position_id=missing_position_id,
    ) is None


def test_create_doctor_position_checks_scope_before_duplicate(
    session,
    department,
    doctor_b,
    position_b,
):
    """Tests that foreign resources are hidden before duplicate checks."""
    existing_assignment = create_test_doctor_position(
        session=session,
        doctor=doctor_b,
        position=position_b,
    )

    doctor_pos_data = DoctorPositionCreate(
        position_id=position_b.id
    )

    with pytest.raises(HTTPException) as exc_info:
        doctor_controllers.create_doctor_position_controller(
            session=session,
            doctor_id=doctor_b.id,
            doctor_pos_data=doctor_pos_data,
            department_id=department.id,
        )

    assert exc_info.value.status_code == 404
    assert "not found" in exc_info.value.detail

    stored_assignment = doctor_repository.get_doctor_position_by_id(
        session=session,
        doctor_id=doctor_b.id,
        position_id=position_b.id,
    )
    assert stored_assignment is not None
    assert stored_assignment.id == existing_assignment.id


def test_list_doctor_positions_controller_returns_own_department(
    session,
    department,
    new_doctor,
    position,
):
    """Tests that a department member can list associations for a doctor in their own department."""
    doctor_position = create_test_doctor_position(
        session=session,
        doctor=new_doctor,
        position=position,
    )

    result = doctor_controllers.list_doctor_positions_controller(
        session=session,
        doctor_id=new_doctor.id,
        department_id=department.id,
    )

    assert isinstance(result, list)
    assert any(item.id == doctor_position.id for item in result)


def test_list_doctor_positions_controller_hides_foreign_doctor(
    session,
    department,
    doctor_b,
    position_b,
):
    """Tests that the list controller returns 404 instead of an empty list for a foreign doctor."""
    create_test_doctor_position(
        session=session,
        doctor=doctor_b,
        position=position_b,
    )

    with pytest.raises(HTTPException) as exc_info:
        doctor_controllers.list_doctor_positions_controller(
            session=session,
            doctor_id=doctor_b.id,
            department_id=department.id,
        )

    assert exc_info.value.status_code == 404
    assert exc_info.value.detail == "Doctor not found."
    assert doctor_repository.get_doctor_position_by_id(
        session=session,
        doctor_id=doctor_b.id,
        position_id=position_b.id,
    ) is not None


def test_list_doctor_positions_controller_hides_deleted_doctor(
    session,
    new_doctor,
    position,
):
    """Tests that the list controller returns 404 for a soft-deleted doctor."""
    create_test_doctor_position(
        session=session,
        doctor=new_doctor,
        position=position,
    )

    new_doctor.is_deleted = True
    session.add(new_doctor)
    session.commit()

    with pytest.raises(HTTPException) as exc_info:
        doctor_controllers.list_doctor_positions_controller(
            session=session,
            doctor_id=new_doctor.id,
            department_id=new_doctor.department_id,
        )

    assert exc_info.value.status_code == 404
    assert exc_info.value.detail == "Doctor not found."


def test_create_pre_assignment_unavailability_conflict(session, new_doctor, unavailability, shift):
    """Tests that a pre assignment cannot be assigned on an unavailable day"""
    pre_assignment_data = DoctorPreAssignmentCreate(
        date=unavailability.date,
        shift_id=shift.id,
    )

    with pytest.raises(HTTPException) as exc_info:
        doctor_controllers.create_doctor_pre_assignment_controller(
            session=session,
            doctor_id=new_doctor.id,
            department_id=new_doctor.department_id,
            pre_assignment_data=pre_assignment_data,
        )

    assert exc_info.value.status_code == 422
    assert "cannot be assigned to an unavailable day" in exc_info.value.detail


def test_create_unavailability_pre_assignment_conflict(session, new_doctor, pre_assignment, department_admin_user):
    """Tests that a doctor cannot be marked unavailable on a date with an existing pre-assignment."""
    unavailability_data = DoctorUnavailabilityCreate(
        date=pre_assignment.date,
    )

    with pytest.raises(HTTPException) as exc_info:
        doctor_controllers.create_doctor_unavailability_controller(
            session=session,
            doctor_id=new_doctor.id,
            department_id=new_doctor.department_id,
            current_user=department_admin_user,
            unavailability_data=unavailability_data,
        )

    assert exc_info.value.status_code == 422
    assert "cannot be marked unavailable on a pre-assigned day" in exc_info.value.detail

#######################
# Route tests
#######################


def test_create_doctor_route(client, session, department, team, department_admin_headers):
    """Test that POST /doctors/ derives department_id from the authenticated admin"""

    response = client.post(
        "api/v1/doctors",
        json={
            "name": "Dr Panos",
            "team_id": str(team.id),
        },
        headers=department_admin_headers,
    )

    assert response.status_code == 201
    data = response.json()
    assert data["name"] == "Dr Panos"
    assert data["department_id"] == str(department.id)
    assert data["team_id"] == str(team.id)
    assert "id" in data
    assert "created_at" in data
    assert "updated_at" in data

    persisted_doctor = doctor_repository.get_doctor_by_id_for_department(
        session=session,
        doctor_id=uuid.UUID(data["id"]),
        department_id=department.id,
    )
    assert persisted_doctor is not None
    assert persisted_doctor.department_id == department.id
    assert persisted_doctor.team_id == team.id


def test_create_doctor_route_missing_required_field(client, team, department_admin_headers):
    """Test the POST /doctors/ route rejects a payload missing a required field"""

    response = client.post(
        "api/v1/doctors",
        json={
            "team_id": str(team.id),
        },
        headers=department_admin_headers,
    )
    assert response.status_code == 422


def test_create_doctor_route_rejects_supplied_department_id(
    client,
    session,
    department,
    team,
    department_admin_headers,
):
    """Tests that POST /doctors/ rejects a client-supplied department_id"""

    response = client.post(
        "api/v1/doctors",
        json={
            "name": "Dr Panos",
            "team_id": str(team.id),
            "department_id": str(department.id),
        },
        headers=department_admin_headers,
    )

    assert response.status_code == 422
    validation_errors = response.json()["detail"]
    department_id_error = next(
        (
            error for error in validation_errors
            if error["loc"] == ["body", "department_id"]
        ),
        None,
    )

    assert department_id_error is not None
    assert department_id_error["type"] == "extra_forbidden"


def test_create_doctor_route_hides_foreign_team(
    client,
    session,
    team_b,
    department_admin_headers,
):
    """Tests that a department admin cannot create a doctor using a foreign team"""

    response = client.post(
        "api/v1/doctors",
        json={
            "name": "Dr Panos",
            "team_id": str(team_b.id),
        },
        headers=department_admin_headers,
    )

    assert response.status_code == 404
    assert response.json() == {"detail": "Team not found."}
    assert response.headers.get("WWW-Authenticate") is None


def test_get_doctor_by_id_route(client, department, team, new_doctor, department_admin_headers):
    """Tests that GET /doctors/{doctor_id} returns full doctor detail for a department admin"""
    response = client.get(
        f"/api/v1/doctors/{new_doctor.id}",
        headers=department_admin_headers,
    )

    assert response.status_code == 200
    data = response.json()
    assert data["name"] == "Dr Panos"
    assert data["department_id"] == str(department.id)
    assert data["team_id"] == str(team.id)
    assert "id" in data
    assert "created_at" in data
    assert "updated_at" in data


def test_get_doctor_by_id_route_nonexistent(client, department_admin_headers):
    """Test the GET /doctors/{doctor_id} route returns error when given a nonexistent id"""
    response = client.get(
        f"/api/v1/doctors/{uuid.uuid4()}",
        headers=department_admin_headers,
    )
    assert response.status_code == 404


def test_get_doctor_by_id_route_hides_foreign_doctor(
    client,
    doctor_b,
    department_admin_headers,
):
    """Tests that a department admin cannot view another department's doctor"""
    response = client.get(
        f"/api/v1/doctors/{doctor_b.id}",
        headers=department_admin_headers,
    )

    assert response.status_code == 404
    assert response.json() == {"detail": "Doctor not found."}
    assert response.headers.get("WWW-Authenticate") is None


# new doctor is needed to add a doctor in the db
def test_list_doctors_route(client, new_doctor, doctor_b, department_admin_headers):
    """Tests that GET /doctors/ returns the full, department-scoped doctor list"""
    response = client.get(
        "api/v1/doctors",
        headers=department_admin_headers,
    )

    assert response.status_code == 200
    data = response.json()

    returned_ids = {item["id"] for item in data}
    assert str(new_doctor.id) in returned_ids
    assert str(doctor_b.id) not in returned_ids

    returned_doctor = next(
        item for item in data if item["id"] == str(new_doctor.id)
    )
    assert returned_doctor["name"] == new_doctor.name


@pytest.mark.parametrize(
    "path_template",
    [
        pytest.param("/api/v1/doctors/", id="list-doctors"),
        pytest.param("/api/v1/doctors/{doctor_id}", id="get-doctor"),
    ],
)
@pytest.mark.parametrize(
    "role",
    [
        UserRole.VIEWER,
        UserRole.DOCTOR,
        UserRole.SUPER_ADMIN,
    ]
)
def test_full_doctor_routes_reject_non_department_admin(
    client,
    role,
    user_factory,
    auth_headers_factory,
    department,
    new_doctor,
    path_template,
):
    """Tests that only department admins may access full doctor list/detail"""
    path = path_template.format(doctor_id=new_doctor.id)
    user = user_factory(
        role=role,
        department_id=None if role == UserRole.SUPER_ADMIN else department.id,
        doctor_id=new_doctor.id if role == UserRole.DOCTOR else None,
    )

    response = client.get(path, headers=auth_headers_factory(user))

    assert response.status_code == 403
    assert response.json() == {
        "detail": "Insufficient permissions for this operation."}
    assert response.headers.get("WWW-Authenticate") is None


def test_list_doctors_roster_route(client, new_doctor, doctor_b, viewer_headers):
    """Tests that GET /doctors/roster returns a reduced, department-scoped roster"""
    response = client.get(
        "/api/v1/doctors/roster",
        headers=viewer_headers,
    )

    assert response.status_code == 200
    data = response.json()

    returned_ids = {item["id"] for item in data}
    assert str(new_doctor.id) in returned_ids
    assert str(doctor_b.id) not in returned_ids

    roster_entry = next(
        item for item in data if item["id"] == str(new_doctor.id)
    )
    assert set(roster_entry.keys()) == {
        "id", "name", "department_id", "team_id"
    }
    assert roster_entry["department_id"] == str(new_doctor.department_id)
    assert roster_entry["team_id"] == str(new_doctor.team_id)
    assert "email" not in roster_entry
    assert "sync_status" not in roster_entry
    assert "created_at" not in roster_entry
    assert "updated_at" not in roster_entry


@pytest.mark.parametrize(
    "role",
    [
        UserRole.DEPARTMENT_ADMIN,
        UserRole.DOCTOR,
        UserRole.VIEWER,
    ]
)
def test_doctor_roster_route_allows_department_members(
    client,
    role,
    user_factory,
    auth_headers_factory,
    department,
    new_doctor,
):
    """Tests that department admins, doctors, and viewers can access the roster"""
    user = user_factory(
        role=role,
        department_id=department.id,
        doctor_id=new_doctor.id if role == UserRole.DOCTOR else None,
    )

    response = client.get(
        "/api/v1/doctors/roster",
        headers=auth_headers_factory(user),
    )

    assert response.status_code == 200


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


def test_create_doctor_pre_assignments_route_invalid_payload(
    client,
    new_doctor,
    department_admin_headers,
):
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


def test_create_doctor_pre_assignments_route_hides_foreign_doctor(
    client,
    session,
    doctor_b,
    shift,
    department_admin_headers,
):
    """Tests that an admin cannot create a pre-assignment for a foreign doctor."""
    response = client.post(
        f"/api/v1/doctors/{doctor_b.id}/pre-assignments",
        json={
            "date": str(datetime.date(2026, 8, 12)),
            "shift_id": str(shift.id),
        },
        headers=department_admin_headers,
    )

    assert response.status_code == 404
    assert response.json() == {"detail": "Doctor or shift not found."}
    assert response.headers.get("WWW-Authenticate") is None
    assert doctor_repository.get_doctor_pre_assignment_by_date(
        session=session,
        doctor_id=doctor_b.id,
        target_date=datetime.date(2026, 8, 12),
    ) is None


def test_create_doctor_pre_assignments_route_hides_foreign_shift(
    client,
    session,
    new_doctor,
    shift_b,
    department_admin_headers,
):
    """Tests that an admin cannot create a pre-assignment using a foreign shift."""
    response = client.post(
        f"/api/v1/doctors/{new_doctor.id}/pre-assignments",
        json={
            "date": str(datetime.date(2026, 8, 12)),
            "shift_id": str(shift_b.id),
        },
        headers=department_admin_headers,
    )

    assert response.status_code == 404
    assert response.json() == {"detail": "Doctor or shift not found."}
    assert response.headers.get("WWW-Authenticate") is None
    assert doctor_repository.get_doctor_pre_assignment_by_date(
        session=session,
        doctor_id=new_doctor.id,
        target_date=datetime.date(2026, 8, 12),
    ) is None


def test_get_doctor_pre_assignments_route_hides_foreign_doctor(
    client,
    doctor_b,
    department_admin_headers,
):
    """Tests that listing pre-assignments for a foreign doctor returns 404, not an empty list."""
    response = client.get(
        f"/api/v1/doctors/{doctor_b.id}/pre-assignments",
        headers=department_admin_headers,
    )

    assert response.status_code == 404
    assert response.json() == {"detail": "Doctor not found."}
    assert response.headers.get("WWW-Authenticate") is None


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


def test_create_doctor_unavailability_pre_assignment_conflict_route(
    client, new_doctor, pre_assignment, department_admin_headers
):
    """Tests that POST /doctors/{doctor_id}/unavailability rejects when pre-assignment exists on that date"""
    response = client.post(
        f"api/v1/doctors/{new_doctor.id}/unavailability",
        json={
            "date": str(pre_assignment.date),
        },
        headers=department_admin_headers,
    )
    assert response.status_code == 422
    assert "cannot be marked unavailable on a pre-assigned day" in response.json()["detail"]


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


def test_get_doctor_unavailability_route_department_admin_own(
    client,
    new_doctor,
    unavailability,
    department_admin_headers,
):
    """Tests that a department admin can list unavailability for their own department's doctor."""
    response = client.get(
        f"/api/v1/doctors/{new_doctor.id}/unavailability",
        headers=department_admin_headers,
    )

    assert response.status_code == 200
    data = response.json()
    assert isinstance(data, list)
    assert any(item["id"] == str(unavailability.id) for item in data)


def test_create_doctor_unavailability_route_hides_foreign_doctor(
    client,
    session,
    doctor_b,
    department_admin_headers,
):
    """Tests that a department admin cannot create unavailability for a foreign doctor."""
    response = client.post(
        f"/api/v1/doctors/{doctor_b.id}/unavailability",
        json={"date": str(datetime.date(2026, 9, 10))},
        headers=department_admin_headers,
    )

    assert response.status_code == 404
    assert response.json() == {"detail": "Doctor not found."}
    assert response.headers.get("WWW-Authenticate") is None
    assert doctor_repository.get_doctor_unavailability_by_date(
        session=session,
        doctor_id=doctor_b.id,
        target_date=datetime.date(2026, 9, 10),
    ) is None


def test_get_doctor_unavailability_route_hides_foreign_doctor(
    client,
    doctor_b,
    department_admin_headers,
):
    """Tests that a department admin cannot list unavailability for a foreign doctor."""
    response = client.get(
        f"/api/v1/doctors/{doctor_b.id}/unavailability",
        headers=department_admin_headers,
    )

    assert response.status_code == 404
    assert response.json() == {"detail": "Doctor not found."}
    assert response.headers.get("WWW-Authenticate") is None


def test_create_doctor_unavailability_route_rejects_doctor_targeting_other_doctor(
    client,
    session,
    other_doctor,
    doctor_headers,
):
    """Tests that a doctor cannot create unavailability for another same-department doctor."""
    response = client.post(
        f"/api/v1/doctors/{other_doctor.id}/unavailability",
        json={"date": str(datetime.date(2026, 9, 11))},
        headers=doctor_headers,
    )

    assert response.status_code == 403
    assert response.json() == {
        "detail": "Cannot access another doctor's unavailability."}
    assert response.headers.get("WWW-Authenticate") is None
    assert doctor_repository.get_doctor_unavailability_by_date(
        session=session,
        doctor_id=other_doctor.id,
        target_date=datetime.date(2026, 9, 11),
    ) is None


def test_get_doctor_unavailability_route_rejects_doctor_targeting_other_doctor(
    client,
    other_doctor,
    doctor_headers,
):
    """Tests that a doctor cannot list another same-department doctor's unavailability."""
    response = client.get(
        f"/api/v1/doctors/{other_doctor.id}/unavailability",
        headers=doctor_headers,
    )

    assert response.status_code == 403
    assert response.json() == {
        "detail": "Cannot access another doctor's unavailability."}
    assert response.headers.get("WWW-Authenticate") is None


def test_create_doctor_unavailability_route_rejects_doctor_targeting_foreign_doctor(
    client,
    session,
    doctor_b,
    doctor_headers,
):
    """Tests that a doctor cannot create unavailability for a foreign-department doctor."""
    response = client.post(
        f"api/v1/doctors/{doctor_b.id}/unavailability",
        json={"date": str(datetime.date(2026, 9, 12))},
        headers=doctor_headers,
    )

    assert response.status_code == 403
    assert response.json() == {
        "detail": "Cannot access another doctor's unavailability."}
    assert response.headers.get("WWW-Authenticate") is None
    assert doctor_repository.get_doctor_unavailability_by_date(
        session=session,
        doctor_id=doctor_b.id,
        target_date=datetime.date(2026, 9, 12),
    ) is None


def test_get_doctor_unavailability_route_rejects_doctor_targeting_foreign_doctor(
    client,
    doctor_b,
    doctor_headers,
):
    """Tests that a doctor cannot list a foreign-department doctor's unavailability."""
    response = client.get(
        f"api/v1/doctors/{doctor_b.id}/unavailability",
        headers=doctor_headers,
    )

    assert response.status_code == 403
    assert response.json() == {
        "detail": "Cannot access another doctor's unavailability."}
    assert response.headers.get("WWW-Authenticate") is None


def test_create_doctor_unavailability_route_checks_ownership_before_duplicate(
    client,
    session,
    other_doctor,
    doctor_headers,
):
    """Tests that ownership is enforced before the duplicate-date check at the route level."""
    existing_unavailability = create_test_unavailability(
        session, other_doctor, datetime.date(2026, 9, 15)
    )

    response = client.post(
        f"api/v1/doctors/{other_doctor.id}/unavailability",
        json={"date": str(datetime.date(2026, 9, 15))},
        headers=doctor_headers,
    )

    assert response.status_code == 403
    assert response.json() == {
        "detail": "Cannot access another doctor's unavailability."}
    assert response.headers.get("WWW-Authenticate") is None

    stored_unavailabilities = doctor_repository.get_doctor_unavailability(
        session=session,
        doctor_id=other_doctor.id,
    )
    assert len(stored_unavailabilities) == 1
    assert stored_unavailabilities[0].id == existing_unavailability.id


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


def test_create_doctor_position_route_hides_foreign_doctor(
    client,
    session,
    doctor_b,
    position,
    department_admin_headers,
):
    """Tests that an admin cannot assign a Position to a foreign doctor."""
    response = client.post(
        f"/api/v1/doctors/{doctor_b.id}/position",
        json={"position_id": str(position.id)},
        headers=department_admin_headers,
    )

    assert response.status_code == 404
    assert response.json() == {"detail": "Doctor or position not found."}
    assert response.headers.get("WWW-Authenticate") is None
    assert doctor_repository.get_doctor_position_by_id(
        session=session,
        doctor_id=doctor_b.id,
        position_id=position.id,
    ) is None


def test_create_doctor_position_route_hides_foreign_position(
    client,
    session,
    new_doctor,
    position_b,
    department_admin_headers,
):
    """Tests that an admin cannot assign a foreign Position to a doctor."""
    response = client.post(
        f"/api/v1/doctors/{new_doctor.id}/position",
        json={"position_id": str(position_b.id)},
        headers=department_admin_headers,
    )

    assert response.status_code == 404
    assert response.json() == {"detail": "Doctor or position not found."}
    assert response.headers.get("WWW-Authenticate") is None
    assert doctor_repository.get_doctor_position_by_id(
        session=session,
        doctor_id=new_doctor.id,
        position_id=position_b.id,
    ) is None


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
        UserRole.DEPARTMENT_ADMIN,
        UserRole.DOCTOR,
        UserRole.VIEWER,
    ]
)
def test_list_doctor_positions_route_hides_foreign_doctor(
    client,
    session,
    department,
    new_doctor,
    doctor_b,
    position_b,
    role,
    user_factory,
    auth_headers_factory,
):
    """Tests that Department A members cannot list a Department B doctor's associations."""
    create_test_doctor_position(
        session=session,
        doctor=doctor_b,
        position=position_b,
    )

    user = user_factory(
        role=role,
        department_id=department.id,
        doctor_id=new_doctor.id if role == UserRole.DOCTOR else None,
    )

    response = client.get(
        f"api/v1/doctors/{doctor_b.id}/position",
        headers=auth_headers_factory(user),
    )

    assert response.status_code == 404
    assert response.json() == {"detail": "Doctor not found."}
    assert response.headers.get("WWW-Authenticate") is None


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
    new_doctor,
):
    """Tests post /api/v1/doctors route with non department admin headers"""
    user = user_factory(
        role=role,
        department_id=(
            None
            if role == UserRole.SUPER_ADMIN
            else department.id
        ),
        doctor_id=new_doctor.id if role == UserRole.DOCTOR else None,
    )
    headers = auth_headers_factory(user)

    response = client.post(
        "/api/v1/doctors",
        json={
            "name": "Test Doctor",
            "team_id": str(team.id),
        },
        headers=headers,
    )

    assert response.status_code == 403
    assert response.json() == {
        "detail": "Insufficient permissions for this operation."}
    assert response.headers.get("WWW-Authenticate") is None
    assert doctor_repository.get_doctor_by_email(
        session=session,
        email="test@example.com",
    ) is None


def test_create_doctor_requires_authentication(client, session, department, team):
    """Tests the post /api/v1/doctors route without auth"""
    response = client.post(
        "/api/v1/doctors",
        json={
            "name": "Dr Panos",
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
        pytest.param(
            "/api/v1/doctors/roster",
            id="doctor-roster",
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
        "detail": "Insufficient permissions for this operation."}
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
        "detail": "Insufficient permissions for this operation."}
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
        "detail": "Insufficient permissions for this operation."}
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
        "detail": "Insufficient permissions for this operation."}
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
        "detail": "Insufficient permissions for this operation."
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


def test_get_doctor_emails_for_department(
    session, department, department_b, new_doctor, other_doctor, doctor_b, user_factory
):
    """Tests batch retrieval of doctor emails for a department, excluding deleted and foreign users."""
    # Active user linked to new_doctor in department
    user_factory(
        email="panos@hospital.com",
        role=UserRole.DOCTOR,
        department_id=department.id,
        doctor_id=new_doctor.id,
    )
    # Active user linked to doctor_b in foreign department_b
    user_factory(
        email="foreign@hospital.com",
        role=UserRole.DOCTOR,
        department_id=department_b.id,
        doctor_id=doctor_b.id,
    )
    # other_doctor has no User account

    # Soft-deleted user linked to another doctor
    doc_deleted = doctor_repository.create_doctor(
        session=session, name="Dr Deleted", department_id=department.id
    )
    user_deleted = user_factory(
        email="deleted@hospital.com",
        role=UserRole.DOCTOR,
        department_id=department.id,
        doctor_id=doc_deleted.id,
    )
    user_deleted.is_deleted = True
    session.add(user_deleted)
    session.commit()

    emails = doctor_repository.get_doctor_emails_for_department(session, department.id)

    assert emails == {new_doctor.id: "panos@hospital.com"}
    assert other_doctor.id not in emails
    assert doctor_b.id not in emails
    assert doc_deleted.id not in emails


def test_get_all_doctor_positions_for_department(
    session, department, department_b, new_doctor, other_doctor, doctor_b, position, position_b
):
    """Tests batch retrieval of doctor positions for a department, excluding deleted and foreign records."""
    # 2 active assignments in department
    dp1 = create_test_doctor_position(session, new_doctor, position)
    dp2 = create_test_doctor_position(session, other_doctor, position)

    # 1 in foreign department
    create_test_doctor_position(session, doctor_b, position_b)

    # 1 soft-deleted assignment in department (using a 3rd doctor)
    doc_3 = doctor_repository.create_doctor(
        session=session, name="Dr Position Deleted", department_id=department.id
    )
    dp_deleted = create_test_doctor_position(session, doc_3, position)
    dp_deleted.is_deleted = True
    session.add(dp_deleted)
    session.commit()

    results = doctor_repository.get_all_doctor_positions_for_department(session, department.id)
    result_ids = {r.id for r in results}

    assert len(results) == 2
    assert dp1.id in result_ids
    assert dp2.id in result_ids
    assert dp_deleted.id not in result_ids


def test_get_department_unavailabilities_for_month(
    session, department, department_b, new_doctor, doctor_b
):
    """Tests batch retrieval of doctor unavailabilities strictly filtered by department and target month."""
    # In-month unavailabilities (November 2026)
    u1 = create_test_unavailability(session, new_doctor, datetime.date(2026, 11, 5))
    u2 = create_test_unavailability(session, new_doctor, datetime.date(2026, 11, 20))

    # Out-of-month (October and December)
    create_test_unavailability(session, new_doctor, datetime.date(2026, 10, 31))
    create_test_unavailability(session, new_doctor, datetime.date(2026, 12, 1))

    # In-month but soft-deleted
    u_deleted = create_test_unavailability(session, new_doctor, datetime.date(2026, 11, 15))
    u_deleted.is_deleted = True
    session.add(u_deleted)
    session.commit()

    # In-month but foreign department
    create_test_unavailability(session, doctor_b, datetime.date(2026, 11, 10))

    results = doctor_repository.get_department_unavailabilities_for_month(
        session, department.id, year=2026, month=11
    )
    result_ids = {r.id for r in results}

    assert len(results) == 2
    assert u1.id in result_ids
    assert u2.id in result_ids
    assert u_deleted.id not in result_ids


def test_get_department_pre_assignments_for_month(
    session, department, department_b, new_doctor, doctor_b, shift, shift_b
):
    """Tests batch retrieval of pre-assignments strictly filtered by department and target month."""
    # In-month pre-assignment (November 2026)
    pa1 = create_test_pre_assignment(session, shift.id, new_doctor, datetime.date(2026, 11, 3))

    # Out-of-month (October and December)
    create_test_pre_assignment(session, shift.id, new_doctor, datetime.date(2026, 10, 31))
    create_test_pre_assignment(session, shift.id, new_doctor, datetime.date(2026, 12, 1))

    # In-month but soft-deleted
    pa_deleted = create_test_pre_assignment(session, shift.id, new_doctor, datetime.date(2026, 11, 15))
    pa_deleted.is_deleted = True
    session.add(pa_deleted)
    session.commit()

    # In-month but foreign department
    create_test_pre_assignment(session, shift_b.id, doctor_b, datetime.date(2026, 11, 7))

    results = doctor_repository.get_department_pre_assignments_for_month(
        session, department.id, year=2026, month=11
    )
    result_ids = {r.id for r in results}

    assert len(results) == 1
    assert pa1.id in result_ids
    assert pa_deleted.id not in result_ids


#####################
# Step 5 Task 5.1 Slice 4:
# Doctor PATCH/DELETE, Sub-resource DELETEs, Restore-in-Place & DoctorRead.position_ids
#####################


# --- Schemas ---


def test_doctor_update_schema_valid():
    """DoctorUpdate accepts name, team_id (including explicit None), or empty payload."""
    team_id = uuid.uuid4()

    rename_only = DoctorUpdate(name="Dr Renamed")
    assert rename_only.model_dump(exclude_unset=True) == {"name": "Dr Renamed"}

    unassign_team = DoctorUpdate(team_id=None)
    assert unassign_team.model_dump(exclude_unset=True) == {"team_id": None}

    reassign_team = DoctorUpdate(team_id=team_id)
    assert reassign_team.model_dump(exclude_unset=True) == {"team_id": team_id}

    empty = DoctorUpdate()
    assert empty.model_dump(exclude_unset=True) == {}


def test_doctor_update_schema_rejects_department_id_and_extra_fields():
    """DoctorUpdate forbids department_id and arbitrary extra fields."""
    with pytest.raises(ValidationError):
        DoctorUpdate(department_id=uuid.uuid4())

    with pytest.raises(ValidationError):
        DoctorUpdate(unknown_field="unexpected")


def test_doctor_read_schema_defaults_position_ids():
    """DoctorRead defaults position_ids to an empty list."""
    now = datetime.datetime.now(datetime.timezone.utc)
    schema = DoctorRead(
        id=uuid.uuid4(),
        name="Dr Schema",
        department_id=uuid.uuid4(),
        created_at=now,
        updated_at=now,
    )
    assert schema.position_ids == []


# --- DoctorRead.position_ids in Doctor Responses & Roster Exclusion ---


def test_list_doctors_route_includes_position_ids_and_excludes_deleted(
    client,
    session,
    department,
    new_doctor,
    other_doctor,
    position,
    department_admin_headers,
):
    """GET /api/v1/doctors populates active position_ids and excludes soft-deleted links or positions."""
    deleted_link_pos = position_repository.create_position(
        session=session,
        position_name="ICU",
        department_id=department.id,
        duty_days=[1, 2],
    )
    deleted_pos = position_repository.create_position(
        session=session,
        position_name="OnCall",
        department_id=department.id,
        duty_days=[4, 5],
    )

    create_test_doctor_position(session, new_doctor, position)
    deleted_link = create_test_doctor_position(session, new_doctor, deleted_link_pos)
    deleted_link.is_deleted = True
    session.add(deleted_link)

    create_test_doctor_position(session, new_doctor, deleted_pos)
    deleted_pos.is_deleted = True
    session.add(deleted_pos)
    session.commit()

    response = client.get("api/v1/doctors", headers=department_admin_headers)
    assert response.status_code == 200

    doctors_by_id = {item["id"]: item for item in response.json()}
    assert doctors_by_id[str(new_doctor.id)]["position_ids"] == [str(position.id)]
    assert doctors_by_id[str(other_doctor.id)]["position_ids"] == []


def test_get_doctor_route_includes_position_ids(
    client,
    session,
    new_doctor,
    position,
    department_admin_headers,
):
    """GET /api/v1/doctors/{doctor_id} returns active position_ids."""
    create_test_doctor_position(session, new_doctor, position)

    response = client.get(
        f"api/v1/doctors/{new_doctor.id}",
        headers=department_admin_headers,
    )
    assert response.status_code == 200
    assert response.json()["position_ids"] == [str(position.id)]


def test_create_doctor_route_includes_empty_position_ids(
    client,
    team,
    department_admin_headers,
):
    """POST /api/v1/doctors returns position_ids defaulting to []."""
    response = client.post(
        "api/v1/doctors",
        json={"name": "Dr Fresh", "team_id": str(team.id)},
        headers=department_admin_headers,
    )
    assert response.status_code == 201
    assert response.json()["position_ids"] == []


def test_list_doctors_roster_route_excludes_position_ids(
    client,
    session,
    new_doctor,
    position,
    viewer_headers,
):
    """GET /api/v1/doctors/roster keeps minimal DoctorRosterRead shape without position_ids."""
    create_test_doctor_position(session, new_doctor, position)

    response = client.get("api/v1/doctors/roster", headers=viewer_headers)
    assert response.status_code == 200
    data = response.json()
    assert len(data) == 1
    assert set(data[0].keys()) == {"id", "name", "department_id", "team_id"}


# --- PATCH /api/v1/doctors/{doctor_id} ---


def test_update_doctor_route_rename_keeps_team_and_position_ids(
    client,
    session,
    new_doctor,
    team,
    position,
    department_admin_headers,
):
    """PATCH /doctors/{id} renaming only leaves team_id untouched and returns position_ids."""
    create_test_doctor_position(session, new_doctor, position)
    new_doctor.sync_status = True
    session.add(new_doctor)
    session.commit()

    response = client.patch(
        f"api/v1/doctors/{new_doctor.id}",
        json={"name": "Dr Panos Updated"},
        headers=department_admin_headers,
    )

    assert response.status_code == 200
    data = response.json()
    assert data["name"] == "Dr Panos Updated"
    assert data["team_id"] == str(team.id)
    assert data["position_ids"] == [str(position.id)]
    assert data["sync_status"] is False

    session.refresh(new_doctor)
    assert new_doctor.name == "Dr Panos Updated"
    assert new_doctor.team_id == team.id
    assert new_doctor.sync_status is False


def test_update_doctor_route_reassign_team(
    client,
    session,
    department,
    new_doctor,
    department_admin_headers,
):
    """PATCH /doctors/{id} can reassign a doctor to another active team in the same department."""
    team_two = create_team(session=session, name="ER Team B", department_id=department.id)

    response = client.patch(
        f"api/v1/doctors/{new_doctor.id}",
        json={"team_id": str(team_two.id)},
        headers=department_admin_headers,
    )

    assert response.status_code == 200
    assert response.json()["team_id"] == str(team_two.id)
    session.refresh(new_doctor)
    assert new_doctor.team_id == team_two.id


def test_update_doctor_route_unassign_team_with_null(
    client,
    session,
    new_doctor,
    department_admin_headers,
):
    """PATCH /doctors/{id} with {"team_id": null} unassigns the doctor's team."""
    assert new_doctor.team_id is not None

    response = client.patch(
        f"api/v1/doctors/{new_doctor.id}",
        json={"team_id": None},
        headers=department_admin_headers,
    )

    assert response.status_code == 200
    assert response.json()["team_id"] is None
    session.refresh(new_doctor)
    assert new_doctor.team_id is None


def test_update_doctor_route_empty_payload_idempotent(
    client,
    new_doctor,
    team,
    department_admin_headers,
):
    """PATCH /doctors/{id} with {} returns 200 and preserves existing name and team_id."""
    response = client.patch(
        f"api/v1/doctors/{new_doctor.id}",
        json={},
        headers=department_admin_headers,
    )

    assert response.status_code == 200
    data = response.json()
    assert data["name"] == "Dr Panos"
    assert data["team_id"] == str(team.id)


def test_update_doctor_route_rejects_department_id(
    client,
    new_doctor,
    department_b,
    department_admin_headers,
):
    """PATCH /doctors/{id} rejects client-supplied department_id with 422."""
    response = client.patch(
        f"api/v1/doctors/{new_doctor.id}",
        json={"department_id": str(department_b.id)},
        headers=department_admin_headers,
    )
    assert response.status_code == 422


def test_update_doctor_route_invalid_team_returns_404(
    client,
    session,
    department,
    new_doctor,
    team_b,
    department_admin_headers,
):
    """PATCH /doctors/{id} returns 404 for nonexistent, foreign, or soft-deleted team_id."""
    deleted_team = create_team(session=session, name="Deleted Team", department_id=department.id)
    deleted_team.is_deleted = True
    session.add(deleted_team)
    session.commit()

    for bad_team_id in (uuid.uuid4(), team_b.id, deleted_team.id):
        response = client.patch(
            f"api/v1/doctors/{new_doctor.id}",
            json={"team_id": str(bad_team_id)},
            headers=department_admin_headers,
        )
        assert response.status_code == 404


def test_update_doctor_route_missing_deleted_or_foreign_doctor_returns_404(
    client,
    session,
    new_doctor,
    doctor_b,
    department_admin_headers,
):
    """PATCH /doctors/{id} returns 404 for unknown, soft-deleted, or cross-department doctor."""
    new_doctor.is_deleted = True
    session.add(new_doctor)
    session.commit()

    for bad_doctor_id in (uuid.uuid4(), new_doctor.id, doctor_b.id):
        response = client.patch(
            f"api/v1/doctors/{bad_doctor_id}",
            json={"name": "Dr Nope"},
            headers=department_admin_headers,
        )
        assert response.status_code == 404


def test_update_doctor_route_auth_guards(
    client,
    new_doctor,
    viewer_headers,
    doctor_headers,
):
    """PATCH /doctors/{id} requires authentication (401) and DEPARTMENT_ADMIN role (403)."""
    unauth = client.patch(
        f"api/v1/doctors/{new_doctor.id}",
        json={"name": "Blocked"},
    )
    assert unauth.status_code == 401

    viewer_resp = client.patch(
        f"api/v1/doctors/{new_doctor.id}",
        json={"name": "Blocked"},
        headers=viewer_headers,
    )
    assert viewer_resp.status_code == 403

    doctor_resp = client.patch(
        f"api/v1/doctors/{new_doctor.id}",
        json={"name": "Blocked"},
        headers=doctor_headers,
    )
    assert doctor_resp.status_code == 403


# --- DELETE /api/v1/doctors/{doctor_id} ---


def test_delete_doctor_route_soft_deletes_doctor_and_positions_keeps_user_active(
    client,
    session,
    new_doctor,
    position,
    doctor_user,
    department_admin_headers,
    viewer_headers,
):
    """DELETE /doctors/{id} soft-deletes Doctor and active DoctorPosition links, keeping User active."""
    doc_pos = create_test_doctor_position(session, new_doctor, position)
    new_doctor.sync_status = True
    doc_pos.sync_status = True
    session.add(new_doctor)
    session.add(doc_pos)
    session.commit()

    response = client.delete(
        f"api/v1/doctors/{new_doctor.id}",
        headers=department_admin_headers,
    )
    assert response.status_code == 204
    assert response.content == b""

    session.refresh(new_doctor)
    session.refresh(doc_pos)
    session.refresh(doctor_user)
    assert new_doctor.is_deleted is True
    assert new_doctor.sync_status is False
    assert doc_pos.is_deleted is True
    assert doc_pos.sync_status is False
    assert doctor_user.is_deleted is False

    list_resp = client.get("api/v1/doctors", headers=department_admin_headers)
    assert list_resp.status_code == 200
    assert all(d["id"] != str(new_doctor.id) for d in list_resp.json())

    roster_resp = client.get("api/v1/doctors/roster", headers=viewer_headers)
    assert roster_resp.status_code == 200
    assert all(d["id"] != str(new_doctor.id) for d in roster_resp.json())


def test_delete_doctor_route_missing_deleted_or_foreign_returns_404(
    client,
    session,
    new_doctor,
    doctor_b,
    department_admin_headers,
):
    """DELETE /doctors/{id} returns 404 for unknown, already deleted, or foreign-department doctor."""
    new_doctor.is_deleted = True
    session.add(new_doctor)
    session.commit()

    for bad_doctor_id in (uuid.uuid4(), new_doctor.id, doctor_b.id):
        response = client.delete(
            f"api/v1/doctors/{bad_doctor_id}",
            headers=department_admin_headers,
        )
        assert response.status_code == 404


def test_delete_doctor_route_auth_guards(
    client,
    new_doctor,
    viewer_headers,
    doctor_headers,
):
    """DELETE /doctors/{id} requires authentication (401) and DEPARTMENT_ADMIN role (403)."""
    assert client.delete(f"api/v1/doctors/{new_doctor.id}").status_code == 401
    assert (
        client.delete(
            f"api/v1/doctors/{new_doctor.id}", headers=viewer_headers
        ).status_code
        == 403
    )
    assert (
        client.delete(
            f"api/v1/doctors/{new_doctor.id}", headers=doctor_headers
        ).status_code
        == 403
    )


# --- DELETE /api/v1/doctors/{doctor_id}/position/{position_id} & Restore-in-Place ---


def test_delete_doctor_position_route_and_restore_in_place(
    client,
    session,
    new_doctor,
    position,
    department_admin_headers,
):
    """DELETE /doctors/{id}/position/{position_id} soft-deletes the link; re-adding restores in place."""
    link = create_test_doctor_position(session, new_doctor, position)
    original_link_id = link.id
    link.sync_status = True
    session.add(link)
    session.commit()

    delete_resp = client.delete(
        f"api/v1/doctors/{new_doctor.id}/position/{position.id}",
        headers=department_admin_headers,
    )
    assert delete_resp.status_code == 204
    assert delete_resp.content == b""

    session.refresh(link)
    assert link.is_deleted is True
    assert link.sync_status is False

    # Position listing and DoctorRead.position_ids omit soft-deleted link
    pos_list_resp = client.get(
        f"api/v1/doctors/{new_doctor.id}/position",
        headers=department_admin_headers,
    )
    assert pos_list_resp.status_code == 200
    assert pos_list_resp.json() == []

    doc_resp = client.get(
        f"api/v1/doctors/{new_doctor.id}",
        headers=department_admin_headers,
    )
    assert doc_resp.status_code == 200
    assert doc_resp.json()["position_ids"] == []

    # Deleting already-deleted link returns 404
    second_delete = client.delete(
        f"api/v1/doctors/{new_doctor.id}/position/{position.id}",
        headers=department_admin_headers,
    )
    assert second_delete.status_code == 404

    # Re-assigning the same position restores the existing link row in place
    restore_resp = client.post(
        f"api/v1/doctors/{new_doctor.id}/position",
        json={"position_id": str(position.id)},
        headers=department_admin_headers,
    )
    assert restore_resp.status_code == 201
    restored_data = restore_resp.json()
    assert restored_data["id"] == str(original_link_id)
    assert restored_data["is_deleted"] is False
    assert restored_data["sync_status"] is False


def test_delete_doctor_position_route_missing_or_foreign_returns_404(
    client,
    new_doctor,
    doctor_b,
    position,
    position_b,
    department_admin_headers,
):
    """DELETE /doctors/{id}/position/{pos_id} returns 404 when link, doctor, or position is missing/foreign."""
    assert (
        client.delete(
            f"api/v1/doctors/{uuid.uuid4()}/position/{position.id}",
            headers=department_admin_headers,
        ).status_code
        == 404
    )
    assert (
        client.delete(
            f"api/v1/doctors/{new_doctor.id}/position/{uuid.uuid4()}",
            headers=department_admin_headers,
        ).status_code
        == 404
    )
    assert (
        client.delete(
            f"api/v1/doctors/{doctor_b.id}/position/{position_b.id}",
            headers=department_admin_headers,
        ).status_code
        == 404
    )


def test_delete_doctor_position_route_auth_guards(
    client,
    session,
    new_doctor,
    position,
    viewer_headers,
    doctor_headers,
):
    """DELETE /doctors/{id}/position/{pos_id} requires authentication (401) and DEPARTMENT_ADMIN (403)."""
    create_test_doctor_position(session, new_doctor, position)
    url = f"api/v1/doctors/{new_doctor.id}/position/{position.id}"

    assert client.delete(url).status_code == 401
    assert client.delete(url, headers=viewer_headers).status_code == 403
    assert client.delete(url, headers=doctor_headers).status_code == 403


# --- DELETE /api/v1/doctors/{doctor_id}/unavailability/{unavailability_id} & Soft-Delete Filters / Restore ---


def test_delete_doctor_unavailability_route_admin_and_doctor_self(
    client,
    session,
    new_doctor,
    unavailability,
    department_admin_headers,
    doctor_headers,
):
    """Both DEPARTMENT_ADMIN and the owning DOCTOR can delete an unavailability entry."""
    unavailability.sync_status = True
    session.add(unavailability)
    session.commit()

    admin_del = client.delete(
        f"api/v1/doctors/{new_doctor.id}/unavailability/{unavailability.id}",
        headers=department_admin_headers,
    )
    assert admin_del.status_code == 204
    assert admin_del.content == b""

    session.refresh(unavailability)
    assert unavailability.is_deleted is True
    assert unavailability.sync_status is False

    list_resp = client.get(
        f"api/v1/doctors/{new_doctor.id}/unavailability",
        headers=department_admin_headers,
    )
    assert list_resp.status_code == 200
    assert list_resp.json() == []

    # Doctor deletes their own newly created unavailability
    self_unavail = create_test_unavailability(
        session, new_doctor, datetime.date(2026, 8, 20)
    )
    doc_del = client.delete(
        f"api/v1/doctors/{new_doctor.id}/unavailability/{self_unavail.id}",
        headers=doctor_headers,
    )
    assert doc_del.status_code == 204
    session.refresh(self_unavail)
    assert self_unavail.is_deleted is True


def test_delete_doctor_unavailability_route_doctor_cannot_delete_other_doctor_returns_403(
    client,
    session,
    other_doctor,
    doctor_headers,
):
    """A DOCTOR cannot delete another doctor's unavailability entry."""
    other_unavail = create_test_unavailability(
        session, other_doctor, datetime.date(2026, 8, 15)
    )

    response = client.delete(
        f"api/v1/doctors/{other_doctor.id}/unavailability/{other_unavail.id}",
        headers=doctor_headers,
    )
    assert response.status_code == 403


def test_delete_doctor_unavailability_route_missing_deleted_wrong_doctor_or_foreign_returns_404(
    client,
    session,
    new_doctor,
    other_doctor,
    doctor_b,
    unavailability,
    department_admin_headers,
):
    """DELETE unavailability returns 404 for unknown, deleted, wrong-doctor, or cross-department IDs."""
    other_unavail = create_test_unavailability(
        session, other_doctor, datetime.date(2026, 8, 16)
    )
    foreign_unavail = create_test_unavailability(
        session, doctor_b, datetime.date(2026, 8, 16)
    )
    unavailability.is_deleted = True
    session.add(unavailability)
    session.commit()

    # Unknown ID
    assert (
        client.delete(
            f"api/v1/doctors/{new_doctor.id}/unavailability/{uuid.uuid4()}",
            headers=department_admin_headers,
        ).status_code
        == 404
    )
    # Already soft-deleted ID
    assert (
        client.delete(
            f"api/v1/doctors/{new_doctor.id}/unavailability/{unavailability.id}",
            headers=department_admin_headers,
        ).status_code
        == 404
    )
    # Valid unavailability owned by other_doctor referenced under new_doctor's URL
    assert (
        client.delete(
            f"api/v1/doctors/{new_doctor.id}/unavailability/{other_unavail.id}",
            headers=department_admin_headers,
        ).status_code
        == 404
    )
    # Foreign department doctor
    assert (
        client.delete(
            f"api/v1/doctors/{doctor_b.id}/unavailability/{foreign_unavail.id}",
            headers=department_admin_headers,
        ).status_code
        == 404
    )


def test_delete_doctor_unavailability_route_auth_guards(
    client,
    new_doctor,
    unavailability,
    viewer_headers,
):
    """DELETE unavailability requires authentication (401) and forbids VIEWER (403)."""
    url = f"api/v1/doctors/{new_doctor.id}/unavailability/{unavailability.id}"
    assert client.delete(url).status_code == 401
    assert client.delete(url, headers=viewer_headers).status_code == 403


def test_deleted_unavailability_does_not_block_pre_assignment_and_restores_in_place(
    client,
    session,
    new_doctor,
    shift,
    unavailability,
    department_admin_headers,
):
    """Soft-deleted unavailability does not block pre-assignment on the same date and restores in place."""
    original_id = unavailability.id
    target_date = unavailability.date

    del_resp = client.delete(
        f"api/v1/doctors/{new_doctor.id}/unavailability/{original_id}",
        headers=department_admin_headers,
    )
    assert del_resp.status_code == 204

    assert (
        doctor_repository.get_doctor_unavailability_by_date(
            session, new_doctor.id, target_date
        )
        is None
    )

    # Deleted unavailability must NOT trigger 422 conflict when adding a pre-assignment
    pa_resp = client.post(
        f"api/v1/doctors/{new_doctor.id}/pre-assignments",
        json={"date": target_date.isoformat(), "shift_id": str(shift.id)},
        headers=department_admin_headers,
    )
    assert pa_resp.status_code == 201
    pa_id = pa_resp.json()["id"]

    # Remove that pre-assignment so we can restore the soft-deleted unavailability row
    pa_del_resp = client.delete(
        f"api/v1/doctors/{new_doctor.id}/pre-assignments/{pa_id}",
        headers=department_admin_headers,
    )
    assert pa_del_resp.status_code == 204

    restore_resp = client.post(
        f"api/v1/doctors/{new_doctor.id}/unavailability",
        json={"date": target_date.isoformat()},
        headers=department_admin_headers,
    )
    assert restore_resp.status_code == 201
    restored = restore_resp.json()
    assert restored["id"] == str(original_id)
    assert restored["is_deleted"] is False
    assert restored["sync_status"] is False


# --- DELETE /api/v1/doctors/{doctor_id}/pre-assignments/{pre_assignment_id} & Soft-Delete Filters / Restore ---


def test_delete_doctor_pre_assignment_route_soft_deletes_and_excludes_from_list(
    client,
    session,
    new_doctor,
    pre_assignment,
    department_admin_headers,
):
    """DELETE pre-assignment soft-deletes the row and excludes it from active queries and GET lists."""
    pre_assignment.sync_status = True
    session.add(pre_assignment)
    session.commit()

    response = client.delete(
        f"api/v1/doctors/{new_doctor.id}/pre-assignments/{pre_assignment.id}",
        headers=department_admin_headers,
    )
    assert response.status_code == 204
    assert response.content == b""

    session.refresh(pre_assignment)
    assert pre_assignment.is_deleted is True
    assert pre_assignment.sync_status is False

    assert (
        doctor_repository.get_doctor_pre_assignment_by_date(
            session, new_doctor.id, pre_assignment.date
        )
        is None
    )

    list_resp = client.get(
        f"api/v1/doctors/{new_doctor.id}/pre-assignments",
        headers=department_admin_headers,
    )
    assert list_resp.status_code == 200
    assert list_resp.json() == []


def test_delete_doctor_pre_assignment_route_missing_deleted_wrong_doctor_or_foreign_returns_404(
    client,
    session,
    new_doctor,
    other_doctor,
    doctor_b,
    shift,
    shift_b,
    pre_assignment,
    department_admin_headers,
):
    """DELETE pre-assignment returns 404 for unknown, deleted, wrong-doctor, or foreign IDs."""
    other_pa = create_test_pre_assignment(
        session, shift.id, other_doctor, datetime.date(2026, 8, 17)
    )
    foreign_pa = create_test_pre_assignment(
        session, shift_b.id, doctor_b, datetime.date(2026, 8, 17)
    )
    pre_assignment.is_deleted = True
    session.add(pre_assignment)
    session.commit()

    assert (
        client.delete(
            f"api/v1/doctors/{new_doctor.id}/pre-assignments/{uuid.uuid4()}",
            headers=department_admin_headers,
        ).status_code
        == 404
    )
    assert (
        client.delete(
            f"api/v1/doctors/{new_doctor.id}/pre-assignments/{pre_assignment.id}",
            headers=department_admin_headers,
        ).status_code
        == 404
    )
    assert (
        client.delete(
            f"api/v1/doctors/{new_doctor.id}/pre-assignments/{other_pa.id}",
            headers=department_admin_headers,
        ).status_code
        == 404
    )
    assert (
        client.delete(
            f"api/v1/doctors/{doctor_b.id}/pre-assignments/{foreign_pa.id}",
            headers=department_admin_headers,
        ).status_code
        == 404
    )


def test_delete_doctor_pre_assignment_route_auth_guards(
    client,
    new_doctor,
    pre_assignment,
    viewer_headers,
    doctor_headers,
):
    """DELETE pre-assignment requires authentication (401) and DEPARTMENT_ADMIN role (403)."""
    url = f"api/v1/doctors/{new_doctor.id}/pre-assignments/{pre_assignment.id}"
    assert client.delete(url).status_code == 401
    assert client.delete(url, headers=viewer_headers).status_code == 403
    assert client.delete(url, headers=doctor_headers).status_code == 403


def test_deleted_pre_assignment_does_not_block_unavailability_and_restores_in_place_with_new_shift(
    client,
    session,
    new_doctor,
    position,
    pre_assignment,
    department_admin_headers,
):
    """Soft-deleted pre-assignment does not block unavailability and restores in place with updated shift_id."""
    original_id = pre_assignment.id
    target_date = pre_assignment.date
    second_shift = shift_repository.create_shift(
        session,
        ShiftCreate(
            name="ER 2",
            position_id=position.id,
            grants_day_off=True,
            doctors_per_shift=1,
        ),
    )

    del_resp = client.delete(
        f"api/v1/doctors/{new_doctor.id}/pre-assignments/{original_id}",
        headers=department_admin_headers,
    )
    assert del_resp.status_code == 204

    # Deleted pre-assignment must NOT trigger 422 conflict when adding unavailability
    unavail_resp = client.post(
        f"api/v1/doctors/{new_doctor.id}/unavailability",
        json={"date": target_date.isoformat()},
        headers=department_admin_headers,
    )
    assert unavail_resp.status_code == 201
    unavail_id = unavail_resp.json()["id"]

    # Remove unavailability so we can restore the soft-deleted pre-assignment with second_shift
    assert (
        client.delete(
            f"api/v1/doctors/{new_doctor.id}/unavailability/{unavail_id}",
            headers=department_admin_headers,
        ).status_code
        == 204
    )

    restore_resp = client.post(
        f"api/v1/doctors/{new_doctor.id}/pre-assignments",
        json={"date": target_date.isoformat(), "shift_id": str(second_shift.id)},
        headers=department_admin_headers,
    )
    assert restore_resp.status_code == 201
    restored = restore_resp.json()
    assert restored["id"] == str(original_id)
    assert restored["shift_id"] == str(second_shift.id)
    assert restored["is_deleted"] is False
    assert restored["sync_status"] is False

