"""
Tests for the shift module
"""

import uuid
import datetime
import pytest
from fastapi import HTTPException
from pydantic import ValidationError
from sqlmodel import Session, select
from sqlalchemy.exc import IntegrityError

from src.shift.schemas import (
    ShiftAssignmentCreate,
    ShiftAssignmentRead,
    ShiftCreate,
    ShiftUpdate,
)
from src.shift.models import Shift as ShiftModel
from src.shift.models import ShiftAssignment as ShiftAssignmentModel
from src.shift import repository as shift_repository
from src.shift import controllers as shift_controllers
from src.position.schemas import PositionCreate
from src.position import repository as position_repository
from src.doctor.models import Doctor as DoctorModel
from src.doctor.schemas import (
    DoctorCreate,
    DoctorPreAssignmentCreate,
    DoctorUnavailabilityCreate,
)
from src.doctor import repository as doctor_repository
from src.department.schemas import DepartmentCreate
from src.department import repository as department_repository
from src.team import repository as team_repository
from src.user.models import UserRole

#####################
# Helpers
#####################


def create_new_doctor(
    session: Session,
    name: str,
    email: str = "",
    department_id: uuid.UUID = None,
    team_id: uuid.UUID = None,
    **kwargs,
) -> DoctorModel:
    """Helper that creates a new doctor in the db"""
    return doctor_repository.create_doctor(
        name=name,
        team_id=team_id,
        session=session,
        department_id=department_id,
    )


def create_new_shift(
    session: Session,
    name: str,
    position_id: uuid.UUID,
    grants_day_off: bool,
    doctors_per_shift: int
) -> ShiftModel:
    """Helper that creates a shift in the db"""

    shift_data = ShiftCreate(
        name=name,
        position_id=position_id,
        grants_day_off=grants_day_off,
        doctors_per_shift=doctors_per_shift
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

#####################
# Fixtures
#####################


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
    return position_repository.create_position(
        session=session,
        position_name="ER",
        department_id=department.id,
        duty_days=[1, 3, 5],
    )


@pytest.fixture(name="team")
def team_fixture(session, department):
    """Creates a reusable team for tests"""
    return team_repository.create_team(
        session=session,
        name="ER Team A",
        department_id=department.id
    )


@pytest.fixture(name="new_doctor")
def doctor_fixture(session, department, team):
    """Creates a reusable doctor for tests"""
    return create_new_doctor(session, "Dr Panos", "drpanos@gmail.com", department.id, team.id)


@pytest.fixture(name="shift_assignment")
def shift_assignment_fixture(session, new_doctor, shift):
    """Creates a reusable shift assignment for tests"""
    return create_new_shift_assignment(session, new_doctor.id, shift.id, datetime.date(2026, 8, 12))


@pytest.fixture(name="unavailability")
def unavailability_fixture(session, new_doctor):
    """Creates a reusable pre-assignment for tests"""
    return create_test_unavailability(session, new_doctor, datetime.date(2026, 8, 13))


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


@pytest.fixture(name="department_b")
def department_b_fixture(session):
    """Creates a second department for tenant-isolation tests."""
    dept_data = DepartmentCreate(name="Radiology", code="RAD")
    return department_repository.create_department(session, dept_data)


@pytest.fixture(name="position_b")
def position_b_fixture(session, department_b):
    """Creates a reusable Department B position for tenant-isolation tests."""
    return position_repository.create_position(
        session=session,
        position_name="ICU",
        department_id=department_b.id,
        duty_days=[2, 4, 6],
    )


@pytest.fixture(name="team_b")
def team_b_fixture(session, department_b):
    """Creates a reusable Department B team for tenant-isolation tests."""
    return team_repository.create_team(
        session=session,
        name="ICU Team B",
        department_id=department_b.id
    )


@pytest.fixture(name="doctor_b")
def doctor_b_fixture(session, department_b, team_b):
    """Creates a reusable Department B doctor for tenant-isolation tests."""
    return create_new_doctor(
        session, "Dr Radiology", "drradiology@gmail.com", department_b.id, team_b.id
    )


@pytest.fixture(name="shift_b")
def shift_b_fixture(session, position_b):
    """Creates a reusable Department B shift for tenant-isolation tests."""
    return create_new_shift(session, "ICU 1", position_b.id, False, 2)


@pytest.fixture(name="shift_factory")
def shift_factory_fixture(session, position):
    """Factory fixture to create shifts with customizable fields and flags."""

    def _create_shift(
        name: str = "Factory Shift",
        position_id: uuid.UUID | None = None,
        doctors_per_shift: int = 1,
        grants_day_off: bool = False,
        is_deleted: bool = False,
        sync_status: bool = False,
    ) -> ShiftModel:
        created = shift_repository.create_shift(
            session=session,
            shift_data=ShiftCreate(
                name=name,
                position_id=position_id or position.id,
                doctors_per_shift=doctors_per_shift,
                grants_day_off=grants_day_off,
            ),
        )
        if is_deleted or sync_status:
            created.is_deleted = is_deleted
            created.sync_status = sync_status
            session.add(created)
            session.commit()
            session.refresh(created)
        return created

    return _create_shift


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


def test_get_shift_by_name_for_position_resolves_by_position(
    session,
    shift,
    position,
    position_b,
):
    """Tests that the same shift name resolves independently under each position."""
    shift_under_b = create_new_shift(
        session=session,
        name=shift.name,
        position_id=position_b.id,
        grants_day_off=False,
        doctors_per_shift=1,
    )

    retrieved_a = shift_repository.get_shift_by_name_for_position(
        session=session,
        position_id=position.id,
        shift_name=shift.name,
    )
    retrieved_b = shift_repository.get_shift_by_name_for_position(
        session=session,
        position_id=position_b.id,
        shift_name=shift.name,
    )

    assert retrieved_a is not None
    assert retrieved_b is not None
    assert retrieved_a.id == shift.id
    assert retrieved_b.id == shift_under_b.id
    assert retrieved_a.id != retrieved_b.id


def test_get_shift_by_name_for_position_includes_deleted_for_reservation(
    session,
    shift,
):
    """Tests that a soft-deleted shift is still found to reserve its name."""
    shift.is_deleted = True
    session.add(shift)
    session.commit()

    retrieved_shift = shift_repository.get_shift_by_name_for_position(
        session=session,
        position_id=shift.position_id,
        shift_name=shift.name,
    )

    assert retrieved_shift is not None
    assert retrieved_shift.id == shift.id
    assert retrieved_shift.is_deleted is True


def test_get_shift_by_id_for_department_returns_own_active_shift(
    session,
    shift,
    position,
):
    """Tests retrieving an active Shift scoped to its own department."""
    retrieved_shift = shift_repository.get_shift_by_id_for_department(
        session=session,
        shift_id=shift.id,
        department_id=position.department_id,
    )

    assert retrieved_shift is not None
    assert retrieved_shift.id == shift.id


def test_get_shift_by_id_for_department_hides_foreign_shift(
    session,
    department,
    shift_b,
):
    """Tests that a scoped ID lookup hides Shifts from another department."""
    retrieved_shift = shift_repository.get_shift_by_id_for_department(
        session=session,
        shift_id=shift_b.id,
        department_id=department.id,
    )

    assert retrieved_shift is None


def test_get_shift_by_id_for_department_hides_deleted_shift(
    session,
    shift,
    position,
):
    """Tests that a scoped ID lookup hides deleted Shifts."""
    shift.is_deleted = True
    session.add(shift)
    session.commit()

    retrieved_shift = shift_repository.get_shift_by_id_for_department(
        session=session,
        shift_id=shift.id,
        department_id=position.department_id,
    )

    assert retrieved_shift is None


def test_get_active_shifts_for_department_excludes_foreign_and_deleted(
    session,
    department,
    position,
    shift,
    shift_b,
):
    """Tests listing only active Shifts within department scope."""
    deleted_shift = create_new_shift(session, "ER 2", position.id, False, 1)
    deleted_shift.is_deleted = True
    session.add(deleted_shift)
    session.commit()

    retrieved_shifts = shift_repository.get_active_shifts_for_department(
        session=session,
        department_id=department.id,
    )

    returned_ids = {retrieved_shift.id for retrieved_shift in retrieved_shifts}
    assert shift.id in returned_ids
    assert deleted_shift.id not in returned_ids
    assert shift_b.id not in returned_ids


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


def test_get_shift_assignments_by_date(session, shift_assignment):
    """Tests retrieving a shift assignemtnt by its date"""
    retrieved_shift_assignments = shift_repository.get_shift_assignments_by_date(
        session=session,
        shift_id=shift_assignment.shift_id,
        target_date=datetime.date(2026, 8, 12)
    )

    assert isinstance(retrieved_shift_assignments, list)
    assert shift_assignment.id == retrieved_shift_assignments[0].id


def test_get_shift_assignment_by_id(session, shift_assignment):
    """Tests retrieving a shift assignemtnt by its date"""
    retrieved_shift_assignment = shift_repository.get_shift_assignment_by_id(
        session=session,
        shift_assignment_id=shift_assignment.id
    )

    assert isinstance(retrieved_shift_assignment.id, uuid.UUID)
    assert shift_assignment.id == retrieved_shift_assignment.id


def test_get_active_shift_assignments_for_department_excludes_foreign_and_deleted(
    session,
    department,
    position,
    new_doctor,
    shift,
    shift_assignment,
    shift_b,
    doctor_b,
):
    """Tests listing only active ShiftAssignments within department scope."""
    deleted_assignment = create_new_shift_assignment(
        session=session,
        doctor_id=new_doctor.id,
        shift_id=shift.id,
        date=datetime.date(2026, 8, 13),
    )
    deleted_assignment.is_deleted = True
    session.add(deleted_assignment)
    session.commit()

    foreign_assignment = create_new_shift_assignment(
        session=session,
        doctor_id=doctor_b.id,
        shift_id=shift_b.id,
        date=datetime.date(2026, 8, 12),
    )

    deleted_shift = create_new_shift(
        session=session,
        name="Deleted Shift",
        position_id=position.id,
        grants_day_off=False,
        doctors_per_shift=1,
    )
    deleted_shift_assignment = create_new_shift_assignment(
        session=session,
        doctor_id=new_doctor.id,
        shift_id=deleted_shift.id,
        date=datetime.date(2026, 8, 14),
    )
    deleted_shift.is_deleted = True
    session.add(deleted_shift)
    session.commit()

    deleted_position = position_repository.create_position(
        session=session,
        position_name="Deleted Position",
        department_id=department.id,
        duty_days=[1, 3, 5],
    )
    deleted_position_shift = create_new_shift(
        session=session,
        name="Shift Under Deleted Position",
        position_id=deleted_position.id,
        grants_day_off=False,
        doctors_per_shift=1,
    )
    deleted_position_assignment = create_new_shift_assignment(
        session=session,
        doctor_id=new_doctor.id,
        shift_id=deleted_position_shift.id,
        date=datetime.date(2026, 8, 15),
    )
    deleted_position.is_deleted = True
    session.add(deleted_position)
    session.commit()

    retrieved_assignments = shift_repository.get_active_shift_assignments_for_department(
        session=session,
        department_id=department.id,
    )

    returned_ids = {
        retrieved_assignment.id
        for retrieved_assignment in retrieved_assignments
    }
    assert shift_assignment.id in returned_ids
    assert deleted_assignment.id not in returned_ids
    assert foreign_assignment.id not in returned_ids
    assert deleted_shift_assignment.id not in returned_ids
    assert deleted_position_assignment.id not in returned_ids


def test_shift_name_can_repeat_across_positions(
    session,
    position,
):
    """Tests that a shift name is allowed under different positions."""
    position_b = position_repository.create_position(
        session=session,
        position_name="Position B",
        department_id=position.department_id,
        duty_days=position.duty_days,
    )

    shift_name = "Shared shift"

    shift_a = create_new_shift(
        session=session,
        name=shift_name,
        position_id=position.id,
        grants_day_off=False,
        doctors_per_shift=1,
    )

    shift_b = create_new_shift(
        session=session,
        name=shift_name,
        position_id=position_b.id,
        grants_day_off=False,
        doctors_per_shift=1,
    )

    assert shift_a.id != shift_b.id
    assert shift_a.position_id == position.id
    assert shift_b.position_id == position_b.id
    assert shift_a.name == shift_b.name == shift_name


def test_shift_name_cannot_repeat_within_position(
    session,
    shift,
):
    """Tests that a shift name is unique within a position."""
    duplicate_shift = ShiftCreate(
        name=shift.name,
        doctors_per_shift=shift.doctors_per_shift,
        grants_day_off=shift.grants_day_off,
        position_id=shift.position_id,
    )

    with pytest.raises(IntegrityError):
        shift_repository.create_shift(
            session=session,
            shift_data=duplicate_shift,
        )

    session.rollback()


def test_soft_deleted_shift_name_remains_reserved_within_position(
    session,
    shift,
):
    """Tests that soft deletion does not release the shift name."""
    shift.is_deleted = True
    session.add(shift)
    session.commit()
    replacement_shift = ShiftCreate(
        name=shift.name,
        doctors_per_shift=shift.doctors_per_shift,
        grants_day_off=shift.grants_day_off,
        position_id=shift.position_id,
    )

    with pytest.raises(IntegrityError):
        shift_repository.create_shift(
            session=session,
            shift_data=replacement_shift,
        )

    session.rollback()
    session.refresh(shift)

    assert shift.is_deleted is True


def test_get_latest_shift_assignment_date_for_department_returns_max_date(
    session,
    new_doctor,
    shift,
    department,
):
    """Verifies that the repository returns the latest date among active assignments."""
    create_new_shift_assignment(
        session, new_doctor.id, shift.id, datetime.date(2026, 10, 5)
    )
    create_new_shift_assignment(
        session, new_doctor.id, shift.id, datetime.date(2026, 10, 28)
    )
    create_new_shift_assignment(
        session, new_doctor.id, shift.id, datetime.date(2026, 10, 15)
    )

    latest_date = shift_repository.get_latest_shift_assignment_date_for_department(
        session=session,
        department_id=department.id,
    )

    assert latest_date == datetime.date(2026, 10, 28)


def test_get_latest_shift_assignment_date_excludes_deleted_and_foreign_records(
    session,
    department,
    shift,
    new_doctor,
    department_b,
    shift_b,
    doctor_b,
):
    """Verifies that soft-deleted assignments and assignments from another department are excluded."""
    # Active assignment for department A on Oct 10
    create_new_shift_assignment(
        session, new_doctor.id, shift.id, datetime.date(2026, 10, 10)
    )

    # Soft-deleted assignment for department A with a later date (Nov 15)
    deleted_assignment = create_new_shift_assignment(
        session, new_doctor.id, shift.id, datetime.date(2026, 11, 15)
    )
    deleted_assignment.is_deleted = True
    session.add(deleted_assignment)
    session.commit()

    # Foreign assignment for department B on Dec 1
    create_new_shift_assignment(
        session, doctor_b.id, shift_b.id, datetime.date(2026, 12, 1)
    )

    latest_date = shift_repository.get_latest_shift_assignment_date_for_department(
        session=session,
        department_id=department.id,
    )

    # Should resolve to Oct 10, ignoring the soft-deleted Nov 15 and foreign Dec 1
    assert latest_date == datetime.date(2026, 10, 10)


def test_get_latest_shift_assignment_date_empty_department_returns_none(
    session,
    department,
):
    """Verifies that None is returned when a department has no shift assignments."""
    latest_date = shift_repository.get_latest_shift_assignment_date_for_department(
        session=session,
        department_id=department.id,
    )
    assert latest_date is None


def test_soft_delete_shift_assignments_for_department_month(
    session,
    department,
    position,
    shift,
    new_doctor,
    department_b,
    shift_b,
    doctor_b,
):
    """Verifies that soft-delete scoped to department and month only affects target records."""
    # 1. Target assignments in Department A for November 2026
    target_1 = create_new_shift_assignment(
        session, new_doctor.id, shift.id, datetime.date(2026, 11, 2)
    )
    target_2 = create_new_shift_assignment(
        session, new_doctor.id, shift.id, datetime.date(2026, 11, 25)
    )

    # 2. Assignment in Department A for a DIFFERENT month (December 2026)
    other_month = create_new_shift_assignment(
        session, new_doctor.id, shift.id, datetime.date(2026, 12, 5)
    )

    # 3. Assignment in Department B for November 2026 (Foreign Department)
    foreign_dept = create_new_shift_assignment(
        session, doctor_b.id, shift_b.id, datetime.date(2026, 11, 2)
    )

    deleted_count = shift_repository.soft_delete_shift_assignments_for_department_month(
        session=session,
        department_id=department.id,
        year=2026,
        month=11,
    )

    assert deleted_count == 2

    session.refresh(target_1)
    session.refresh(target_2)
    session.refresh(other_month)
    session.refresh(foreign_dept)

    # Target records must be soft-deleted
    assert target_1.is_deleted is True
    assert target_2.is_deleted is True

    # Other month and foreign department records must remain active
    assert other_month.is_deleted is False
    assert foreign_dept.is_deleted is False


def test_bulk_create_shift_assignments(
    session,
    new_doctor,
    shift,
):
    """Verifies that multiple ShiftAssignment rows are created in a single batch."""
    assignments = [
        ShiftAssignmentModel(
            doctor_id=new_doctor.id,
            shift_id=shift.id,
            date=datetime.date(2026, 11, 1),
        ),
        ShiftAssignmentModel(
            doctor_id=new_doctor.id,
            shift_id=shift.id,
            date=datetime.date(2026, 11, 2),
        ),
    ]

    created = shift_repository.bulk_create_shift_assignments(
        session=session,
        assignments=assignments,
    )

    assert len(created) == 2
    assert all(a.id is not None for a in created)
    assert all(a.is_deleted is False for a in created)


def test_soft_delete_shift_assignments_empty_month_returns_zero(
    session,
    department,
):
    """Verifies that soft-deleting a month with zero assignments returns 0."""
    deleted_count = shift_repository.soft_delete_shift_assignments_for_department_month(
        session=session,
        department_id=department.id,
        year=2026,
        month=11,
    )
    assert deleted_count == 0


def test_soft_delete_shift_assignments_ignores_already_deleted(
    session,
    department,
    shift,
    new_doctor,
):
    """Verifies that already soft-deleted assignments are not re-processed or counted."""
    assignment = create_new_shift_assignment(
        session, new_doctor.id, shift.id, datetime.date(2026, 11, 10)
    )
    assignment.is_deleted = True
    session.add(assignment)
    session.commit()

    deleted_count = shift_repository.soft_delete_shift_assignments_for_department_month(
        session=session,
        department_id=department.id,
        year=2026,
        month=11,
    )
    assert deleted_count == 0


def test_bulk_create_shift_assignments_empty_list(session):
    """Verifies that passing an empty list to bulk_create returns an empty list without error."""
    created = shift_repository.bulk_create_shift_assignments(
        session=session,
        assignments=[],
    )
    assert created == []


def test_soft_delete_shift_assignments_commit_false_can_be_rolled_back(
    session,
    department,
    shift,
    new_doctor,
):
    """Verifies that commit=False allows rolling back the soft-delete transaction."""
    assignment = create_new_shift_assignment(
        session, new_doctor.id, shift.id, datetime.date(2026, 11, 10)
    )

    deleted_count = shift_repository.soft_delete_shift_assignments_for_department_month(
        session=session,
        department_id=department.id,
        year=2026,
        month=11,
        commit=False,
    )
    assert deleted_count == 1

    # Roll back the transaction
    session.rollback()
    session.refresh(assignment)

    # Must NOT be deleted because commit=False didn't make it permanent
    assert assignment.is_deleted is False


def test_bulk_create_shift_assignments_commit_false_can_be_rolled_back(
    session,
    new_doctor,
    shift,
):
    """Verifies that commit=False allows rolling back bulk-created assignments."""
    assignment = ShiftAssignmentModel(
        doctor_id=new_doctor.id,
        shift_id=shift.id,
        date=datetime.date(2026, 11, 1),
    )

    created = shift_repository.bulk_create_shift_assignments(
        session=session,
        assignments=[assignment],
        commit=False,
    )
    created_id = created[0].id

    # Roll back the transaction
    session.rollback()

    assert session.get(ShiftAssignmentModel, created_id) is None


def test_stage_shift_flushes_without_committing(session, position):
    """Verify stage_shift creates a shift queryable in-session that rolls back cleanly."""
    shift_data = ShiftCreate(
        name="Staged Shift",
        doctors_per_shift=2,
        grants_day_off=True,
        position_id=position.id,
    )
    staged = shift_repository.stage_shift(
        session=session,
        shift_data=shift_data,
    )

    assert staged.id is not None
    assert staged.name == "Staged Shift"
    assert staged.position_id == position.id

    fetched = session.get(ShiftModel, staged.id)
    assert fetched is not None

    session.rollback()
    assert session.get(ShiftModel, staged.id) is None


#####################
# Controller Tests
#####################


def test_create_shift_controller_duplicate_name(session, department, shift):
    """Tests that creating a shift with a duplicate name returns error and no extra row is created."""
    shift2_data = ShiftCreate(
        name=shift.name,
        position_id=shift.position_id,
        grants_day_off=False,
        doctors_per_shift=1
    )

    with pytest.raises(HTTPException) as exc_info:
        shift_controllers.create_shift_controller(
            shift_data=shift2_data,
            department_id=department.id,
            session=session,
        )

    assert exc_info.value.status_code == 400
    assert "already exists" in exc_info.value.detail

    stored_shifts = session.exec(
        select(ShiftModel).where(ShiftModel.position_id == shift.position_id)
    ).all()
    assert len(stored_shifts) == 1


def test_create_shift_controller_nonexistent_position(session, department):
    """Tests that creating a shift under a non existent position returns 404 Position not found."""
    shift_data = ShiftCreate(
        name="ER 1",
        position_id=uuid.uuid4(),
        grants_day_off=False,
        doctors_per_shift=1
    )

    with pytest.raises(HTTPException) as exc_info:
        shift_controllers.create_shift_controller(
            shift_data=shift_data,
            department_id=department.id,
            session=session,
        )

    assert exc_info.value.status_code == 404
    assert "Position not found" in exc_info.value.detail


def test_create_shift_controller_foreign_position(session, department, position_b):
    """Tests that creating a shift under a Department B position from Department A returns 404 and creates no Shift."""
    shift_data = ShiftCreate(
        name="Foreign Position Shift",
        position_id=position_b.id,
        grants_day_off=False,
        doctors_per_shift=1
    )

    with pytest.raises(HTTPException) as exc_info:
        shift_controllers.create_shift_controller(
            shift_data=shift_data,
            department_id=department.id,
            session=session,
        )

    assert exc_info.value.status_code == 404
    assert "Position not found" in exc_info.value.detail

    stored_shifts = session.exec(
        select(ShiftModel).where(
            ShiftModel.position_id == position_b.id,
            ShiftModel.name == "Foreign Position Shift",
        )
    ).all()
    assert stored_shifts == []


def test_get_shift_controller_nonexistent(session, department):
    """Tests that trying to retrieve a non existent shift returns error"""
    with pytest.raises(HTTPException) as exc_info:
        shift_controllers.get_shift_controller(
            shift_id=uuid.uuid4(),
            department_id=department.id,
            session=session,
        )

    assert exc_info.value.status_code == 404
    assert "not found" in exc_info.value.detail


def test_get_shift_controller_deleted(session, shift, position):
    """Tests that trying to retrieve a deleted shift returns error"""
    shift.is_deleted = True
    session.add(shift)
    session.commit()

    with pytest.raises(HTTPException) as exc_info:
        shift_controllers.get_shift_controller(
            shift_id=shift.id,
            department_id=position.department_id,
            session=session,
        )

    assert exc_info.value.status_code == 404
    assert "not found" in exc_info.value.detail


def test_get_shift_controller_foreign_shift(session, department, shift_b):
    """Tests that a Department A caller cannot retrieve a Department B shift by id."""
    with pytest.raises(HTTPException) as exc_info:
        shift_controllers.get_shift_controller(
            shift_id=shift_b.id,
            department_id=department.id,
            session=session,
        )

    assert exc_info.value.status_code == 404
    assert "not found" in exc_info.value.detail


def test_create_shift_assignment_controller_same_doctor_duplicate(session, shift_assignment, department):
    """Tests that creating a new shift assignment with the same doctor returns error"""
    new_shift_assignment_data = ShiftAssignmentCreate(
        doctor_id=shift_assignment.doctor_id,
        date=shift_assignment.date
    )

    with pytest.raises(HTTPException) as exc_info:
        shift_controllers.create_shift_assignment_controller(
            shift_id=shift_assignment.shift_id,
            department_id=department.id,
            session=session,
            shift_assignment_data=new_shift_assignment_data
        )

    assert exc_info.value.status_code == 400
    assert "Doctor is already assigned on this dat" in exc_info.value.detail


def test_create_shift_assignment_controller_rejects_doctor_assigned_to_another_shift(
    session,
    department,
    position,
    shift,
    new_doctor,
):
    """Tests that a Doctor cannot work two Shifts on the same date."""
    target_date = datetime.date(2026, 8, 12)
    second_shift = create_new_shift(
        session=session,
        name="ER 2",
        position_id=position.id,
        grants_day_off=False,
        doctors_per_shift=2,
    )
    existing_assignment = create_new_shift_assignment(
        session=session,
        doctor_id=new_doctor.id,
        shift_id=shift.id,
        date=target_date,
    )

    assignment_data = ShiftAssignmentCreate(
        doctor_id=new_doctor.id,
        date=target_date,
    )

    with pytest.raises(HTTPException) as exc_info:
        shift_controllers.create_shift_assignment_controller(
            shift_id=second_shift.id,
            department_id=department.id,
            session=session,
            shift_assignment_data=assignment_data,
        )

    assert exc_info.value.status_code == 400
    assert exc_info.value.detail == "Doctor is already assigned on this date."
    assert shift_repository.get_shift_assignments_by_date(
        session=session,
        shift_id=second_shift.id,
        target_date=target_date,
    ) == []

    session.refresh(existing_assignment)
    assert existing_assignment.is_deleted is False


def test_create_shift_assignment_controller_foreign_shift(session, department, shift_b, new_doctor):
    """Tests that creating an assignment against a Department B shift returns 404 and creates no assignment."""
    shift_assignment_data = ShiftAssignmentCreate(
        doctor_id=new_doctor.id,
        date=datetime.date(2026, 8, 12),
    )

    with pytest.raises(HTTPException) as exc_info:
        shift_controllers.create_shift_assignment_controller(
            shift_id=shift_b.id,
            department_id=department.id,
            session=session,
            shift_assignment_data=shift_assignment_data,
        )

    assert exc_info.value.status_code == 404
    assert "Shift or Doctor not found" in exc_info.value.detail
    assert shift_repository.get_shift_assignments_by_date(
        session=session,
        shift_id=shift_b.id,
        target_date=datetime.date(2026, 8, 12),
    ) == []


def test_create_shift_assignment_controller_foreign_doctor(session, department, shift, doctor_b):
    """Tests that creating an assignment for a Department B doctor returns 404 and creates no assignment."""
    shift_assignment_data = ShiftAssignmentCreate(
        doctor_id=doctor_b.id,
        date=datetime.date(2026, 8, 12),
    )

    with pytest.raises(HTTPException) as exc_info:
        shift_controllers.create_shift_assignment_controller(
            shift_id=shift.id,
            department_id=department.id,
            session=session,
            shift_assignment_data=shift_assignment_data,
        )

    assert exc_info.value.status_code == 404
    assert "Doctor not found" in exc_info.value.detail
    assert shift_repository.get_shift_assignments_by_date(
        session=session,
        shift_id=shift.id,
        target_date=datetime.date(2026, 8, 12),
    ) == []


def test_create_shift_assignment_controller_nonexistent_doctor(
    session,
    department,
    shift,
):
    """Tests that a valid but nonexistent Doctor UUID returns 404."""
    nonexistent_doctor_id = uuid.uuid4()
    target_date = datetime.date(2026, 8, 12)
    shift_assignment_data = ShiftAssignmentCreate(
        doctor_id=nonexistent_doctor_id,
        date=target_date,
    )

    with pytest.raises(HTTPException) as exc_info:
        shift_controllers.create_shift_assignment_controller(
            shift_id=shift.id,
            department_id=department.id,
            session=session,
            shift_assignment_data=shift_assignment_data,
        )

    assert exc_info.value.status_code == 404
    assert exc_info.value.detail == "Shift or Doctor not found."
    assert shift_repository.get_shift_assignments_by_date(
        session=session,
        shift_id=shift.id,
        target_date=target_date,
    ) == []


def test_create_shift_assignment_controller_deleted_doctor(
    session,
    department,
    shift,
    new_doctor,
):
    """Tests that a soft-deleted Doctor cannot be assigned."""
    new_doctor.is_deleted = True
    session.add(new_doctor)
    session.commit()

    target_date = datetime.date(2026, 8, 12)
    shift_assignment_data = ShiftAssignmentCreate(
        doctor_id=new_doctor.id,
        date=target_date,
    )

    with pytest.raises(HTTPException) as exc_info:
        shift_controllers.create_shift_assignment_controller(
            shift_id=shift.id,
            department_id=department.id,
            session=session,
            shift_assignment_data=shift_assignment_data,
        )

    assert exc_info.value.status_code == 404
    assert exc_info.value.detail == "Shift or Doctor not found."
    assert shift_repository.get_shift_assignments_by_date(
        session=session,
        shift_id=shift.id,
        target_date=target_date,
    ) == []


def test_create_shift_assignment_controller_deleted_shift(
    session,
    department,
    shift,
    new_doctor,
):
    """Tests that a soft-deleted Shift cannot receive assignments."""
    shift.is_deleted = True
    session.add(shift)
    session.commit()

    target_date = datetime.date(2026, 8, 12)
    shift_assignment_data = ShiftAssignmentCreate(
        doctor_id=new_doctor.id,
        date=target_date,
    )

    with pytest.raises(HTTPException) as exc_info:
        shift_controllers.create_shift_assignment_controller(
            shift_id=shift.id,
            department_id=department.id,
            session=session,
            shift_assignment_data=shift_assignment_data,
        )

    assert exc_info.value.status_code == 404
    assert exc_info.value.detail == "Shift or Doctor not found."
    assert shift_repository.get_shift_assignments_by_date(
        session=session,
        shift_id=shift.id,
        target_date=target_date,
    ) == []


def test_create_shift_assignment_controller_scope_check_precedes_capacity_check(
    session,
    department,
    position_b,
    doctor_b,
    new_doctor,
):
    """Tests that the foreign-shift scope check runs before the capacity business rule."""
    full_foreign_shift = create_new_shift(
        session=session,
        name="Full Foreign Shift",
        position_id=position_b.id,
        grants_day_off=False,
        doctors_per_shift=1,
    )
    create_new_shift_assignment(
        session=session,
        doctor_id=doctor_b.id,
        shift_id=full_foreign_shift.id,
        date=datetime.date(2026, 8, 12),
    )
    shift_assignment_data = ShiftAssignmentCreate(
        doctor_id=new_doctor.id,
        date=datetime.date(2026, 8, 12),
    )

    with pytest.raises(HTTPException) as exc_info:
        shift_controllers.create_shift_assignment_controller(
            shift_id=full_foreign_shift.id,
            department_id=department.id,
            session=session,
            shift_assignment_data=shift_assignment_data,
        )

    assert exc_info.value.status_code == 404
    assert "Shift or Doctor not found" in exc_info.value.detail


# def test_create_shift_assignment_controller_different_doctor_conflict(session, department, team, new_doctor, shift):
#     """Tests that creating a new shift assignment where another doctor is assigned returns error"""

#     new_doctor2 = create_new_doctor(
#         session=session,
#         name="Dr panostest",
#         email="drpanostest@gmail.com",
#         department_id=department.id,
#         team_id=team.id
#     )

#     create_new_shift_assignment(
#         session=session,
#         doctor_id=new_doctor.id,
#         date=datetime.date(2026, 8, 12),
#         shift_id=shift.id,
#     )

#     new_shift_assignment_data = ShiftAssignmentCreate(
#         date=datetime.date(2026, 8, 12),
#         doctor_id=new_doctor2.id
#     )

#     with pytest.raises(HTTPException) as exc_info:
#         shift_controllers.create_shift_assignment_controller(
#             shift_id=shift.id,
#             session=session,
#             shift_assignment_data=new_shift_assignment_data
#         )

#     assert exc_info.value.status_code == 422
#     assert "Another doctor is assigned on this shift" in exc_info.value.detail


def test_create_shift_assignment_unavailability_conflict(session, unavailability, shift, department):
    """Tests that creating a shift assignment when a doctor is unavailable returns error"""

    shift_assignment_data = ShiftAssignmentCreate(
        doctor_id=unavailability.doctor_id,
        date=unavailability.date
    )

    with pytest.raises(HTTPException) as exc_info:
        shift_controllers.create_shift_assignment_controller(
            session=session,
            shift_id=shift.id,
            department_id=department.id,
            shift_assignment_data=shift_assignment_data
        )

    assert exc_info.value.status_code == 400
    assert "Doctor is unavailable" in exc_info.value.detail


def test_create_shift_assignment_capacity_limit(session, shift, new_doctor, department, team):
    """Tests that a shift assignment cannot be created when a shift already has reached the assignment capacity"""
    new_doctor2 = create_new_doctor(
        session, "2nd doc", "2nddoc@gmail.com", department.id, team.id)
    new_doctor3 = create_new_doctor(
        session, "3rd doc", "3rddoc@gmail.com", department.id, team.id)

    target_date = datetime.date(2026, 8, 12)

    create_new_shift_assignment(session, new_doctor.id, shift.id, target_date)
    create_new_shift_assignment(session, new_doctor2.id, shift.id, target_date)

    shift_assignment_data = ShiftAssignmentCreate(
        doctor_id=new_doctor3.id,
        date=target_date
    )

    with pytest.raises(HTTPException) as exc_info:
        shift_controllers.create_shift_assignment_controller(
            session=session,
            shift_id=shift.id,
            department_id=department.id,
            shift_assignment_data=shift_assignment_data
        )

    assert exc_info.value.status_code == 400
    assert "Shift already has max assignments" in exc_info.value.detail


def test_create_shift_assignment_ignores_deleted_assignment_for_capacity(
    session,
    shift,
    new_doctor,
    department,
    team,
):
    """Tests that a deleted assignment does not consume Shift capacity."""
    target_date = datetime.date(2026, 8, 12)
    shift.doctors_per_shift = 1
    session.add(shift)
    session.commit()

    deleted_assignment = create_new_shift_assignment(
        session=session,
        doctor_id=new_doctor.id,
        shift_id=shift.id,
        date=target_date,
    )
    deleted_assignment.is_deleted = True
    session.add(deleted_assignment)
    session.commit()

    replacement_doctor = create_new_doctor(
        session=session,
        name="Replacement Doctor",
        email="replacement@example.com",
        department_id=department.id,
        team_id=team.id,
    )

    created_assignment = shift_controllers.create_shift_assignment_controller(
        session=session,
        shift_id=shift.id,
        department_id=department.id,
        shift_assignment_data=ShiftAssignmentCreate(
            doctor_id=replacement_doctor.id,
            date=target_date,
        ),
    )

    assert created_assignment.doctor_id == replacement_doctor.id
    assert created_assignment.shift_id == shift.id
    assert created_assignment.date == target_date

    active_assignments = shift_repository.get_shift_assignments_by_date(
        session=session,
        shift_id=shift.id,
        target_date=target_date,
    )
    assert [assignment.id for assignment in active_assignments] == [
        created_assignment.id
    ]

    session.refresh(deleted_assignment)
    assert deleted_assignment.is_deleted is True


def test_create_shift_assignment_controller_nonexistent_shift(session, department, new_doctor):
    """Tests that a valid but nonexistent Shift UUID returns 404 instead of raising an AttributeError."""
    nonexistent_shift_id = uuid.uuid4()
    shift_assignment_data = ShiftAssignmentCreate(
        doctor_id=new_doctor.id,
        date=datetime.date(2026, 8, 12),
    )

    with pytest.raises(HTTPException) as exc_info:
        shift_controllers.create_shift_assignment_controller(
            shift_id=nonexistent_shift_id,
            department_id=department.id,
            session=session,
            shift_assignment_data=shift_assignment_data,
        )

    assert exc_info.value.status_code == 404
    assert "Shift or Doctor not found." in exc_info.value.detail
    assert session.exec(
        select(ShiftAssignmentModel).where(
            ShiftAssignmentModel.doctor_id == new_doctor.id
        )
    ).all() == []


def test_create_shift_assignment_controller_scope_check_precedes_unavailability_check(
    session,
    department,
    shift,
    doctor_b,
):
    """Tests that the foreign-doctor scope check runs before the unavailability business rule."""
    target_date = datetime.date(2026, 8, 20)
    create_test_unavailability(session, doctor_b, target_date)

    shift_assignment_data = ShiftAssignmentCreate(
        doctor_id=doctor_b.id,
        date=target_date,
    )

    with pytest.raises(HTTPException) as exc_info:
        shift_controllers.create_shift_assignment_controller(
            shift_id=shift.id,
            department_id=department.id,
            session=session,
            shift_assignment_data=shift_assignment_data,
        )

    assert exc_info.value.status_code == 404
    assert "Shift or Doctor not found." in exc_info.value.detail
    assert shift_repository.get_shift_assignments_by_date(
        session=session,
        shift_id=shift.id,
        target_date=target_date,
    ) == []


def test_create_shift_assignment_controller_scope_check_precedes_duplicate_check(
    session,
    department,
    shift,
    doctor_b,
):
    """Tests that the foreign-doctor scope check runs before the duplicate-assignment business rule."""
    target_date = datetime.date(2026, 8, 21)
    # Seeds a Department B doctor directly onto a Department A shift, bypassing
    # the controller, so the duplicate check would otherwise match on retry.
    preexisting_assignment = create_new_shift_assignment(
        session, doctor_b.id, shift.id, target_date
    )

    shift_assignment_data = ShiftAssignmentCreate(
        doctor_id=doctor_b.id,
        date=target_date,
    )

    with pytest.raises(HTTPException) as exc_info:
        shift_controllers.create_shift_assignment_controller(
            shift_id=shift.id,
            department_id=department.id,
            session=session,
            shift_assignment_data=shift_assignment_data,
        )

    assert exc_info.value.status_code == 404
    assert "Shift or Doctor not found." in exc_info.value.detail
    remaining_assignments = shift_repository.get_shift_assignments_by_date(
        session=session,
        shift_id=shift.id,
        target_date=target_date,
    )
    assert [assignment.id for assignment in remaining_assignments] == [
        preexisting_assignment.id
    ]


def test_create_shift_controller_restores_soft_deleted_shift(
    session,
    department,
    position,
    shift_factory,
):
    """Tests creating a shift with a soft-deleted shift's name restores it and updates settings."""
    deleted_shift = shift_factory(
        name="Archived Night",
        position_id=position.id,
        doctors_per_shift=1,
        grants_day_off=False,
        is_deleted=True,
        sync_status=True,
    )
    previous_updated_at = deleted_shift.updated_at

    restored = shift_controllers.create_shift_controller(
        shift_data=ShiftCreate(
            name="Archived Night",
            position_id=position.id,
            doctors_per_shift=3,
            grants_day_off=True,
        ),
        department_id=department.id,
        session=session,
    )

    assert restored.id == deleted_shift.id
    assert restored.name == "Archived Night"
    assert restored.doctors_per_shift == 3
    assert restored.grants_day_off is True
    assert restored.is_deleted is False
    assert restored.sync_status is False
    assert restored.updated_at >= previous_updated_at


@pytest.mark.parametrize(
    ("update_kwargs", "expected_name", "expected_doctors", "expected_day_off"),
    [
        pytest.param({"name": "Updated Shift"}, "Updated Shift", 2, False, id="name-only"),
        pytest.param({"doctors_per_shift": 4}, "Initial Shift", 4, False, id="doctors-only"),
        pytest.param({"grants_day_off": True}, "Initial Shift", 2, True, id="day-off-only"),
        pytest.param(
            {"name": "Night Shift", "doctors_per_shift": 3, "grants_day_off": True},
            "Night Shift",
            3,
            True,
            id="all-fields",
        ),
    ],
)
def test_update_shift_controller_updates_fields(
    session,
    department,
    shift_factory,
    update_kwargs,
    expected_name,
    expected_doctors,
    expected_day_off,
):
    """Tests partial and full updates via update_shift_controller."""
    target_shift = shift_factory(
        name="Initial Shift",
        doctors_per_shift=2,
        grants_day_off=False,
        sync_status=True,
    )
    previous_updated_at = target_shift.updated_at

    updated = shift_controllers.update_shift_controller(
        shift_id=target_shift.id,
        shift_data=ShiftUpdate(**update_kwargs),
        department_id=department.id,
        session=session,
    )

    assert updated.id == target_shift.id
    assert updated.name == expected_name
    assert updated.doctors_per_shift == expected_doctors
    assert updated.grants_day_off is expected_day_off
    assert updated.sync_status is False
    assert updated.updated_at >= previous_updated_at


@pytest.mark.parametrize("other_is_deleted", [False, True])
def test_update_shift_controller_rejects_colliding_name(
    session,
    department,
    shift_factory,
    other_is_deleted,
):
    """Tests renaming a shift to an existing (active or soft-deleted) name in the same position raises 400."""
    shift_a = shift_factory(name="Shift Alpha")
    shift_b_same_pos = shift_factory(name="Shift Beta", is_deleted=other_is_deleted)

    with pytest.raises(HTTPException) as exc_info:
        shift_controllers.update_shift_controller(
            shift_id=shift_a.id,
            shift_data=ShiftUpdate(name=shift_b_same_pos.name),
            department_id=department.id,
            session=session,
        )

    assert exc_info.value.status_code == 400
    assert "already exists" in exc_info.value.detail


def test_update_shift_controller_allows_same_name_under_another_position(
    session,
    department,
    position,
    shift_factory,
):
    """Tests that shifts in two different positions of the same department may share a name."""
    other_position = position_repository.create_position(
        session=session,
        position_name="ICU",
        department_id=department.id,
        duty_days=[0, 1, 2],
    )
    shift_pos_a = shift_factory(name="Day", position_id=position.id)
    shift_pos_b = shift_factory(name="Evening", position_id=other_position.id)

    updated = shift_controllers.update_shift_controller(
        shift_id=shift_pos_b.id,
        shift_data=ShiftUpdate(name=shift_pos_a.name),
        department_id=department.id,
        session=session,
    )

    assert updated.name == "Day"


def test_update_shift_controller_nonexistent_deleted_or_foreign_raises_404(
    session,
    department,
    shift_factory,
    shift_b,
):
    """Tests updating a missing, soft-deleted, or foreign shift raises 404."""
    deleted_shift = shift_factory(name="Deleted Shift", is_deleted=True)

    for missing_id in (uuid.uuid4(), deleted_shift.id, shift_b.id):
        with pytest.raises(HTTPException) as exc_info:
            shift_controllers.update_shift_controller(
                shift_id=missing_id,
                shift_data=ShiftUpdate(name="New Name"),
                department_id=department.id,
                session=session,
            )

        assert exc_info.value.status_code == 404
        assert exc_info.value.detail == "Shift not found."


def test_delete_shift_controller_soft_deletes_shift_and_upcoming_pre_assignments(
    session,
    department,
    position,
    new_doctor,
    shift_factory,
):
    """Tests soft-deleting a shift soft-deletes upcoming pre-assignments while keeping history."""
    target_shift = shift_factory(name="Target Shift", sync_status=True)
    sibling_shift = shift_factory(name="Sibling Shift")

    today = datetime.date.today()
    upcoming_today_pre = doctor_repository.create_doctor_pre_assignment(
        session=session,
        doctor_id=new_doctor.id,
        pre_assignment_data=DoctorPreAssignmentCreate(
            shift_id=target_shift.id,
            date=today,
        ),
    )
    upcoming_today_pre.sync_status = True
    session.add(upcoming_today_pre)
    session.commit()

    upcoming_future_pre = doctor_repository.create_doctor_pre_assignment(
        session=session,
        doctor_id=new_doctor.id,
        pre_assignment_data=DoctorPreAssignmentCreate(
            shift_id=target_shift.id,
            date=today + datetime.timedelta(days=4),
        ),
    )
    historical_pre = doctor_repository.create_doctor_pre_assignment(
        session=session,
        doctor_id=new_doctor.id,
        pre_assignment_data=DoctorPreAssignmentCreate(
            shift_id=target_shift.id,
            date=today - datetime.timedelta(days=1),
        ),
    )
    sibling_pre = doctor_repository.create_doctor_pre_assignment(
        session=session,
        doctor_id=new_doctor.id,
        pre_assignment_data=DoctorPreAssignmentCreate(
            shift_id=sibling_shift.id,
            date=today + datetime.timedelta(days=6),
        ),
    )

    future_assignment = create_new_shift_assignment(
        session=session,
        doctor_id=new_doctor.id,
        shift_id=target_shift.id,
        date=today + datetime.timedelta(days=2),
    )
    past_assignment = create_new_shift_assignment(
        session=session,
        doctor_id=new_doctor.id,
        shift_id=target_shift.id,
        date=today - datetime.timedelta(days=3),
    )

    shift_controllers.delete_shift_controller(
        shift_id=target_shift.id,
        department_id=department.id,
        session=session,
    )

    for obj in (
        target_shift,
        upcoming_today_pre,
        upcoming_future_pre,
        historical_pre,
        future_assignment,
        past_assignment,
        sibling_shift,
        sibling_pre,
    ):
        session.refresh(obj)

    assert target_shift.is_deleted is True
    assert target_shift.sync_status is False
    assert upcoming_today_pre.is_deleted is True
    assert upcoming_today_pre.sync_status is False
    assert upcoming_future_pre.is_deleted is True

    # Historical pre-assignments, all ShiftAssignments, and sibling entities stay active
    assert historical_pre.is_deleted is False
    assert future_assignment.is_deleted is False
    assert past_assignment.is_deleted is False
    assert sibling_shift.is_deleted is False
    assert sibling_pre.is_deleted is False


def test_delete_shift_controller_nonexistent_deleted_or_foreign_raises_404(
    session,
    department,
    shift_factory,
    shift_b,
):
    """Tests deleting a missing, already soft-deleted, or foreign shift raises 404."""
    deleted_shift = shift_factory(name="Already Gone", is_deleted=True)

    for missing_id in (uuid.uuid4(), deleted_shift.id, shift_b.id):
        with pytest.raises(HTTPException) as exc_info:
            shift_controllers.delete_shift_controller(
                shift_id=missing_id,
                department_id=department.id,
                session=session,
            )

        assert exc_info.value.status_code == 404
        assert exc_info.value.detail == "Shift not found."


#####################
# Route Tests
#####################


def test_create_shift_route(client, position, department_admin_headers):
    """Tests post /api/v1/shifts route"""
    response = client.post(
        "api/v1/shifts",
        json={
            "name": "ER 1",
            "doctors_per_shift": 1,
            "grants_day_off": False,
            "position_id": str(position.id)
        },
        headers=department_admin_headers
    )

    assert response.status_code == 201
    data = response.json()
    assert data["name"] == "ER 1"
    assert data["position_id"] == str(position.id)
    assert "id" in data
    assert "created_at" in data
    assert "updated_at" in data


def test_create_shift_route_invalid_payload(client, position, department_admin_headers):
    """Tests post /api/v1/shifts route with invalid payload"""
    response = client.post(
        "api/v1/shifts",
        json={
            "doctors_per_shift": 1,
            "grants_day_off": False,
            "position_id": str(position.id)
        },
        headers=department_admin_headers
    )

    assert response.status_code == 422


def test_list_shifts_route(client, session, department, department_b, shift, shift_b, viewer_headers):
    """Tests that get /api/v1/shifts only returns shifts scoped to the authenticated department."""
    response = client.get(
        "/api/v1/shifts",
        headers=viewer_headers,
    )

    assert response.status_code == 200
    data = response.json()
    assert isinstance(data, list)

    returned_ids = {item["id"] for item in data}
    assert str(shift.id) in returned_ids
    assert str(shift_b.id) not in returned_ids


def test_create_shift_route_same_name_under_different_position_in_department(
    client,
    session,
    department,
    shift,
    department_admin_headers,
):
    """Tests that a shift name may repeat under a different Position within the same department."""
    other_position = position_repository.create_position(
        session=session,
        position_name="ICU",
        department_id=department.id,
        duty_days=[2, 4, 6],
    )

    response = client.post(
        "api/v1/shifts",
        json={
            "name": shift.name,
            "doctors_per_shift": 1,
            "grants_day_off": False,
            "position_id": str(other_position.id)
        },
        headers=department_admin_headers,
    )

    assert response.status_code == 201
    data = response.json()
    assert data["name"] == shift.name
    assert data["position_id"] == str(other_position.id)


def test_create_shift_route_rejects_foreign_position(
    client,
    session,
    position_b,
    department_admin_headers,
):
    """Tests that a Department A admin cannot create a Shift under a Department B position."""
    shift_name = "Foreign Position Route Shift"

    response = client.post(
        "api/v1/shifts",
        json={
            "name": shift_name,
            "doctors_per_shift": 1,
            "grants_day_off": False,
            "position_id": str(position_b.id)
        },
        headers=department_admin_headers,
    )

    assert response.status_code == 404
    assert response.json() == {"detail": "Position not found."}
    assert response.headers.get("WWW-Authenticate") is None

    stored_shifts = session.exec(
        select(ShiftModel).where(ShiftModel.name == shift_name)
    ).all()
    assert stored_shifts == []


def test_create_shift_route_duplicate_name_within_same_position(
    client,
    session,
    shift,
    department_admin_headers,
):
    """Tests that duplicating a Shift name under the same Position returns 400 and creates no extra row."""
    response = client.post(
        "api/v1/shifts",
        json={
            "name": shift.name,
            "doctors_per_shift": 1,
            "grants_day_off": False,
            "position_id": str(shift.position_id)
        },
        headers=department_admin_headers,
    )

    assert response.status_code == 400
    assert response.json() == {"detail": "Shift already exists"}
    assert response.headers.get("WWW-Authenticate") is None

    stored_shifts = session.exec(
        select(ShiftModel).where(ShiftModel.position_id == shift.position_id)
    ).all()
    assert len(stored_shifts) == 1
    assert stored_shifts[0].id == shift.id


def test_get_shift_route(client, shift, viewer_headers):
    """Tests get /api/v1/shifts/{shift_id} route"""
    response = client.get(
        f"/api/v1/shifts/{shift.id}",
        headers=viewer_headers,
    )

    assert response.status_code == 200
    data = response.json()
    assert data["id"] == str(shift.id)


def test_get_shift_route_malformed_id(client, viewer_headers):
    """Tests get /api/v1/shifts/{shift_id} route with a malformed UUID"""
    response = client.get(
        "/api/v1/shifts/not-a-uuid",
        headers=viewer_headers,
    )

    assert response.status_code == 422


def test_get_shift_route_nonexistent_id(client, viewer_headers):
    """Tests get /api/v1/shifts/{shift_id} route with invalid payload"""
    response = client.get(
        f"/api/v1/shifts/{uuid.uuid4()}",
        headers=viewer_headers,
    )

    assert response.status_code == 404


@pytest.mark.parametrize(
    "role",
    [
        UserRole.DEPARTMENT_ADMIN,
        UserRole.DOCTOR,
        UserRole.VIEWER,
    ],
)
def test_department_member_cannot_get_shift_from_another_department(
    client,
    department,
    shift_b,
    role,
    user_factory,
    auth_headers_factory,
    new_doctor,
):
    """Tests that foreign Shift IDs are hidden from department members."""
    department_user = user_factory(
        role=role,
        department_id=department.id,
        doctor_id=new_doctor.id if role == UserRole.DOCTOR else None
    )

    response = client.get(
        f"/api/v1/shifts/{shift_b.id}",
        headers=auth_headers_factory(department_user),
    )

    assert response.status_code == 404
    assert response.json() == {"detail": "Shift not found."}
    assert response.headers.get("WWW-Authenticate") is None


def test_create_shift_assignment_route(client, shift, new_doctor, department_admin_headers):
    """Tests post /api/v1/shifts/{shift_id}/assignments"""
    response = client.post(
        f"/api/v1/shifts/{shift.id}/assignments",
        json={
            "doctor_id": str(new_doctor.id),
            "date": str(datetime.date(2026, 8, 12))
        },
        headers=department_admin_headers,
    )

    assert response.status_code == 201
    data = response.json()
    assert data["doctor_id"] == str(new_doctor.id)
    assert "id" in data
    assert "created_at" in data
    assert "updated_at" in data


def test_create_shift_assignment_route_invalid_payload(client, shift, new_doctor, department_admin_headers):
    """Tests post /api/v1/shifts/{shift_id}/assignments with invalid payload"""
    response = client.post(
        f"/api/v1/shifts/{shift.id}/assignments",
        json={
            "doctor_id": str(new_doctor.id),
        },
        headers=department_admin_headers,
    )

    assert response.status_code == 422


def test_create_shift_assignment_route_malformed_shift_id(client, new_doctor, department_admin_headers):
    """Tests post /api/v1/shifts/{shift_id}/assignments with a malformed shift id."""
    response = client.post(
        "/api/v1/shifts/not-a-uuid/assignments",
        json={
            "doctor_id": str(new_doctor.id),
            "date": str(datetime.date(2026, 8, 12))
        },
        headers=department_admin_headers,
    )

    assert response.status_code == 422


def test_create_shift_assignment_route_duplicate(client, shift, new_doctor, department_admin_headers):
    """Tests post /api/v1/shifts/{shift_id}/assignments with invalid payload"""
    client.post(
        f"/api/v1/shifts/{shift.id}/assignments",
        json={
            "doctor_id": str(new_doctor.id),
            "date": str(datetime.date(2026, 8, 12))
        },
        headers=department_admin_headers,
    )

    response = client.post(
        f"/api/v1/shifts/{shift.id}/assignments",
        json={
            "doctor_id": str(new_doctor.id),
            "date": str(datetime.date(2026, 8, 12))
        },
        headers=department_admin_headers,
    )

    assert response.status_code == 400


def test_create_shift_assignment_route_mixed_department_shift_and_doctor(
    client,
    session,
    shift,
    doctor_b,
    department_admin_headers,
):
    """Tests that a Department A shift with a Department B doctor returns 404 and creates no assignment."""
    response = client.post(
        f"/api/v1/shifts/{shift.id}/assignments",
        json={
            "doctor_id": str(doctor_b.id),
            "date": str(datetime.date(2026, 8, 12))
        },
        headers=department_admin_headers,
    )

    assert response.status_code == 404
    assert response.json() == {"detail": "Shift or Doctor not found."}
    assert response.headers.get("WWW-Authenticate") is None
    assert shift_repository.get_shift_assignments_by_date(
        session=session,
        shift_id=shift.id,
        target_date=datetime.date(2026, 8, 12),
    ) == []


def test_create_shift_assignment_route_foreign_shift(
    client,
    session,
    shift_b,
    new_doctor,
    department_admin_headers,
):
    """Tests that a Department B shift with a Department A doctor returns 404 and creates no assignment."""
    response = client.post(
        f"/api/v1/shifts/{shift_b.id}/assignments",
        json={
            "doctor_id": str(new_doctor.id),
            "date": str(datetime.date(2026, 8, 12))
        },
        headers=department_admin_headers,
    )

    assert response.status_code == 404
    assert response.json() == {"detail": "Shift or Doctor not found."}
    assert response.headers.get("WWW-Authenticate") is None
    assert shift_repository.get_shift_assignments_by_date(
        session=session,
        shift_id=shift_b.id,
        target_date=datetime.date(2026, 8, 12),
    ) == []


def test_create_shift_assignment_route_nonexistent_shift(
    client,
    session,
    new_doctor,
    department_admin_headers,
):
    """Tests that a valid but nonexistent Shift UUID returns 404 and creates no assignment."""
    nonexistent_shift_id = uuid.uuid4()

    response = client.post(
        f"/api/v1/shifts/{nonexistent_shift_id}/assignments",
        json={
            "doctor_id": str(new_doctor.id),
            "date": str(datetime.date(2026, 8, 12))
        },
        headers=department_admin_headers,
    )

    assert response.status_code == 404
    assert response.json() == {"detail": "Shift or Doctor not found."}
    assert response.headers.get("WWW-Authenticate") is None
    assert session.exec(
        select(ShiftAssignmentModel).where(
            ShiftAssignmentModel.doctor_id == new_doctor.id
        )
    ).all() == []


def test_list_shift_assignments_route(client, shift_assignment, viewer_headers):
    """Tests get /api/v1/shifts/assignments"""
    response = client.get(
        "/api/v1/shifts/assignments",
        headers=viewer_headers
    )

    assert response.status_code == 200
    data = response.json()
    assert isinstance(data, list)

    returned_shift_assignment = next(
        (
            item
            for item in data
            if item["id"] == str(shift_assignment.id)
        ),
        None
    )

    assert returned_shift_assignment is not None
    assert returned_shift_assignment["id"] == str(shift_assignment.id)


@pytest.mark.parametrize(
    "role",
    [
        UserRole.VIEWER,
        UserRole.DOCTOR,
        UserRole.SUPER_ADMIN,
    ]
)
def test_non_department_admin_cannot_create_shift(
    client,
    role,
    position,
    user_factory,
    auth_headers_factory,
    department,
    session,
    new_doctor,
):
    """Tests post /api/v1/shifts route with non department admin headers"""
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
        "/api/v1/shifts",
        json={
            "name": "ER 1",
            "doctors_per_shift": 1,
            "grants_day_off": False,
            "position_id": str(position.id)
        },
        headers=headers,
    )

    assert response.status_code == 403
    assert response.json() == {
        "detail": "Insufficient permissions for this operation."}
    assert response.headers.get("WWW-Authenticate") is None
    assert shift_repository.get_shift_by_name_for_position(
        session=session, shift_name="ER 1", position_id=position.id
    ) is None


def test_create_shift_requires_authentication(client, session, position):
    """Tests post /api/v1/shifts route without auth"""
    response = client.post(
        "/api/v1/shifts",
        json={
            "name": "ER 1",
            "doctors_per_shift": 1,
            "grants_day_off": False,
            "position_id": str(position.id)
        }
    )

    assert response.status_code == 401
    assert response.json() == {"detail": "Unauthorized"}
    assert response.headers.get("WWW-Authenticate") == "Bearer"
    assert shift_repository.get_shift_by_name_for_position(
        session=session, shift_name="ER 1", position_id=position.id) is None


@pytest.mark.parametrize(
    "role",
    [
        UserRole.VIEWER,
        UserRole.DOCTOR,
        UserRole.SUPER_ADMIN,
    ]
)
def test_non_department_admin_cannot_create_shift_assignment(
    client,
    role,
    user_factory,
    auth_headers_factory,
    department,
    new_doctor,
    shift,
    session,
):
    """tests post /api/v1/shifts/{shift_id}/assignments with non department admin headers"""
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
        f"/api/v1/shifts/{shift.id}/assignments",
        json={
            "doctor_id": str(new_doctor.id),
            "date": str(datetime.date(2026, 8, 12))
        },
        headers=headers,
    )

    assert response.status_code == 403
    assert response.json() == {
        "detail": "Insufficient permissions for this operation."}
    assert response.headers.get("WWW-Authenticate") is None
    retrieved_shift_assignments = shift_repository.get_shift_assignments_by_date(
        session=session,
        shift_id=shift.id,
        target_date=datetime.date(2026, 8, 12)
    )
    assert retrieved_shift_assignments == []


def test_create_shift_assignment_requires_authentication(client, shift, new_doctor, session):
    """Tests post /api/v1/shifts/{shift_id}/assignments without auth"""

    response = client.post(
        f"/api/v1/shifts/{shift.id}/assignments",
        json={
            "doctor_id": str(new_doctor.id),
            "date": str(datetime.date(2026, 8, 12))
        },
    )

    assert response.status_code == 401
    assert response.json() == {"detail": "Unauthorized"}
    assert response.headers.get("WWW-Authenticate") == "Bearer"
    retrieved_shift_assignments = shift_repository.get_shift_assignments_by_date(
        session=session,
        shift_id=shift.id,
        target_date=datetime.date(2026, 8, 12)
    )
    assert retrieved_shift_assignments == []


@pytest.mark.parametrize(
    "path_template",
    [
        pytest.param(
            "/api/v1/shifts/",
            id="list-shifts",
        ),
        pytest.param(
            "/api/v1/shifts/{shift_id}",
            id="get-shift",
        ),
        pytest.param(
            "/api/v1/shifts/assignments",
            id="list-shift-assignments",
        ),
    ],
)
def test_shift_read_routes_require_authentication(
    client,
    shift,
    path_template,
):
    """Tests that the shift read routes require auth"""
    path = path_template.format(shift_id=shift.id)

    response = client.get(path)

    assert response.status_code == 401
    assert response.json() == {"detail": "Unauthorized"}
    assert response.headers.get("WWW-Authenticate") == "Bearer"


def test_list_shift_assignments_route_excludes_foreign_department(
    client,
    session,
    shift_assignment,
    shift_b,
    doctor_b,
    viewer_headers,
):
    """Tests that the assignment list excludes assignments from another department."""
    foreign_assignment = create_new_shift_assignment(
        session=session,
        doctor_id=doctor_b.id,
        shift_id=shift_b.id,
        date=datetime.date(2026, 8, 12),
    )

    response = client.get(
        "api/v1/shifts/assignments",
        headers=viewer_headers,
    )

    assert response.status_code == 200
    returned_ids = {item["id"] for item in response.json()}
    assert str(shift_assignment.id) in returned_ids
    assert str(foreign_assignment.id) not in returned_ids


def test_shift_create_schema_rejects_zero_or_negative_doctors_per_shift(position):
    """Verifies that ShiftCreate enforces doctors_per_shift > 0."""
    from pydantic import ValidationError
    from src.shift.schemas import ShiftCreate

    with pytest.raises(ValidationError) as exc_info:
        ShiftCreate(
            name="Invalid Zero",
            position_id=position.id,
            doctors_per_shift=0,
        )
    assert "greater than 0" in str(exc_info.value)

    with pytest.raises(ValidationError) as exc_info:
        ShiftCreate(
            name="Invalid Negative",
            position_id=position.id,
            doctors_per_shift=-2,
        )
    assert "greater than 0" in str(exc_info.value)


def test_shift_update_schema_validates_doctors_per_shift_and_grants_day_off(position):
    """Verifies ShiftUpdate type correctness, gt=0 check, and grants_day_off attribute."""
    from pydantic import ValidationError
    from src.shift.schemas import ShiftUpdate

    # Valid partial update
    update_data = ShiftUpdate(
        name="Updated Name",
        doctors_per_shift=3,
        grants_day_off=True,
    )
    assert update_data.doctors_per_shift == 3
    assert update_data.grants_day_off is True

    # Invalid non-positive doctors_per_shift
    with pytest.raises(ValidationError) as exc_info:
        ShiftUpdate(doctors_per_shift=0)
    assert "greater than 0" in str(exc_info.value)

    with pytest.raises(ValidationError) as exc_info:
        ShiftUpdate(doctors_per_shift=-1)
    assert "greater than 0" in str(exc_info.value)


def test_shift_update_schema_rejects_position_id():
    """Verifies ShiftUpdate forbids client-supplied position_id."""
    with pytest.raises(ValidationError):
        ShiftUpdate(name="Moved Shift", position_id=uuid.uuid4())


def test_restore_soft_deleted_shift_route(
    client,
    session,
    position,
    shift_factory,
    department_admin_headers,
):
    """Tests POST /shifts/ restores a soft-deleted shift and updates its configuration."""
    deleted_shift = shift_factory(
        name="Restorable Night",
        position_id=position.id,
        doctors_per_shift=1,
        grants_day_off=False,
        is_deleted=True,
        sync_status=True,
    )

    response = client.post(
        "/api/v1/shifts/",
        json={
            "name": "Restorable Night",
            "position_id": str(position.id),
            "doctors_per_shift": 3,
            "grants_day_off": True,
        },
        headers=department_admin_headers,
    )

    assert response.status_code == 201
    data = response.json()
    assert data["id"] == str(deleted_shift.id)
    assert data["name"] == "Restorable Night"
    assert data["doctors_per_shift"] == 3
    assert data["grants_day_off"] is True
    assert data["is_deleted"] is False
    assert data["sync_status"] is False

    session.refresh(deleted_shift)
    assert deleted_shift.is_deleted is False
    assert deleted_shift.doctors_per_shift == 3
    assert deleted_shift.grants_day_off is True


@pytest.mark.parametrize(
    ("payload", "expected_name", "expected_doctors", "expected_day_off"),
    [
        pytest.param({"name": "Renamed ER 1"}, "Renamed ER 1", 2, False, id="name-only"),
        pytest.param({"doctors_per_shift": 5}, "ER 1", 5, False, id="doctors-only"),
        pytest.param({"grants_day_off": True}, "ER 1", 2, True, id="day-off-only"),
        pytest.param(
            {"name": "ER 1", "doctors_per_shift": 2, "grants_day_off": False},
            "ER 1",
            2,
            False,
            id="same-values",
        ),
        pytest.param({}, "ER 1", 2, False, id="empty-body"),
    ],
)
def test_update_shift_route(
    client,
    session,
    shift,
    department_admin_headers,
    payload,
    expected_name,
    expected_doctors,
    expected_day_off,
):
    """Tests PATCH /shifts/{shift_id} supports partial, full, and idempotent updates."""
    response = client.patch(
        f"/api/v1/shifts/{shift.id}",
        json=payload,
        headers=department_admin_headers,
    )

    assert response.status_code == 200
    data = response.json()
    assert data["id"] == str(shift.id)
    assert data["name"] == expected_name
    assert data["doctors_per_shift"] == expected_doctors
    assert data["grants_day_off"] is expected_day_off
    assert data["position_id"] == str(shift.position_id)
    assert data["sync_status"] is False

    session.refresh(shift)
    assert shift.name == expected_name
    assert shift.doctors_per_shift == expected_doctors
    assert shift.grants_day_off is expected_day_off


def test_update_shift_route_rejects_supplied_position_id(
    client,
    session,
    department,
    shift,
    department_admin_headers,
):
    """Tests PATCH /shifts/{shift_id} rejects position_id in payload (extra='forbid')."""
    other_position = position_repository.create_position(
        session=session,
        position_name="Other Pos",
        department_id=department.id,
        duty_days=[0, 1],
    )
    original_position_id = shift.position_id

    response = client.patch(
        f"/api/v1/shifts/{shift.id}",
        json={
            "name": "Attempted Move",
            "position_id": str(other_position.id),
        },
        headers=department_admin_headers,
    )

    assert response.status_code == 422
    validation_errors = response.json()["detail"]
    position_id_error = next(
        (
            error for error in validation_errors
            if error["loc"] == ["body", "position_id"]
        ),
        None,
    )
    assert position_id_error is not None
    assert position_id_error["type"] == "extra_forbidden"

    session.refresh(shift)
    assert shift.name == "ER 1"
    assert shift.position_id == original_position_id


@pytest.mark.parametrize("bad_doctors", [0, -1])
def test_update_shift_route_rejects_non_positive_doctors_per_shift(
    client,
    shift,
    department_admin_headers,
    bad_doctors,
):
    """Tests PATCH /shifts/{shift_id} returns 422 when doctors_per_shift <= 0."""
    response = client.patch(
        f"/api/v1/shifts/{shift.id}",
        json={"doctors_per_shift": bad_doctors},
        headers=department_admin_headers,
    )

    assert response.status_code == 422


def test_update_shift_route_rejects_duplicate_name_in_same_position(
    client,
    session,
    shift,
    shift_factory,
    department_admin_headers,
):
    """Tests PATCH /shifts/{shift_id} returns 400 when renaming to another shift in the same position."""
    other_shift = shift_factory(name="ER 2", position_id=shift.position_id)

    response = client.patch(
        f"/api/v1/shifts/{shift.id}",
        json={"name": other_shift.name},
        headers=department_admin_headers,
    )

    assert response.status_code == 400
    assert "already exists" in response.json()["detail"]

    session.refresh(shift)
    assert shift.name == "ER 1"


def test_delete_shift_route(
    client,
    session,
    shift,
    new_doctor,
    department_admin_headers,
):
    """Tests DELETE /shifts/{shift_id} soft-deletes the shift and upcoming pre-assignments."""
    today = datetime.date.today()
    upcoming_pre = doctor_repository.create_doctor_pre_assignment(
        session=session,
        doctor_id=new_doctor.id,
        pre_assignment_data=DoctorPreAssignmentCreate(
            shift_id=shift.id,
            date=today + datetime.timedelta(days=3),
        ),
    )
    assignment = create_new_shift_assignment(
        session=session,
        doctor_id=new_doctor.id,
        shift_id=shift.id,
        date=today + datetime.timedelta(days=4),
    )

    response = client.delete(
        f"/api/v1/shifts/{shift.id}",
        headers=department_admin_headers,
    )

    assert response.status_code == 204
    assert response.content == b""

    for obj in (shift, upcoming_pre, assignment):
        session.refresh(obj)

    assert shift.is_deleted is True
    assert shift.sync_status is False
    assert upcoming_pre.is_deleted is True
    assert assignment.is_deleted is False

    list_response = client.get("/api/v1/shifts/", headers=department_admin_headers)
    assert list_response.status_code == 200
    assert all(item["id"] != str(shift.id) for item in list_response.json())

    get_response = client.get(
        f"/api/v1/shifts/{shift.id}",
        headers=department_admin_headers,
    )
    assert get_response.status_code == 404


@pytest.mark.parametrize("method", ["PATCH", "DELETE"])
@pytest.mark.parametrize("target_kind", ["missing", "deleted", "foreign_department"])
def test_shift_mutation_routes_return_404_for_missing_deleted_or_foreign_shift(
    client,
    shift_factory,
    shift_b,
    department_admin_headers,
    method,
    target_kind,
):
    """Tests PATCH/DELETE return 404 for nonexistent, soft-deleted, or foreign shifts."""
    if target_kind == "missing":
        target_id = uuid.uuid4()
    elif target_kind == "deleted":
        target_id = shift_factory(name="Deleted Target", is_deleted=True).id
    else:
        target_id = shift_b.id

    request_kwargs = {"json": {"name": "New Name"}} if method == "PATCH" else {}
    response = client.request(
        method,
        f"/api/v1/shifts/{target_id}",
        headers=department_admin_headers,
        **request_kwargs,
    )

    assert response.status_code == 404
    assert response.json() == {"detail": "Shift not found."}
    assert response.headers.get("WWW-Authenticate") is None


@pytest.mark.parametrize("method", ["PATCH", "DELETE"])
def test_shift_mutation_routes_require_authentication(client, shift, method):
    """Tests PATCH/DELETE /shifts/{shift_id} require authentication."""
    request_kwargs = {"json": {"name": "No Auth"}} if method == "PATCH" else {}
    response = client.request(method, f"/api/v1/shifts/{shift.id}", **request_kwargs)

    assert response.status_code == 401
    assert response.json() == {"detail": "Unauthorized"}
    assert response.headers.get("WWW-Authenticate") == "Bearer"


@pytest.mark.parametrize("method", ["PATCH", "DELETE"])
@pytest.mark.parametrize(
    "role",
    [
        UserRole.DOCTOR,
        UserRole.VIEWER,
        UserRole.SUPER_ADMIN,
    ],
)
def test_non_department_admin_cannot_update_or_delete_shift(
    client,
    session,
    department,
    shift,
    new_doctor,
    user_factory,
    auth_headers_factory,
    method,
    role,
):
    """Tests that only DEPARTMENT_ADMIN can PATCH or DELETE shifts."""
    user = user_factory(
        role=role,
        department_id=None if role == UserRole.SUPER_ADMIN else department.id,
        doctor_id=new_doctor.id if role == UserRole.DOCTOR else None,
    )
    headers = auth_headers_factory(user)
    request_kwargs = {"json": {"name": "Blocked Rename"}} if method == "PATCH" else {}

    response = client.request(
        method,
        f"/api/v1/shifts/{shift.id}",
        headers=headers,
        **request_kwargs,
    )

    assert response.status_code == 403
    assert response.json() == {"detail": "Insufficient permissions for this operation."}
    assert response.headers.get("WWW-Authenticate") is None

    session.refresh(shift)
    assert shift.name == "ER 1"
    assert shift.is_deleted is False


def test_shift_assignment_read_schema_supports_optional_display_fields(
    shift_assignment,
    position,
):
    """Verifies ShiftAssignmentRead defaults display fields to None and accepts enriched values."""
    from_orm_dto = ShiftAssignmentRead.model_validate(shift_assignment)
    assert from_orm_dto.id == shift_assignment.id
    assert from_orm_dto.doctor_name is None
    assert from_orm_dto.position_id is None
    assert from_orm_dto.position_name is None
    assert from_orm_dto.shift_name is None

    enriched_dto = ShiftAssignmentRead(
        id=shift_assignment.id,
        doctor_id=shift_assignment.doctor_id,
        shift_id=shift_assignment.shift_id,
        date=shift_assignment.date,
        created_at=shift_assignment.created_at,
        updated_at=shift_assignment.updated_at,
        is_deleted=shift_assignment.is_deleted,
        sync_status=shift_assignment.sync_status,
        doctor_name="Dr. House",
        position_id=position.id,
        position_name=position.name,
        shift_name="Night",
    )
    assert enriched_dto.doctor_name == "Dr. House"
    assert enriched_dto.position_id == position.id
    assert enriched_dto.position_name == position.name
    assert enriched_dto.shift_name == "Night"


def test_get_published_shift_assignments_for_department_filters_by_date_and_orders_asc(
    session,
    department,
    position,
    new_doctor,
    shift,
    shift_b,
    doctor_b,
):
    """Tests month-bounded published assignment query, ordering, and joined entities."""
    oct_assignment = create_new_shift_assignment(
        session=session,
        doctor_id=new_doctor.id,
        shift_id=shift.id,
        date=datetime.date(2026, 10, 15),
    )
    nov_late = create_new_shift_assignment(
        session=session,
        doctor_id=new_doctor.id,
        shift_id=shift.id,
        date=datetime.date(2026, 11, 20),
    )
    nov_early = create_new_shift_assignment(
        session=session,
        doctor_id=new_doctor.id,
        shift_id=shift.id,
        date=datetime.date(2026, 11, 5),
    )
    deleted_nov = create_new_shift_assignment(
        session=session,
        doctor_id=new_doctor.id,
        shift_id=shift.id,
        date=datetime.date(2026, 11, 12),
    )
    deleted_nov.is_deleted = True
    session.add(deleted_nov)
    session.commit()

    # Foreign department assignment in the same month must be excluded
    create_new_shift_assignment(
        session=session,
        doctor_id=doctor_b.id,
        shift_id=shift_b.id,
        date=datetime.date(2026, 11, 10),
    )

    nov_rows = shift_repository.get_published_shift_assignments_for_department(
        session=session,
        department_id=department.id,
        start_date=datetime.date(2026, 11, 1),
        end_date=datetime.date(2026, 11, 30),
    )

    assert [row[0].id for row in nov_rows] == [nov_early.id, nov_late.id]
    first_assignment, first_doctor, first_shift, first_position = nov_rows[0]
    assert first_assignment.date == datetime.date(2026, 11, 5)
    assert first_doctor.id == new_doctor.id
    assert first_doctor.name == new_doctor.name
    assert first_shift.id == shift.id
    assert first_shift.name == shift.name
    assert first_position.id == position.id
    assert first_position.name == position.name

    all_rows = shift_repository.get_published_shift_assignments_for_department(
        session=session,
        department_id=department.id,
    )
    assert [row[0].id for row in all_rows] == [
        oct_assignment.id,
        nov_early.id,
        nov_late.id,
    ]


def test_get_published_shift_assignments_preserves_history_when_parents_soft_deleted(
    session,
    department,
    position,
    new_doctor,
    shift,
):
    """Tests active published assignments remain readable after doctor/shift/position soft-delete."""
    assignment = create_new_shift_assignment(
        session=session,
        doctor_id=new_doctor.id,
        shift_id=shift.id,
        date=datetime.date(2026, 11, 8),
    )

    new_doctor.is_deleted = True
    shift.is_deleted = True
    position.is_deleted = True
    session.add(new_doctor)
    session.add(shift)
    session.add(position)
    session.commit()

    rows = shift_repository.get_published_shift_assignments_for_department(
        session=session,
        department_id=department.id,
        start_date=datetime.date(2026, 11, 1),
        end_date=datetime.date(2026, 11, 30),
    )

    assert len(rows) == 1
    ret_assignment, ret_doctor, ret_shift, ret_position = rows[0]
    assert ret_assignment.id == assignment.id
    assert ret_doctor.name == new_doctor.name
    assert ret_shift.name == shift.name
    assert ret_position.name == position.name


def test_get_published_months_for_department_returns_distinct_months_descending(
    session,
    department,
    new_doctor,
    shift,
    shift_b,
    doctor_b,
):
    """Tests distinct YYYY-MM published months sorted newest-first, excluding deleted/foreign rows."""
    create_new_shift_assignment(session, new_doctor.id, shift.id, datetime.date(2026, 9, 4))
    create_new_shift_assignment(session, new_doctor.id, shift.id, datetime.date(2026, 11, 2))
    create_new_shift_assignment(session, new_doctor.id, shift.id, datetime.date(2026, 11, 19))
    create_new_shift_assignment(session, new_doctor.id, shift.id, datetime.date(2026, 10, 11))

    deleted_dec = create_new_shift_assignment(
        session, new_doctor.id, shift.id, datetime.date(2026, 12, 3)
    )
    deleted_dec.is_deleted = True
    session.add(deleted_dec)
    session.commit()

    create_new_shift_assignment(session, doctor_b.id, shift_b.id, datetime.date(2027, 1, 7))

    months = shift_repository.get_published_months_for_department(
        session=session,
        department_id=department.id,
    )

    assert months == ["2026-11", "2026-10", "2026-09"]


def test_list_shift_assignments_controller_filters_by_target_month_and_enriches_display_fields(
    session,
    department,
    position,
    new_doctor,
    shift,
):
    """Tests controller month filtering and joined display field population."""
    oct_assignment = create_new_shift_assignment(
        session=session,
        doctor_id=new_doctor.id,
        shift_id=shift.id,
        date=datetime.date(2026, 10, 25),
    )
    nov_assignment = create_new_shift_assignment(
        session=session,
        doctor_id=new_doctor.id,
        shift_id=shift.id,
        date=datetime.date(2026, 11, 14),
    )

    nov_results = shift_controllers.list_shift_assignments_controller(
        session=session,
        department_id=department.id,
        target_month="2026-11",
    )

    assert len(nov_results) == 1
    assert nov_results[0].id == nov_assignment.id
    assert nov_results[0].doctor_id == new_doctor.id
    assert nov_results[0].doctor_name == new_doctor.name
    assert nov_results[0].shift_id == shift.id
    assert nov_results[0].shift_name == shift.name
    assert nov_results[0].position_id == position.id
    assert nov_results[0].position_name == position.name

    all_results = shift_controllers.list_shift_assignments_controller(
        session=session,
        department_id=department.id,
    )
    assert [item.id for item in all_results] == [oct_assignment.id, nov_assignment.id]


@pytest.mark.parametrize("bad_month", ["2026", "2026-13", "2026-00", "nov-2026", "2026/11"])
def test_list_shift_assignments_route_rejects_invalid_target_month(
    client,
    viewer_headers,
    bad_month,
):
    """Tests GET /api/v1/shifts/assignments rejects malformed target_month with 400."""
    response = client.get(
        f"/api/v1/shifts/assignments?target_month={bad_month}",
        headers=viewer_headers,
    )

    assert response.status_code == 400
    assert "YYYY-MM" in response.json()["detail"]


def test_list_shift_assignments_route_filters_target_month_and_returns_enriched_fields(
    client,
    session,
    department,
    position,
    new_doctor,
    shift,
    user_factory,
    auth_headers_factory,
):
    """Tests GET /api/v1/shifts/assignments?target_month=YYYY-MM for DOCTOR role."""
    doctor_user = user_factory(
        role=UserRole.DOCTOR,
        department_id=department.id,
        doctor_id=new_doctor.id,
    )
    doctor_headers = auth_headers_factory(doctor_user)

    create_new_shift_assignment(
        session=session,
        doctor_id=new_doctor.id,
        shift_id=shift.id,
        date=datetime.date(2026, 10, 29),
    )
    nov_assignment = create_new_shift_assignment(
        session=session,
        doctor_id=new_doctor.id,
        shift_id=shift.id,
        date=datetime.date(2026, 11, 3),
    )

    response = client.get(
        "/api/v1/shifts/assignments?target_month=2026-11",
        headers=doctor_headers,
    )

    assert response.status_code == 200
    payload = response.json()
    assert len(payload) == 1
    assert payload[0]["id"] == str(nov_assignment.id)
    assert payload[0]["doctor_id"] == str(new_doctor.id)
    assert payload[0]["doctor_name"] == new_doctor.name
    assert payload[0]["shift_id"] == str(shift.id)
    assert payload[0]["shift_name"] == shift.name
    assert payload[0]["position_id"] == str(position.id)
    assert payload[0]["position_name"] == position.name
    assert payload[0]["date"] == "2026-11-03"


def test_list_published_months_route_returns_months_descending_and_requires_auth(
    client,
    session,
    new_doctor,
    shift,
    shift_b,
    doctor_b,
    viewer_headers,
):
    """Tests GET /api/v1/shifts/published-months enforces auth and returns distinct YYYY-MM descending."""
    unauth_response = client.get("/api/v1/shifts/published-months")
    assert unauth_response.status_code == 401

    create_new_shift_assignment(session, new_doctor.id, shift.id, datetime.date(2026, 9, 10))
    create_new_shift_assignment(session, new_doctor.id, shift.id, datetime.date(2026, 11, 2))
    create_new_shift_assignment(session, doctor_b.id, shift_b.id, datetime.date(2027, 1, 5))

    response = client.get(
        "/api/v1/shifts/published-months",
        headers=viewer_headers,
    )

    assert response.status_code == 200
    assert response.json() == ["2026-11", "2026-09"]




