"""
Shift repository functions for handling database operations.
"""

# from typing import Optional
import datetime
import uuid
import calendar
from sqlmodel import Session, not_, select, func
from src.doctor.models import Doctor as DoctorModel
from src.position.models import Position as PositionModel
from src.shift.schemas import ShiftCreate, ShiftAssignmentCreate
from src.shift.models import Shift as ShiftModel
from src.shift.models import ShiftAssignment as ShiftAssignmentModel


def _sync_metadata(
    session: Session,
    shift: ShiftModel,
    *,
    commit: bool = False
) -> ShiftModel:
    """Syncs metadata and commits or flushes"""
    shift.sync_status = False
    shift.updated_at = datetime.datetime.now(datetime.timezone.utc)

    session.add(shift)
    session.flush()
    if commit:
        session.commit()
        session.refresh(shift)

    return shift


def create_shift(session: Session, shift_data: ShiftCreate) -> ShiftModel:
    """Create a new shift in the database"""
    new_shift = ShiftModel(
        name=shift_data.name,
        doctors_per_shift=shift_data.doctors_per_shift,
        grants_day_off=shift_data.grants_day_off,
        position_id=shift_data.position_id
    )
    session.add(new_shift)
    session.commit()
    session.refresh(new_shift)
    return new_shift


def stage_shift(session: Session, shift_data: ShiftCreate) -> ShiftModel:
    """Creates a new shift but does not commit in database"""
    new_shift = ShiftModel(
        name=shift_data.name,
        doctors_per_shift=shift_data.doctors_per_shift,
        grants_day_off=shift_data.grants_day_off,
        position_id=shift_data.position_id
    )
    session.add(new_shift)
    session.flush()
    return new_shift


def get_shift_by_id_for_department(
    session: Session,
    shift_id: uuid.UUID,
    department_id: uuid.UUID,
) -> ShiftModel | None:
    """Retrieves an active Shift through its Position's department."""
    statement = (
        select(ShiftModel)
        .join(
            PositionModel,
            ShiftModel.position_id == PositionModel.id
        )
        .where(
            ShiftModel.id == shift_id,
            PositionModel.department_id == department_id,
            not_(ShiftModel.is_deleted),
            not_(PositionModel.is_deleted),
        )
    )
    return session.exec(statement).first()


def get_shift_by_name_for_position(
    session: Session,
    shift_name: str,
    position_id: uuid.UUID,
) -> ShiftModel | None:
    """Retrieves a shift by its name"""
    statement = select(ShiftModel).where(
        ShiftModel.name == shift_name,
        ShiftModel.position_id == position_id,
    )

    return session.exec(statement).first()


def get_active_shifts_for_department(
    session: Session,
    department_id: uuid.UUID,
) -> list[ShiftModel]:
    """Retrieves all active shifts"""
    statement = (
        select(ShiftModel)
        .join(
            PositionModel,
            ShiftModel.position_id == PositionModel.id
        )
        .where(
            PositionModel.department_id == department_id,
            not_(ShiftModel.is_deleted),
            not_(PositionModel.is_deleted),
        )
    )

    return list(session.exec(statement).all())


def create_shift_assignment(session: Session, shift_id: uuid.UUID, shift_assignment_data: ShiftAssignmentCreate) -> ShiftAssignmentModel:
    """Creates a shift assignment in the db"""
    new_shift_assignment = ShiftAssignmentModel(
        shift_id=shift_id,
        doctor_id=shift_assignment_data.doctor_id,
        date=shift_assignment_data.date,
    )

    session.add(new_shift_assignment)
    session.commit()
    session.refresh(new_shift_assignment)
    return new_shift_assignment


def get_shift_assignment_by_id(session: Session, shift_assignment_id: uuid.UUID) -> list[ShiftAssignmentModel]:
    """Retrieves a shift assignment by its id"""
    statement = select(ShiftAssignmentModel).where(
        ShiftAssignmentModel.id == shift_assignment_id
    )

    return session.exec(statement).first()


def get_shift_assignments_by_date(
    session: Session,
    shift_id: uuid.UUID,
    target_date: datetime.date
) -> list[ShiftAssignmentModel]:
    """Retrieves a shift assignments for a shift on a date"""
    statement = select(ShiftAssignmentModel).where(
        ShiftAssignmentModel.date == target_date,
        ShiftAssignmentModel.shift_id == shift_id,
        not_(ShiftAssignmentModel.is_deleted),
    )

    return session.exec(statement).all()


def get_active_shift_assignment_for_doctor_on_date(
    session: Session,
    doctor_id: uuid.UUID,
    target_date: datetime.date,
) -> ShiftAssignmentModel | None:
    """Retrieves a shift_assignment for a doctor on a specific date"""
    statement = select(ShiftAssignmentModel).where(
        ShiftAssignmentModel.doctor_id == doctor_id,
        ShiftAssignmentModel.date == target_date,
        not_(ShiftAssignmentModel.is_deleted),
    )

    return session.exec(statement).first()


def get_active_shift_assignments_for_department(
    session: Session,
    department_id: uuid.UUID,
) -> list[ShiftAssignmentModel]:
    """Retrieves all active shift assignments of a department through its shifts and positions."""
    statement = (
        select(ShiftAssignmentModel)
        .join(
            ShiftModel,
            ShiftAssignmentModel.shift_id == ShiftModel.id
        )
        .join(
            PositionModel,
            ShiftModel.position_id == PositionModel.id
        )
        .where(
            PositionModel.department_id == department_id,
            not_(ShiftAssignmentModel.is_deleted),
            not_(ShiftModel.is_deleted),
            not_(PositionModel.is_deleted),
        )
    )

    return list(session.exec(statement).all())


def get_published_shift_assignments_for_department(
    session: Session,
    department_id: uuid.UUID,
    *,
    start_date: datetime.date | None = None,
    end_date: datetime.date | None = None,
) -> list[tuple[ShiftAssignmentModel, DoctorModel, ShiftModel, PositionModel]]:
    """
    Retrieves active published shift assignments joined with Doctor, Shift, and Position
    for a department, optionally bounded by start_date/end_date, ordered by date ascending.
    Preserves historical assignments even if parent Doctor, Shift, or Position was soft-deleted.
    """
    statement = (
        select(ShiftAssignmentModel, DoctorModel, ShiftModel, PositionModel).join(
            DoctorModel, DoctorModel.id == ShiftAssignmentModel.doctor_id
        ).join(
            ShiftModel, ShiftModel.id == ShiftAssignmentModel.shift_id
        ).join(
            PositionModel, PositionModel.id == ShiftModel.position_id
        ).where(
            PositionModel.department_id == department_id,
            not_(ShiftAssignmentModel.is_deleted)
        )
    )

    if start_date:
        statement = statement.where(ShiftAssignmentModel.date >= start_date)

    if end_date:
        statement = statement.where(ShiftAssignmentModel.date <= end_date)

    statement = statement.order_by(ShiftAssignmentModel.date.asc())

    return list(session.exec(statement).all())


def get_published_months_for_department(
    session: Session,
    department_id: uuid.UUID,
) -> list[str]:
    """
    Retrieves distinct YYYY-MM strings for months with active shift assignments
    in a department, ordered descending (newest month first).
    """
    statement = (
        select(ShiftAssignmentModel.date)
        .distinct()
        .join(
            ShiftModel, ShiftModel.id == ShiftAssignmentModel.shift_id
        ).join(
            PositionModel, PositionModel.id == ShiftModel.position_id
        ).where(
            PositionModel.department_id == department_id,
            not_(ShiftAssignmentModel.is_deleted)
        )
    )
    dates = session.exec(statement).all()

    return sorted({f"{d.year:04d}-{d.month:02d}" for d in dates}, reverse=True)


def get_latest_shift_assignment_date_for_department(
    session: Session,
    department_id: uuid.UUID,
) -> datetime.date | None:
    """Retrieves the latest shift assignment date for a department"""
    statement = select(func.max(ShiftAssignmentModel.date)).join(
        ShiftModel,
        ShiftAssignmentModel.shift_id == ShiftModel.id
    ).join(
        PositionModel,
        ShiftModel.position_id == PositionModel.id
    ).where(
        PositionModel.department_id == department_id,
        not_(ShiftAssignmentModel.is_deleted),
        not_(ShiftModel.is_deleted),
        not_(PositionModel.is_deleted),
    )

    return session.exec(statement).first()


def soft_delete_shift_assignments_for_department_month(
    session: Session,
    department_id: uuid.UUID,
    year: int,
    month: int,
    commit: bool = True,
) -> int:
    """
    Sets is_deleted = True for the shift assignments of a department
    For a given month
    """
    start_date = datetime.date(year, month, 1)
    last_day = calendar.monthrange(year, month)[1]
    end_date = datetime.date(year, month, last_day)

    statement = (
        select(ShiftAssignmentModel)
        .join(ShiftModel, ShiftAssignmentModel.shift_id == ShiftModel.id)
        .join(PositionModel, ShiftModel.position_id == PositionModel.id)
        .where(
            PositionModel.department_id == department_id,
            ShiftAssignmentModel.date >= start_date,
            ShiftAssignmentModel.date <= end_date,
            not_(ShiftAssignmentModel.is_deleted),
        )
    )

    assignments = list(session.exec(statement).all())

    for assignment in assignments:
        assignment.is_deleted = True
        session.add(assignment)

    if commit:
        session.commit()
    else:
        session.flush()

    return len(assignments)


def bulk_create_shift_assignments(
    session: Session,
    assignments: list[ShiftAssignmentModel],
    commit: bool = True,
) -> list[ShiftAssignmentModel]:
    """Batch-inserts a list of ShiftAssignmentModel instances into the database."""
    session.add_all(assignments)
    if commit:
        session.commit()
    else:
        session.flush()

    for assignment in assignments:
        session.refresh(assignment)

    return assignments


def soft_delete_shifts_for_position(
    session: Session,
    position_id: uuid.UUID,
    department_id: uuid.UUID,
    *,
    commit: bool = False,
) -> list[uuid.UUID]:
    """Soft deletes shifts of a position given the position UUID."""
    statement = (
        select(ShiftModel)
        .join(PositionModel, ShiftModel.position_id == PositionModel.id)
        .where(
            ShiftModel.position_id == position_id,
            PositionModel.department_id == department_id,
            not_(ShiftModel.is_deleted)
        )
    )

    shifts_to_delete = session.exec(statement).all()
    deleted_uuids: list[uuid.UUID] = []
    for shift in shifts_to_delete:
        shift.is_deleted = True
        shift = _sync_metadata(session, shift, commit=commit)
        deleted_uuids.append(shift.id)

    return deleted_uuids


def restore_shift(
    session: Session,
    shift: ShiftModel,
    doctors_per_shift: int,
    grants_day_off: bool,
    *,
    commit: bool = False,
) -> ShiftModel:
    """Restores a soft-deleted shift and updates its settings."""
    shift.is_deleted = False
    shift.doctors_per_shift = doctors_per_shift
    shift.grants_day_off = grants_day_off
    return _sync_metadata(session, shift, commit=commit)


def update_shift(
    session: Session,
    shift: ShiftModel,
    *,
    name: str | None = None,
    doctors_per_shift: int | None = None,
    grants_day_off: bool | None = None,
    commit: bool = False,
) -> ShiftModel:
    """Updates editable attributes on an active shift."""
    if name is not None:
        shift.name = name

    if doctors_per_shift is not None:
        shift.doctors_per_shift = doctors_per_shift

    if grants_day_off is not None:
        shift.grants_day_off = grants_day_off

    return _sync_metadata(session, shift, commit=commit)


def soft_delete_shift(
    session: Session,
    shift: ShiftModel,
    *,
    commit: bool = False,
) -> ShiftModel:
    """Soft-deletes a shift."""
    shift.is_deleted = True
    return _sync_metadata(session, shift, commit=commit)
