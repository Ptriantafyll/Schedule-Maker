"""
Shift repository functions for handling database operations.
"""

# from typing import Optional
import datetime
import uuid
import calendar
from sqlmodel import Session, not_, select, func
from src.position.models import Position as PositionModel
from src.shift.schemas import ShiftCreate, ShiftAssignmentCreate
from src.shift.models import Shift as ShiftModel
from src.shift.models import ShiftAssignment as ShiftAssignmentModel


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
