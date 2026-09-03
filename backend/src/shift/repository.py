"""
Shift repository functions for handling database operations.
"""

# from typing import Optional
import datetime
import uuid
from sqlmodel import Session, not_, select
from src.position.models import Position as PositionModel
from src.shift.schemas import ShiftCreate, ShiftAssignmentCreate
from src.shift.models import Shift as ShiftModel
from src.shift.models import ShiftAssignment as ShiftAssignmentModel


def create_shift(session: Session, shift_data=ShiftCreate) -> ShiftModel:
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


def get_shift_by_name(session: Session, shift_name: str) -> ShiftModel:
    """Retrieves a shift by its name"""
    statement = select(ShiftModel).where(
        ShiftModel.name == shift_name
    )

    return session.exec(statement).first()


def get_active_shifts(session: Session) -> list[ShiftModel]:
    """Retrieves all active shifts"""
    statement = select(ShiftModel).where(
        not_(ShiftModel.is_deleted)
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


def get_shift_assignments_by_date(session: Session, shift_id: uuid.UUID, target_date: datetime.date) -> list[ShiftAssignmentModel]:
    """Retrieves a shift assignment by its date"""
    statement = select(ShiftAssignmentModel).where(
        ShiftAssignmentModel.date == target_date,
        ShiftAssignmentModel.shift_id == shift_id
    )

    return session.exec(statement).all()


def get_active_shift_assignments(session: Session) -> list[ShiftAssignmentModel]:
    """Retrieves all active shift assignments"""
    statement = select(ShiftAssignmentModel).where(
        not_(ShiftAssignmentModel.is_deleted)
    )

    return list(session.exec(statement).all())
