"""
Shift repository functions for handling database operations.
"""

# from typing import Optional
import uuid
from sqlmodel import Session, not_, select

from src.shift.schemas import ShiftCreate
from src.shift.models import Shift as ShiftModel


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


def get_shift_by_id(session: Session, shift_id: uuid.UUID) -> ShiftModel:
    """Retrieves a shift by its id"""
    statement = select(ShiftModel).where(
        ShiftModel.id == shift_id
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
