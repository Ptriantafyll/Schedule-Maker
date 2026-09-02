"""
Position repository functions for handling database operations.
"""

# from typing import Optional
import uuid
from sqlmodel import Session, not_, select

from src.position.models import Position as PositionModel


def create_position(
    session: Session,
    position_name: str,
    duty_days: list[int],
    department_id: uuid.UUID,
) -> PositionModel:
    """Creates a new position in the database"""
    new_position = PositionModel(
        name=position_name,
        department_id=department_id,
        duty_days=duty_days,
    )

    session.add(new_position)
    session.commit()
    session.refresh(new_position)
    return new_position


def get_position_by_id(session: Session, position_id: uuid.UUID) -> PositionModel:
    """Retrieves a position by its id"""
    statement = select(PositionModel).where(
        PositionModel.id == position_id
    )
    return session.exec(statement).first()


def get_position_by_name(session: Session, position_name: str) -> PositionModel:
    """Retrieves a position by its id"""
    statement = select(PositionModel).where(
        PositionModel.name == position_name
    )

    return session.exec(statement).first()


def get_active_positions(session: Session) -> list[PositionModel]:
    """Retrieves all active (non deleted) positions"""
    statement = select(PositionModel).where(
        not_(PositionModel.is_deleted)
    )

    return session.exec(statement).all()
