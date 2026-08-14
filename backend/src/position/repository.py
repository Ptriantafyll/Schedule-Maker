"""
Position repository functions for handling database operations.
"""

# from typing import Optional
import uuid
from sqlmodel import Session, not_, select

from src.position.schemas import PositionCreate
from src.position.models import Position as PositionModel


def create_position(session: Session, position_data: PositionCreate) -> PositionModel:
    """Creates a new position in the database"""
    new_position = PositionModel(
        name="ER",
        department_id=position_data.department_id,
        duty_days=position_data.duty_days,
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