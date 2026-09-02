"""
Position controller functions for handling business logic related to position management.
"""

import uuid
from fastapi import HTTPException, status
from sqlmodel import Session

from src.position import repository
from src.position.schemas import PositionCreate
from src.position.models import Position as PositionModel
from src.department import repository as department_repository


def create_position_controller(
    position_data: PositionCreate,
    department_id: uuid.UUID,
    session: Session,
) -> PositionModel:
    """Handled the logic for creating a new position"""
    existing_position = repository.get_position_by_name(
        session=session,
        position_name=position_data.name
    )

    if existing_position:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Position already exists"
        )

    department = department_repository.get_department_by_id(
        session, position_data.department_id)
    if not department:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_CONTENT,
            detail="Department does not exist"
        )

    return repository.create_position(
        session=session,
        position_name=position_data.name,
        duty_days=position_data.duty_days,
        department_id=department_id,
    )


def list_positions_controller(session: Session) -> list[PositionModel]:
    """Handles the logic for listing all active positions"""
    return repository.get_active_positions(session)


def get_position_controller(position_name: str, session: Session) -> PositionModel:
    """Handles the logic for retrieving a position by its name"""
    position = repository.get_position_by_name(session, position_name)

    if not position or position.is_deleted:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Position not found"
        )

    return position
