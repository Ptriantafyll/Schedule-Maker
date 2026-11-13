"""
Position controller functions for handling business logic related to position management.
"""

import uuid
import datetime
from fastapi import HTTPException, status
from sqlmodel import Session

from src.position import repository
from src.position.schemas import PositionCreate, PositionUpdate
from src.position.models import Position as PositionModel
from src.doctor import repository as doctor_repository
from src.shift import repository as shift_repository


def create_position_controller(
    position_data: PositionCreate,
    department_id: uuid.UUID,
    session: Session,
) -> PositionModel | None:
    """Handled the logic for creating a new position"""
    existing_position = repository.get_position_by_name_for_department(
        session=session,
        position_name=position_data.name,
        department_id=department_id
    )

    if existing_position:
        if not existing_position.is_deleted:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail="Position already exists"
            )

        return repository.restore_position(
            session=session,
            position=existing_position,
            duty_days=position_data.duty_days,
            commit=True
        )

    return repository.create_position(
        session=session,
        position_name=position_data.name,
        duty_days=position_data.duty_days,
        department_id=department_id,
    )


def list_positions_controller(
    session: Session,
    department_id: uuid.UUID,
) -> list[PositionModel]:
    """Handles the logic for listing all active positions"""
    return repository.get_active_positions_for_department(
        session=session,
        department_id=department_id,
    )


def get_position_controller(
    position_id: uuid.UUID,
    department_id: uuid.UUID,
    session: Session,
) -> PositionModel:
    """Handles the logic for retrieving a position by its name"""
    position = repository.get_position_by_id_for_department(
        session=session,
        position_id=position_id,
        department_id=department_id,
    )

    if not position or position.is_deleted:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Position not found"
        )

    return position


def update_position_controller(
    position_id: uuid.UUID,
    position_data: PositionUpdate,
    department_id: uuid.UUID,
    session: Session,
) -> PositionModel:
    """Handles logic for updating a specific position."""
    position_to_update = repository.get_position_by_id_for_department(
        session, position_id, department_id
    )

    if position_to_update is None:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Position not found"
        )

    update_data = position_data.model_dump(exclude_unset=True)
    new_name = update_data.get("name")
    new_duty_days = update_data.get("duty_days")

    name_changed = new_name is not None and new_name != position_to_update.name
    duty_days_changed = (
        new_duty_days is not None
        and new_duty_days != position_to_update.duty_days
    )

    if not name_changed and not duty_days_changed:
        return position_to_update

    if name_changed:
        existing_position = repository.get_position_by_name_for_department(
            session=session,
            position_name=new_name,
            department_id=department_id,
        )

        if existing_position and existing_position.id != position_to_update.id:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail=f"A position named '{new_name}' already exists.",
            )

    return repository.update_position(
        session=session,
        position=position_to_update,
        name=new_name,
        duty_days=new_duty_days,
        commit=True
    )


def delete_position_controller(
    position_id: uuid.UUID,
    department_id: uuid.UUID,
    session: Session,
) -> None:
    """Handles logic for deleting a position by its uuid."""
    position_to_delete = repository.get_position_by_id_for_department(
        session, position_id, department_id
    )

    if position_to_delete is None:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Position not found"
        )

    shift_ids = shift_repository.soft_delete_shifts_for_position(
        session, position_id, department_id, commit=False
    )
    doctor_repository.soft_delete_upcoming_pre_assignments_for_shifts(
        session, shift_ids, department_id, from_date=datetime.date.today(), commit=False
    )
    doctor_repository.soft_delete_doctor_positions_for_position(
        session, position_id, department_id, commit=False
    )
    repository.soft_delete_position(session, position_to_delete, commit=False)

    session.commit()
    return None
