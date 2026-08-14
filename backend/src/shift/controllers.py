"""
Shift controller functions for handling business logic related to team management.
"""

import uuid
from fastapi import HTTPException, status
from sqlmodel import Session
from src.shift import repository
from src.shift.schemas import ShiftCreate, ShiftAssignmentCreate
from src.shift.models import Shift as ShiftModel
from src.shift.models import ShiftAssignment as ShiftAssignmentModel

from src.position import repository as position_repository


def create_shift_controller(shift_data: ShiftCreate, session: Session) -> ShiftModel:
    """Handles logic for creating a shift"""
    existing_shift = repository.get_shift_by_name(session, shift_data.name)
    if existing_shift:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Shift already exists"
        )

    position = position_repository.get_position_by_id(
        session, shift_data.position_id)
    if not position:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Position does not exist"
        )

    return repository.create_shift(session, shift_data)


def list_shifts_controller(session: Session) -> list[ShiftModel]:
    """Handles logic for listing all active shifts"""
    return repository.get_active_shifts(session)


def get_shift_controller(shift_id: uuid.UUID, session: Session) -> ShiftModel:
    """Handles logic for retrieving a shift"""
    shift = repository.get_shift_by_id(session, shift_id)

    if not shift or shift.is_deleted:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Shift not found"
        )

    return shift


def create_shift_assignment_controller(shift_id: uuid.UUID, shift_assignment_data: ShiftAssignmentCreate, session: Session) -> ShiftAssignmentModel:
    """Handles logic for creating a shift assignment"""
    existing_shift_assignment = repository.get_shift_assignment_by_date(
        session=session,
        shift_id=shift_id,
        target_date=shift_assignment_data.date
    )

    if existing_shift_assignment and existing_shift_assignment.doctor_id == shift_assignment_data.doctor_id:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Doctor is already assigned on this date"
        )

    return repository.create_shift_assignment(session, shift_id, shift_assignment_data)


def list_shift_assignments_controller(session: Session) -> ShiftAssignmentModel:
    """Handles logic for retrieving the active shift assignments"""
    return repository.get_active_shift_assignments(session)
