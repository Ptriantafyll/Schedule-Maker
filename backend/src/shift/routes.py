"""
Shift routes for handling API requests related to shift management.
"""

import uuid
from fastapi import APIRouter, Depends, status
from sqlmodel import Session
from src.db.connection import get_session

from src.shift.schemas import ShiftCreate, ShiftRead, ShiftAssignmentCreate, ShiftAssignmentRead
from src.shift import controllers as shift_controllers
from src.auth.dependencies import (
    require_department_member,
    require_department_admin,
    require_department_scope,
)

from src.user.models import User as UserModel

router = APIRouter(
    prefix="/shifts",
    tags=["Shifts"]
)


@router.post("/", response_model=ShiftRead, status_code=status.HTTP_201_CREATED)
def create_shift(
    shift_data: ShiftCreate,
    session: Session = Depends(get_session),
    _current_user: UserModel = Depends(require_department_admin),
    department_id: uuid.UUID = Depends(require_department_scope),
):
    """Endpoint to create a new shift"""
    return shift_controllers.create_shift_controller(
        department_id=department_id,
        shift_data=shift_data,
        session=session,
    )


@router.get("/", response_model=list[ShiftRead])
def list_shifts(
    session: Session = Depends(get_session),
    _current_user: UserModel = Depends(require_department_member),
    department_id: uuid.UUID = Depends(require_department_scope),
):
    """Endpoint to list all shifts"""
    return shift_controllers.list_shifts_controller(session, department_id=department_id)


@router.get("/assignments", response_model=list[ShiftAssignmentRead])
def list_shift_assignments(
    session: Session = Depends(get_session),
    _current_user: UserModel = Depends(require_department_member),
    department_id: uuid.UUID = Depends(require_department_scope),
):
    """Retrieves all shift assignments of a department."""
    return shift_controllers.list_shift_assignments_controller(
        session=session,
        department_id=department_id,
    )


@router.get("/{shift_id}", response_model=ShiftRead)
def get_shift(
    shift_id: uuid.UUID,
    session: Session = Depends(get_session),
    _current_user: UserModel = Depends(require_department_member),
    department_id: uuid.UUID = Depends(require_department_scope),
):
    """Fetches a specific shift by its ID."""
    return shift_controllers.get_shift_controller(
        shift_id=shift_id,
        department_id=department_id,
        session=session,
    )


@router.post("/{shift_id}/assignments", response_model=ShiftAssignmentRead, status_code=status.HTTP_201_CREATED)
def create_shift_assignment(
    shift_id: uuid.UUID,
    shift_assignment_data: ShiftAssignmentCreate,
    session: Session = Depends(get_session),
    _current_user: UserModel = Depends(require_department_admin),
    department_id: uuid.UUID = Depends(require_department_scope),
):
    """Endpoint to create a new shift"""
    return shift_controllers.create_shift_assignment_controller(
        session=session,
        shift_assignment_data=shift_assignment_data,
        shift_id=shift_id,
        department_id=department_id
    )
