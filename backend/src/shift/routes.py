"""
Shift routes for handling API requests related to shift management.
"""

import uuid
from fastapi import APIRouter, Depends, status
from sqlmodel import Session
from src.db.connection import get_session

from src.shift.schemas import ShiftCreate, ShiftRead
from src.shift.controllers import create_shift_controller, list_shifts_controller, get_shift_controller 

router = APIRouter(
    prefix="/shifts",
    tags=["Shifts"]
)


@router.post("/", response_model=ShiftRead, status_code=status.HTTP_201_CREATED)
def create_shift(shift_data: ShiftCreate, session: Session = Depends(get_session)):
    """Endpoint to create a new shift"""
    return create_shift_controller(shift_data, session)


@router.get("/", response_model=list[ShiftRead])
def list_shifts(session: Session = Depends(get_session)):
    """Endpoint to list all shifts"""
    return list_shifts_controller(session)


@router.get("/{shift_id}", response_model=ShiftRead)
def get_shift(shift_id: uuid.UUID, session: Session = Depends(get_session)):
    """Fetches a specific shift by its UUID."""
    return get_shift_controller(shift_id, session)