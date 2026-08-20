
"""
Position routes for handling API requests related to position management.
"""

import uuid
from fastapi import APIRouter, Depends, status
from sqlmodel import Session
from src.db.connection import get_session

from src.position.schemas import PositionCreate, PositionRead
from src.position import controllers as position_controllers

router = APIRouter(
    prefix="/positions",
    tags=["Positions"]
)


@router.post("/", response_model=PositionRead, status_code=status.HTTP_201_CREATED)
def create_position(position_data: PositionCreate, session: Session = Depends(get_session)):
    """Endpoint to create a new position"""
    return position_controllers.create_position_controller(position_data, session)


@router.get("/", response_model=list[PositionRead])
def list_positions(session: Session = Depends(get_session)):
    """Endpoint to list all positions"""
    return position_controllers.list_positions_controller(session)


@router.get("/{position_name}", response_model=PositionRead)
def get_positions(position_name: str, session: Session = Depends(get_session)):
    """Fetches a specific position by its name."""
    return position_controllers.get_position_controller(position_name, session)
