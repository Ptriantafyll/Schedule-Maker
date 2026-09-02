
"""
Position routes for handling API requests related to position management.
"""

import uuid
from fastapi import APIRouter, Depends, status
from sqlmodel import Session
from src.db.connection import get_session

from src.position.schemas import PositionCreate, PositionRead
from src.position import controllers as position_controllers
from src.auth.dependencies import (
    require_department_admin,
    require_department_member,
    require_department_scope,
)
from src.user.models import User as UserModel

router = APIRouter(
    prefix="/positions",
    tags=["Positions"]
)


@router.post("/", response_model=PositionRead, status_code=status.HTTP_201_CREATED)
def create_position(
    position_data: PositionCreate,
    session: Session = Depends(get_session),
    _current_user: UserModel = Depends(require_department_admin),
    department_id: uuid.UUID = Depends(require_department_scope),
):
    """Endpoint to create a new position"""
    return position_controllers.create_position_controller(
        position_data=position_data,
        department_id=department_id,
        session=session,
    )


@router.get("/", response_model=list[PositionRead])
def list_positions(
    session: Session = Depends(get_session),
    _current_user: UserModel = Depends(require_department_member),
    department_id: uuid.UUID = Depends(require_department_scope),
):
    """Endpoint to list all positions"""
    return position_controllers.list_positions_controller(
        session=session,
        department_id=department_id,
    )


@router.get("/{position_id}", response_model=PositionRead)
def get_positions(
    position_id: uuid.UUID,
    session: Session = Depends(get_session),
    _current_user: UserModel = Depends(require_department_member),
    department_id: uuid.UUID = Depends(require_department_scope),
):
    """Fetches a specific position by its name."""
    return position_controllers.get_position_controller(
        position_id=position_id,
        department_id=department_id,
        session=session,
    )
