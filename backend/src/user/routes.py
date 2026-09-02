"""
User routes for handling API requests related to user management.
"""

import uuid
from fastapi import APIRouter, Depends
from sqlmodel import Session

from src.db.connection import get_session
from src.user.schemas import UserRead
from src.user.models import User as UserModel
from src.user import controllers as user_controllers

from src.auth.dependencies import require_department_admin, require_department_scope

router = APIRouter(
    prefix="/users",
    tags=["Users"]
)


@router.get("/", response_model=list[UserRead])
def list_users(
    session: Session = Depends(get_session),
    _current_user: UserModel = Depends(require_department_admin),
    department_id: uuid.UUID = Depends(require_department_scope),
):
    """Endpoint to list active users in the authenticated user's department"""
    return user_controllers.list_users_controller(
        session=session,
        department_id=department_id,
    )
