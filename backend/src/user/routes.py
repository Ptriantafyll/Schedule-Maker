"""
User routes for handling API requests related to user management.
"""

from fastapi import APIRouter, Depends, HTTPException, status
from sqlmodel import Session

from src.db.connection import get_session
from src.user.schemas import UserRead
from src.user.models import User as UserModel
from src.user import controllers as user_controllers

from src.auth.dependencies import require_department_admin

router = APIRouter(
    prefix="/users",
    tags=["Users"]
)


@router.get("/", response_model=list[UserRead])
def list_users(
    session: Session = Depends(get_session),
    current_user: UserModel = Depends(require_department_admin),
):
    """Endpoint to list active users in the authenticated user's department"""
    if current_user.department_id is None:
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="Invalid account scope.",
        )

    return user_controllers.list_users_controller(session, current_user.department_id)
