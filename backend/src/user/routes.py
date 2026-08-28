"""
User routes for handling API requests related to user management.
"""

from fastapi import APIRouter, Depends, status
from sqlmodel import Session

from src.db.connection import get_session
from src.user.schemas import UserCreate, UserRead
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
    """Endpoint to lsit all active users"""
    return user_controllers.list_users_controller(session)


@router.post("/signup", response_model=UserRead, status_code=status.HTTP_201_CREATED)
def create_user(user_data: UserCreate, session: Session = Depends(get_session)):
    """Endpoint to create a new user"""
    return user_controllers.create_user_controller(user_data, session)


@router.get("/{user_email}", response_model=UserRead)
def get_user(user_email: str, session: Session = Depends(get_session)):
    """Fetches a specific user by their email"""
    return user_controllers.get_user_controller(user_email, session)
