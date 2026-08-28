"""
Authentication routes
"""

from fastapi import APIRouter, Depends
from fastapi.security import OAuth2PasswordRequestForm
from sqlmodel import Session

from src.auth import controllers as auth_controllers
from src.auth.schemas import Token
from src.auth.dependencies import get_current_user
from src.user.models import User as UserModel
from src.db.connection import get_session

router = APIRouter(
    prefix="/auth",
    tags=["Authentication"]
)


@router.post("/me", response_model=Token)
def get_current_user_profile(current_user: UserModel = Depends(get_current_user)):
    """Returns the profile of the currently authenticated user."""
    return current_user


@router.post("/login", response_model=Token)
def login(
    form_data: OAuth2PasswordRequestForm = Depends(),
    session: Session = Depends(get_session),
):
    """Returns the token for the user that logs in"""
    return auth_controllers.login_controller(
        email=form_data.username,
        password=form_data.password,
        session=session,
    )
