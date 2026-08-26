"""
Authentication routes
"""

from fastapi import APIRouter, Depends

from src.auth.schemas import Token
from src.auth.dependencies import get_current_user
from src.user.models import User as UserModel

router = APIRouter(
    prefix="/auth",
    tags=["Authentication"]
)


@router.post("/me", response_model=Token)
def get_current_user_profile(current_user: UserModel = Depends(get_current_user)):
    """Returns the profile of the currently authenticated user."""
    return current_user
