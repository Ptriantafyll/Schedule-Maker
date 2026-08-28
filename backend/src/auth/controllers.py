"""
Authentication controller functions handling business logic
"""

from fastapi import HTTPException, status
from sqlmodel import Session
from src.user import repository as user_repository
from src.auth.schemas import Token
from src.auth.security import verify_password, create_access_token


def login_controller(email: str, password: str, session: Session) -> Token:
    """Handles logic for logging in"""
    user = user_repository.get_user_by_email(session, email)

    if not user or user.is_deleted or not verify_password(password, user.hashed_password):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Username or password is incorrect",
            headers={"WWW-Authenticate": "Bearer"}
        )

    access_token = create_access_token({"sub": str(user.id)})

    return Token(
        access_token=access_token,
        token_type="bearer"
    )
