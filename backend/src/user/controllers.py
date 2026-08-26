"""
User controller functions for handling business logic related to user management
"""

from fastapi import HTTPException, status
from sqlmodel import Session

from src.user import repository
from src.user.schemas import UserCreate, UserLogin
from src.user.models import User as UserModel
from src.auth.security import hash_password
from src.auth.schemas import Token
from src.auth.security import verify_password, create_access_token


def create_user_controller(user_data: UserCreate, session: Session) -> UserModel:
    """Handles the logic for creating a new user"""
    existing_user = repository.get_user_by_email(
        session=session,
        user_email=user_data.email
    )

    if existing_user:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="User already exists"
        )

    user_data.password = hash_password(user_data.password)

    return repository.create_user(session, user_data)


def list_users_controller(session: Session) -> list[UserModel]:
    """Handles the logic for listing all active users"""
    return repository.get_active_users(session)


def get_user_controller(user_email: str, session: Session) -> UserModel:
    """Handles the logic for retrieving a user by their email"""
    user = repository.get_user_by_email(session, user_email)

    if not user or user.is_deleted:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="User not found"
        )

    return user


def login_controller(login_data: UserLogin, session: Session) -> Token:
    """Verifies credentials and issues a JWT access token."""
    user = repository.get_user_by_email(session, login_data.email)

    if not user or user.is_deleted or not verify_password(login_data.password, user.hashed_password):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Username or password is incorrect",
            headers={"WWW-Authenticate": "Bearer"}
        )

    payload = {
        "sub": str(user.id),
        "email": user.email,
        "role": user.role.value if hasattr(user.role, "value") else user.role,
        "department_id": str(user.department_id) if user.department_id else None,
        "doctor__id": str(user.doctor_id) if user.doctor_id else None
    }

    access_token = create_access_token(data=payload)
    return Token(access_token=access_token, token_type="bearer")
