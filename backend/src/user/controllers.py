"""
User controller functions for handling business logic related to user management
"""

import uuid
from fastapi import HTTPException, status
from sqlmodel import Session

from src.user import repository
from src.user.schemas import UserCreate
from src.user.models import User as UserModel
from src.auth.security import hash_password


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


def list_users_controller(session: Session, department_id: uuid.UUID) -> list[UserModel]:
    """Handles the logic for listing users in a department"""
    return repository.get_active_users_by_department(
        session=session,
        department_id=department_id,
    )


# def get_user_controller_global(user_email: str, session: Session) -> UserModel:
#     """Handles the logic for retrieving a user by their email"""
#     user = repository.get_user_by_email(session, user_email)

#     if not user or user.is_deleted:
#         raise HTTPException(
#             status_code=status.HTTP_404_NOT_FOUND,
#             detail="User not found"
#         )

#     return user
