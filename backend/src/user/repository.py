
"""
Authentication repository function for handling database operations.
"""

import uuid
from sqlmodel import Session, not_, select

from src.user.schemas import UserCreate
from src.user.models import User as UserModel


def create_user(session: Session, user_data: UserCreate) -> UserModel:
    """Creates a new user in the database"""
    new_user = UserModel(
        email=user_data.email,
        full_name=user_data.full_name,
        role=user_data.role,
        hashed_password=user_data.password,
        doctor_id=user_data.doctor_id,
        department_id=user_data.department_id
    )

    session.add(new_user)
    session.commit()
    session.refresh(new_user)
    return new_user


def get_user_by_id(session: Session, user_id: uuid.UUID) -> UserModel:
    """Retrieves a user by their id"""
    statement = select(UserModel).where(
        UserModel.id == user_id
    )
    return session.exec(statement).first()


def get_user_by_email(session: Session, user_email: str) -> UserModel:
    """Retrieves a user by their email"""
    statement = select(UserModel).where(
        UserModel.email == user_email
    )
    return session.exec(statement).first()


def get_active_users(session: Session) -> list[UserModel]:
    """Retrieves all active (non deleted) users"""
    statement = select(UserModel).where(
        not_(UserModel.is_deleted)
    )

    return session.exec(statement).all()


def get_active_users_by_department(
    session: Session,
    department_id: uuid.UUID,
) -> list[UserModel]:
    """Retrieved all active users of a department"""
    statement = select(UserModel).where(
        UserModel.department_id == department_id,
        not_(UserModel.is_deleted)
    )

    return session.exec(statement).all()
