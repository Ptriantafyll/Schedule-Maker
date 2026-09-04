"""
User service module for handling user-related operations, including account creation and persistence.
"""

from sqlalchemy.exc import IntegrityError, SQLAlchemyError
from sqlmodel import Session

from src.auth.security import hash_password
from src.user import repository
from src.user.models import User as UserModel
from src.user.schemas import UserAccountCreate, UserPersistenceCreate
from src.utils.email import normalize_email


class UserEmailAlreadyExistsError(Exception):
    """Raised when a canonical email is already reserved."""


def stage_user_account(
    session: Session,
    account_data: UserAccountCreate,
) -> UserModel:
    """Stages a prepared User without committing the transaction"""
    existing_user = repository.get_user_by_email(
        session=session,
        user_email=account_data.email
    )
    if existing_user:
        raise UserEmailAlreadyExistsError(
            "A user with this email already exists."
        )

    normalized_email = normalize_email(account_data.email)
    hashed_password = hash_password(account_data.password)

    user_persistence_data = UserPersistenceCreate(
        email=normalized_email,
        full_name=account_data.full_name,
        role=account_data.role,
        hashed_password=hashed_password,
        doctor_id=account_data.doctor_id,
        department_id=account_data.department_id
    )
    return repository.add_user(
        session=session,
        user_data=user_persistence_data,
    )


def create_user_account(
    session: Session,
    account_data: UserAccountCreate,
) -> UserModel:
    """Creates a new user account and commits the transaction"""
    try:
        staged_user = stage_user_account(
            session=session,
            account_data=account_data,
        )
        session.commit()
    except IntegrityError as exc:
        session.rollback()

        existing_user = repository.get_user_by_email(
            session=session,
            user_email=account_data.email,
        )
        if existing_user:
            raise UserEmailAlreadyExistsError(
                "A user with this email already exists."
            ) from exc
        raise 
    except SQLAlchemyError:
        session.rollback()
        raise

    session.refresh(staged_user)
    return staged_user
