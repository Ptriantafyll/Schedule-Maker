"""
User service module for handling user-related operations, including account creation and persistence.
"""

from sqlalchemy.exc import IntegrityError, SQLAlchemyError
from sqlmodel import Session, select, not_

from src.auth.security import hash_password
from src.user import repository
from src.user.models import User as UserModel, UserRole
from src.user.schemas import UserAccountCreate, UserPersistenceCreate
from src.utils.email import normalize_email
from src.department.models import Department as DepartmentModel
from src.doctor.models import Doctor as DoctorModel


class UserEmailAlreadyExistsError(Exception):
    """Raised when a canonical email is already reserved."""


class InvalidUserAccountRelationshipError(Exception):
    """Raised when a User role has invalid Department or Doctor links."""


class DoctorAlreadyLinkedError(Exception):
    """Raised when a user is created with a link to a doctor that is already linked"""


def validate_user_role_shape(
    account_data: UserAccountCreate,
) -> bool:
    """Validate required and forbidden links for a User role."""
    department_id = account_data.department_id
    doctor_id = account_data.doctor_id

    if account_data.role == UserRole.SUPER_ADMIN:
        is_valid = department_id is None and doctor_id is None
    elif account_data.role == UserRole.DEPARTMENT_ADMIN:
        is_valid = department_id is not None
    elif account_data.role == UserRole.DOCTOR:
        is_valid = department_id is not None and doctor_id is not None
    elif account_data.role == UserRole.VIEWER:
        is_valid = department_id is not None and doctor_id is None
    else:
        is_valid = False

    return is_valid


def validate_user_relationships(account_data: UserAccountCreate, session: Session) -> bool:
    """Validate that the relationships of a user exist and are valid"""
    is_valid = validate_user_role_shape(account_data)
    if not is_valid:
        return False

    if account_data.role != UserRole.SUPER_ADMIN:
        dept = session.get(DepartmentModel, account_data.department_id)
        if not dept or dept.is_deleted:
            return False

    if account_data.doctor_id is not None:
        doctor = session.get(DoctorModel, account_data.doctor_id)
        if not doctor or doctor.is_deleted or (doctor.department_id != dept.id):
            return False

        linked_user = repository.get_user_by_doctor_id(
            session=session,
            doctor_id=account_data.doctor_id,
        )

        if linked_user:
            raise DoctorAlreadyLinkedError(
                "Doctor is already linked to an active user account."
            )

    return True


def stage_user_account(
    session: Session,
    account_data: UserAccountCreate,
) -> UserModel:
    """Stages a prepared User without committing the transaction"""
    is_valid = validate_user_relationships(
        account_data=account_data,
        session=session,
    )

    if not is_valid:
        raise InvalidUserAccountRelationshipError(
            "Invalid relationship between role and account links."
        )

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

        if account_data.doctor_id is not None:
            existing_doctor_user = repository.get_user_by_doctor_id(
                session,
                account_data.doctor_id,
            )

            if existing_doctor_user:
                raise DoctorAlreadyLinkedError(
                    "Doctor is already linked to an active user account."
                ) from exc

        raise

    except SQLAlchemyError:
        session.rollback()
        raise

    session.refresh(staged_user)
    return staged_user
