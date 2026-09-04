"""
Bootstrap to create privileged users
"""

from sqlmodel import Session
from sqlalchemy.exc import IntegrityError

from src.user.schemas import UserCreate
from src.user.models import UserRole
from src.user.models import User as UserModel
from src.user import repository as user_repository
from src.user import services as user_services
from src.auth.security import hash_password
from src.utils.email import normalize_email


class SuperAdminAlreadyExistsError(Exception):
    """Raised when a super-admin email already exists"""


def create_super_admin(
    session: Session,
    *,
    email: str,
    full_name: str,
    password: str,
) -> UserModel:
    """Creates super admin user"""
    existing_user = user_repository.get_user_by_email(session, email)

    if existing_user:
        raise SuperAdminAlreadyExistsError(
            "A user with this email already exists"
        )

    email = normalize_email(email)
    hashed_password = hash_password(password)

    super_admin_data = UserCreate(
        email=email,
        full_name=full_name,
        password=hashed_password,
        role=UserRole.SUPER_ADMIN,
        department_id=None,
        doctor_id=None,
    )

    try:
        return user_repository.create_user(
            session=session,
            user_data=super_admin_data
        )
    except IntegrityError as exc:
        session.rollback()
        raise SuperAdminAlreadyExistsError(
            "A user with this email already exists"
        ) from exc
