"""
Bootstrap to create privileged users
"""

from sqlmodel import Session

from src.user.schemas import UserAccountCreate
from src.user.models import UserRole
from src.user.models import User as UserModel
from src.user import services as user_services


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
    account_data = UserAccountCreate(
        email=email,
        full_name=full_name,
        password=password,
        role=UserRole.SUPER_ADMIN,
        department_id=None,
        doctor_id=None,
    )

    try:
        return user_services.create_user_account(
            session=session,
            account_data=account_data,
        )
    except user_services.UserEmailAlreadyExistsError as exc:
        raise SuperAdminAlreadyExistsError(
            "A user with this email already exists"
        ) from exc
