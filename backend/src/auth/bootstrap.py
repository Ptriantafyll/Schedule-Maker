"""
Bootstrap to create privileged users
"""

from sqlmodel import Session

from src.user.schemas import UserAccountCreate
from src.user.models import UserRole
from src.user.models import User as UserModel
from src.user import services as user_services
from src.utils.logger import log_audit_event
from src.auth.security import validate_password_strength


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
    validate_password_strength(password)
    account_data = UserAccountCreate(
        email=email,
        full_name=full_name,
        password=password,
        role=UserRole.SUPER_ADMIN,
        department_id=None,
        doctor_id=None,
    )

    try:
        super_admin_user = user_services.create_user_account(
            session=session,
            account_data=account_data,
        )

        log_audit_event(
            action="super_admin.bootstrap",
            outcome="success",
            message="Super admin created",
            user_id=super_admin_user.id,
            role=super_admin_user.role,
        )

        return super_admin_user
    except user_services.UserEmailAlreadyExistsError as exc:
        raise SuperAdminAlreadyExistsError(
            "A user with this email already exists"
        ) from exc
