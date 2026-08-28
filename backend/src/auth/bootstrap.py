"""
Bootstrap to create privileged users
"""

from sqlmodel import Session

from src.user.schemas import UserCreate
from src.user.models import UserRole
from src.user.models import User as UserModel
from src.user import repository as user_repository
from src.auth.security import hash_password


def create_super_admin(
    session: Session,
    *,
    email: str,
    full_name: str,
    password: str,
) -> UserModel:
    """Creates super admin user"""
    hashed_password = hash_password(password)

    super_admin_data = UserCreate(
        email=email,
        full_name=full_name,
        password=hashed_password,
        role=UserRole.SUPER_ADMIN.value,
        department_id=None,
        doctor_id=None,
    )

    super_admin_user = user_repository.create_user(
        session=session,
        user_data=super_admin_data
    )

    return super_admin_user
