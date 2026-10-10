"""
Bootstrap to create privileged users
"""

from sqlmodel import Session

from src.user.schemas import UserAccountCreate
from src.user.models import UserRole
from src.user.models import User as UserModel
from src.department.models import Department as DepartmentModel
from src.department.schemas import DepartmentCreate
from src.department import repository as department_repository
from src.user import services as user_services
from src.utils.logger import log_audit_event
from src.auth.security import validate_password_strength


class SuperAdminAlreadyExistsError(Exception):
    """Raised when a super-admin email already exists"""


class DepartmentAdminAlreadyExistsError(Exception):
    """Raised when a department-admin email already exists"""


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


def create_department_admin(
    session: Session,
    *,
    department_name: str,
    department_code: str = None,
    email: str,
    full_name: str,
    password: str
) -> tuple[DepartmentModel, UserModel]:
    """Creates a depatrment admin user"""
    validate_password_strength(password)
    clean_name = department_name.strip()
    existing_department = department_repository.get_department_by_name_global(
        session=session,
        name=clean_name
    )

    new_department: DepartmentModel = None
    if not existing_department:
        code = (
            department_code.strip()
            if department_code else clean_name[:4]
        ).upper()

        dept_data = DepartmentCreate(
            name=clean_name,
            code=code,
        )

        new_department = department_repository.create_department(
            session=session,
            department_data=dept_data,
            commit=False
        )

    department = existing_department or new_department
    account_data = UserAccountCreate(
        email=email,
        full_name=full_name,
        password=password,
        role=UserRole.DEPARTMENT_ADMIN,
        department_id=department.id,
        doctor_id=None,
    )

    try:
        department_admin_user = user_services.create_user_account(
            session=session,
            account_data=account_data,
        )

        log_audit_event(
            action="department_admin.bootstrap",
            outcome="success",
            message="Department admin created",
            user_id=department_admin_user.id,
            role=department_admin_user.role,
            department_id=department.id
        )

        return (department, department_admin_user)
    except user_services.UserEmailAlreadyExistsError as exc:
        session.rollback()
        raise DepartmentAdminAlreadyExistsError(
            "A user with this email already exists"
        ) from exc
    except Exception:
        session.rollback()
        raise
