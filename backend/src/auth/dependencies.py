"""
Security utils for authentication
"""
import uuid
import jwt
from fastapi import Depends, HTTPException, status, Request
from fastapi.security import OAuth2PasswordBearer
from sqlmodel import Session

from src.auth.security import decode_access_token
from src.db.connection import get_session
from src.user.models import User as UserModel
from src.user.models import UserRole
from src.user import repository as user_repository

oauth2_scheme = OAuth2PasswordBearer(
    tokenUrl="/api/v1/auth/login", auto_error=False
)


def get_current_user(
        token: str | None = Depends(oauth2_scheme),
        session: Session = Depends(get_session),
) -> UserModel:
    """Extract authenticated user from request state or raise 401"""
    if token is None:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Unauthorized",
            headers={"WWW-Authenticate": "Bearer"}
        )

    try:
        payload = decode_access_token(token)
        user_id = uuid.UUID(payload["sub"])
    except (ValueError, TypeError, jwt.InvalidTokenError) as exc:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Unauthorized",
            headers={"WWW-Authenticate": "Bearer"}
        ) from exc

    user = user_repository.get_user_by_id(session, user_id)
    if not user or user.is_deleted:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="User account no longer active",
            headers={"WWW-Authenticate": "Bearer"}
        )

    return user


def require_role(*allowed_roles: UserRole):
    """Factory dependenct for role enforcement"""
    def role_guard(current_user: UserModel = Depends(get_current_user)) -> UserModel:
        if current_user.role not in allowed_roles:
            raise HTTPException(
                status_code=status.HTTP_403_FORBIDDEN,
                detail="Insufficient permissions for this operation"
            )

        return current_user
    return role_guard


require_super_admin = require_role(UserRole.SUPER_ADMIN)
require_department_admin = require_role(UserRole.DEPARTMENT_ADMIN)
require_doctor_or_department_admin = require_role(
    UserRole.DEPARTMENT_ADMIN, UserRole.DOCTOR)
