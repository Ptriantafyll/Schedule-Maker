"""
Authentication controller functions handling business logic
"""

import uuid
from typing import Optional
from fastapi import HTTPException, status, Response, Request
from sqlmodel import Session
from src.user import repository as user_repository
from src.auth import repository as auth_repository
from src.auth.schemas import (
    Token,
    StaffInvitationCreate,
    DepartmentAdminProvisioningCreate,
    DepartmentAdminInvitationCreate,
    DepartmentProvisioningResponse,
    InvitationCreatedResponse,
    InvitationSignupRequest,
    RefreshTokenRequest,
)
from src.auth.security import (
    verify_password,
    create_access_token,
    hash_refresh_token,
    hash_csrf_token,
    SECURE_COOKIE,
)
from src.user.models import UserRole, User as UserModel
from src.auth.models import Invitation as InvitationModel
from src.auth.services import (
    create_staff_invitation,
    provision_department_with_admin,
    create_department_admin_invitation,
    consume_invitation_and_signup,
    issue_refresh_session,
    rotate_refresh_token,
    terminate_refresh_session,
    UnauthorizedInvitationActionError,
    InvalidInvitationDoctorError,
    DoctorAlreadyLinkedError,
    DoctorInvitationAlreadyPendingError,
    DepartmentAlreadyExistsError,
    DepartmentNotFoundError,
    InvitationExpiredError,
    InvitationRevokedError,
    InvitationNotFoundError,
    InvitationAlreadyUsedError,
    InvalidRefreshTokenError,
    RefreshTokenReuseDetectedError,
    RevokedRefreshTokenError,
    ExpiredRefreshTokenError,

)
from src.user.services import (
    UserEmailAlreadyExistsError,
    InvalidUserAccountRelationshipError,
)
from src.department.schemas import DepartmentRead


def login_controller(
    email: str,
    password: str,
    session: Session,
    response: Optional[Response] = None,
) -> Token:
    """Handles logic for logging in"""
    user = user_repository.get_user_by_email(session, email)

    if not user or user.is_deleted or not verify_password(password, user.hashed_password):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Username or password is incorrect",
            headers={"WWW-Authenticate": "Bearer"}
        )

    access_token = create_access_token({"sub": str(user.id)})

    _, raw_refresh, raw_csrf = issue_refresh_session(
        session=session, user_id=user.id)

    if response is not None:
        response.set_cookie(
            key="refresh_token",
            value=raw_refresh,
            httponly=True,
            secure=SECURE_COOKIE,
            samesite="lax",
            max_age=14 * 24 * 3600,
        )

        response.set_cookie(
            key="csrf_token",
            value=raw_csrf,
            httponly=False,  # JavaScript must be able to read this cookie
            secure=SECURE_COOKIE,
            samesite="lax",
            max_age=14 * 24 * 3600,
        )

    return Token(
        access_token=access_token,
        token_type="bearer",
        refresh_token=raw_refresh,
        csrf_token=raw_csrf,
    )


def signup_controller(session: Session, data: InvitationSignupRequest) -> UserModel:
    """Handles logic for user signup"""
    try:
        return consume_invitation_and_signup(
            session=session,
            data=data,
        )
    except InvitationNotFoundError as exc:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail=str(exc)
        ) from exc
    except UserEmailAlreadyExistsError as exc:
        raise HTTPException(
            status_code=status.HTTP_409_CONFLICT,
            detail=str(exc)
        ) from exc
    except (
        InvitationAlreadyUsedError,
        InvitationRevokedError,
        InvitationExpiredError,
        DoctorAlreadyLinkedError,
        InvalidUserAccountRelationshipError,
    ) as exc:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=str(exc)
        ) from exc


def create_staff_invitation_controller(
    current_user: UserModel,
    data: StaffInvitationCreate,
    session: Session
) -> InvitationCreatedResponse:
    """Handles logic for creating a staff invitation"""
    try:
        return create_staff_invitation(
            session=session,
            current_user=current_user,
            data=data
        )
    except (UnauthorizedInvitationActionError) as exc:
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail=str(exc)
        ) from exc
    except (InvalidInvitationDoctorError) as exc:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=str(exc)
        ) from exc
    except (DoctorAlreadyLinkedError,  DoctorInvitationAlreadyPendingError) as exc:
        raise HTTPException(
            status_code=status.HTTP_409_CONFLICT,
            detail=str(exc)
        ) from exc


def provision_department_controller(
    session: Session,
    current_user: UserModel,
    data: DepartmentAdminProvisioningCreate,
) -> DepartmentProvisioningResponse:
    """Handles logic for provitioning a department and creating the dept admins invitation token."""
    try:
        dept, inv = provision_department_with_admin(
            session=session,
            current_user=current_user,
            data=data,
        )
    except (UnauthorizedInvitationActionError) as exc:
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail=str(exc)
        ) from exc
    except (DepartmentAlreadyExistsError) as exc:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=str(exc)
        ) from exc

    return DepartmentProvisioningResponse(
        department=DepartmentRead.model_validate(dept),
        invitation=inv,
    )


def create_department_admin_invitation_controller(
    session: Session,
    current_user: UserModel,
    data: DepartmentAdminInvitationCreate,
) -> InvitationCreatedResponse:
    """Handles logic for creating a department admin invitation for an existing department"""
    try:
        return create_department_admin_invitation(
            session=session,
            current_user=current_user,
            data=data,
        )
    except (UnauthorizedInvitationActionError) as exc:
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail=str(exc)
        ) from exc
    except (DepartmentNotFoundError) as exc:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail=str(exc)
        ) from exc


def list_invitations_controller(
    session: Session,
    current_user: UserModel
) -> list[InvitationModel]:
    """Handles logic for listing invitations."""
    if current_user.role == UserRole.SUPER_ADMIN:
        return auth_repository.list_all_admin_invitations(session)

    if current_user.role == UserRole.DEPARTMENT_ADMIN:
        return auth_repository.list_invitations_for_department(
            session=session,
            department_id=current_user.department_id,
        )

    raise HTTPException(
        status_code=status.HTTP_403_FORBIDDEN,
        detail="Insufficient permissions for this operation.",
    )


def revoke_invitation_controller(
    session: Session,
    current_user: UserModel,
    invitation_id: uuid.UUID,
) -> InvitationModel:
    """Handles the logic for revoking an invitation."""
    if current_user.role == UserRole.SUPER_ADMIN:
        invitation = auth_repository.get_invitation_by_id(
            session=session,
            invitation_id=invitation_id
        )
    elif current_user.role == UserRole.DEPARTMENT_ADMIN:
        invitation = auth_repository.get_invitation_by_id_for_department(
            session=session,
            invitation_id=invitation_id,
            department_id=current_user.department_id,
        )
    else:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Invitation not found."
        )

    if not invitation:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Invitation not found."
        )

    return auth_repository.revoke_invitation(
        session=session,
        invitation=invitation,
    )


def refresh_token_controller(
    session: Session,
    request: Request,
    response: Response,
    body: Optional[RefreshTokenRequest] = None,
) -> Token:
    """Handles the logic for refreshing a token."""
    if body and body.refresh_token:
        raw_refresh = body.refresh_token
        from_cookie = False
    else:
        raw_refresh = request.cookies.get("refresh_token")
        from_cookie = True

    if raw_refresh is None:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Refresh token required."
        )

    if from_cookie:
        csrf_header = request.headers.get("X-CSRF-Token")
        if not csrf_header:
            raise HTTPException(
                status_code=status.HTTP_403_FORBIDDEN,
                detail="CSRF token header (X-CSRF-Token) missing",
            )

        refresh_hash = hash_refresh_token(raw_refresh)
        current_session = auth_repository.get_refresh_session_by_token_hash(
            session=session,
            token_hash=refresh_hash,
        )

        if current_session and current_session.csrf_token_hash:
            if hash_csrf_token(csrf_header) != current_session.csrf_token_hash:
                raise HTTPException(
                    status_code=status.HTTP_403_FORBIDDEN,
                    detail="Invalid CSRF token."
                )

    try:
        new_session, new_raw_refresh, new_raw_csrf = rotate_refresh_token(
            session=session,
            raw_refresh_token=raw_refresh,
        )
    except (InvalidRefreshTokenError) as exc:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Invalid refresh token",
        ) from exc
    except (RefreshTokenReuseDetectedError) as exc:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Refresh token reuse detected. All sessions in this family have been revoked."
        ) from exc
    except (RevokedRefreshTokenError) as exc:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Refresh token has been revoked",
        ) from exc
    except (ExpiredRefreshTokenError) as exc:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Refresh token expired",
        ) from exc

    access_token = create_access_token({"sub": str(new_session.user_id)})
    response.set_cookie(
        key="refresh_token",
        value=new_raw_refresh,
        httponly=True,
        secure=SECURE_COOKIE,
        samesite="lax",
        max_age=14*24*3600,
    )
    response.set_cookie(
        key="csrf_token",
        value=new_raw_csrf,
        httponly=False,
        secure=SECURE_COOKIE,
        samesite="lax",
        max_age=14*24*3600,
    )

    return Token(
        access_token=access_token,
        token_type="bearer",
        refresh_token=new_raw_refresh,
        csrf_token=new_raw_csrf,
    )


def logout_controller(
    session: Session,
    request: Request,
    response: Response,
    body: Optional[RefreshTokenRequest] = None
) -> dict[str, str]:
    """Handles the logic for a user logging out."""
    if body and body.refresh_token:
        raw_refresh = body.refresh_token
    else:
        raw_refresh = request.cookies.get("refresh_token")

    if raw_refresh:
        terminate_refresh_session(
            session=session,
            raw_refresh_token=raw_refresh
        )

    response.delete_cookie(
        key="refresh_token",
        httponly=True,
        secure=SECURE_COOKIE,
        samesite="lax",
    )
    response.delete_cookie(
        key="csrf_token",
        httponly=False,
        secure=SECURE_COOKIE,
        samesite="lax",
    )

    return {"detail": "Successfully logged out"}
