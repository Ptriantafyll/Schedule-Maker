"""
Authentication controller functions handling business logic
"""

import uuid
from fastapi import HTTPException, status
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
)
from src.auth.security import verify_password, create_access_token
from src.user.models import UserRole, User as UserModel
from src.auth.models import Invitation as InvitationModel
from src.auth.services import (
    create_staff_invitation,
    provision_department_with_admin,
    create_department_admin_invitation,
    consume_invitation_and_signup,
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
)
from src.user.services import (
    UserEmailAlreadyExistsError,
    InvalidUserAccountRelationshipError,
)
from src.department.schemas import DepartmentRead


def login_controller(email: str, password: str, session: Session) -> Token:
    """Handles logic for logging in"""
    user = user_repository.get_user_by_email(session, email)

    if not user or user.is_deleted or not verify_password(password, user.hashed_password):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Username or password is incorrect",
            headers={"WWW-Authenticate": "Bearer"}
        )

    access_token = create_access_token({"sub": str(user.id)})

    return Token(
        access_token=access_token,
        token_type="bearer"
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
