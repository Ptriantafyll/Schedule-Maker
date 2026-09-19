"""
Auth service module for handling auth-related operations, including invitations.
"""

import uuid
import datetime
from sqlmodel import Session
from src.user.schemas import (
    UserRole,
    UserAccountCreate,
)
from src.auth.schemas import (
    StaffInvitationCreate,
    InvitationCreatedResponse,
    DepartmentAdminProvisioningCreate,
    DepartmentAdminInvitationCreate,
    InvitationSignupRequest,
)
from src.user.services import UserEmailAlreadyExistsError
from src.auth import security
from src.auth.models import RefreshSession as RefreshSessionModel
from src.auth import repository as auth_repository
from src.doctor import repository as doctor_repository
from src.user import repository as user_repository
from src.department import repository as department_repository

from src.user.models import User as UserModel
from src.department.models import Department as DepartmentModel
from src.department.schemas import DepartmentCreate

from src.user import services as user_services


class UnauthorizedInvitationActionError(Exception):
    """
    Raised when the caller is not a DEPARTMENT_ADMIN
    (e.g. Doctor, Viewer, or Super-Admin) attempting to issue staff invites)
    """


class InvalidInvitationDoctorError(Exception):
    """
    Raised if the doctor doesn't exist, is soft-deleted, belongs to another department,
    or if a Viewer invitation attempts to attach a doctor_id.
    """


class DoctorAlreadyLinkedError(Exception):
    """Raised if the referenced doctor is already linked to an active User account."""


class DoctorInvitationAlreadyPendingError(Exception):
    """Raised if the referenced doctor already has an active, unexpired, unrevoked invitation."""


class InvalidInvitationRoleError(Exception):
    """Raised if an invalid role is provided."""


class DepartmentAlreadyExistsError(Exception):
    """Raised when attempting to provision a department with an existing name."""


class DepartmentNotFoundError(Exception):
    """Raised when an invitation references a department that does not exist or is soft-deleted."""


class InvitationNotFoundError(Exception):
    """Raised when trying to consume an invitation that does not exist or is soft-deleted."""


class InvitationAlreadyUsedError(Exception):
    """Raised when trying to consume an invitation that already been used."""


class InvitationRevokedError(Exception):
    """Raised when trying to consume an invitation that has been revoked."""


class InvitationExpiredError(Exception):
    """Raised when trying to consume an invitation that has expired"""


class InvalidRefreshTokenError(Exception):
    """Raised when the presented token is not found in the database."""


class ExpiredRefreshTokenError(Exception):
    """Raised when the session's expires_at timestamp is in the past."""


class RevokedRefreshTokenError(Exception):
    """Raised when the session was previously revoked (e.g. logged out)."""


class RefreshTokenReuseDetectedError(Exception):
    """Raised when an already-replaced predecessor token is presented."""


def _issue_invitation(
    session: Session,
    creator_user_id: uuid.UUID,
    role: UserRole,
    department_id: uuid.UUID,
    doctor_id: uuid.UUID | None = None,
) -> InvitationCreatedResponse:
    """Helper to generate token, calculate expiry, persist invitation, and build response."""
    raw_token = security.generate_invitation_token()
    token_hash = security.hash_invitation_token(raw_token)

    now = datetime.datetime.now(datetime.timezone.utc)
    expires_at = now + datetime.timedelta(days=7)

    persisted = auth_repository.create_invitation(
        session=session,
        token_hash=token_hash,
        role=role,
        department_id=department_id,
        created_by_user_id=creator_user_id,
        expires_at=expires_at,
        doctor_id=doctor_id,
    )

    return InvitationCreatedResponse(
        id=persisted.id,
        role=role,
        department_id=department_id,
        doctor_id=doctor_id,
        expires_at=expires_at,
        raw_token=raw_token,
    )


def create_staff_invitation(
    session: Session,
    current_user: UserModel,
    data: StaffInvitationCreate,
) -> InvitationCreatedResponse:
    """Handles the business logic to create a staff invitation from a department admin"""
    if current_user.role != UserRole.DEPARTMENT_ADMIN:
        raise UnauthorizedInvitationActionError(
            "Insufficient permissions to perform this operation."
        )

    if data.role == UserRole.VIEWER and data.doctor_id is not None:
        raise InvalidInvitationDoctorError(
            "Viewers cannot be linked to doctors."
        )

    if data.doctor_id is not None:
        doctor = doctor_repository.get_doctor_by_id_for_department(
            session=session,
            doctor_id=data.doctor_id,
            department_id=current_user.department_id
        )

        if doctor is None or doctor.is_deleted:
            raise InvalidInvitationDoctorError(
                "Doctor not found."
            )

        existing_linked_user = user_repository.get_user_by_doctor_id(
            session=session,
            doctor_id=data.doctor_id,
        )

        if existing_linked_user:
            raise DoctorAlreadyLinkedError(
                "Doctor already linked to a user."
            )

        active_invitation = auth_repository.get_active_invitation_for_doctor(
            session=session,
            doctor_id=data.doctor_id,
        )

        if active_invitation:
            raise DoctorInvitationAlreadyPendingError(
                "Doctor already has a pending invitation."
            )

    return _issue_invitation(
        session=session,
        creator_user_id=current_user.id,
        role=data.role,
        department_id=current_user.department_id,
        doctor_id=data.doctor_id,
    )


def provision_department_with_admin(
    session: Session,
    current_user: UserModel,
    data: DepartmentAdminProvisioningCreate,
) -> tuple[DepartmentModel, InvitationCreatedResponse]:
    """
    Purpose: A SUPER_ADMIN creates a brand new department and simultaneously 
    issues the first DEPARTMENT_ADMIN invitation token.
    """
    if current_user.role != UserRole.SUPER_ADMIN:
        raise UnauthorizedInvitationActionError(
            "Insufficient permissions for this operation."
        )

    existing_dept = department_repository.get_department_by_name_global(
        session=session,
        name=data.department_name,
    )

    if existing_dept:
        raise DepartmentAlreadyExistsError(
            "Department already exists."
        )

    dept_data = DepartmentCreate(
        name=data.department_name,
        code=data.department_code,
    )
    created_department = department_repository.create_department(
        session=session,
        department_data=dept_data
    )

    invitation_created_response = _issue_invitation(
        session=session,
        creator_user_id=current_user.id,
        role=UserRole.DEPARTMENT_ADMIN,
        department_id=created_department.id,
    )

    return (created_department, invitation_created_response)


def create_department_admin_invitation(
    session: Session,
    current_user: UserModel,
    data: DepartmentAdminInvitationCreate,
) -> InvitationCreatedResponse:
    """
    Purpose: A SUPER_ADMIN issues a DEPARTMENT_ADMIN invitation token
    for an existing department.
    """
    if current_user.role != UserRole.SUPER_ADMIN:
        raise UnauthorizedInvitationActionError(
            "Insufficient permissions for this operation."
        )

    department = department_repository.get_department_by_id_global(
        session=session,
        department_id=data.department_id,
    )

    if department is None or department.is_deleted:
        raise DepartmentNotFoundError(
            "Department not found."
        )

    return _issue_invitation(
        session=session,
        creator_user_id=current_user.id,
        role=UserRole.DEPARTMENT_ADMIN,
        department_id=department.id,
    )


def consume_invitation_and_signup(
    session: Session,
    data: InvitationSignupRequest
) -> UserModel:
    """Consumes an invitation and signs up the user."""
    token_hash = security.hash_invitation_token(data.invitation_token)
    invitation = auth_repository.get_invitation_by_token_hash(
        session=session,
        token_hash=token_hash
    )
    if not invitation or invitation.is_deleted:
        raise InvitationNotFoundError(
            "Invitation not found."
        )

    if invitation.used_at is not None:
        raise InvitationAlreadyUsedError(
            "Invitation already used."
        )

    if invitation.revoked_at is not None:
        raise InvitationRevokedError(
            "Invitation has been revoked."
        )

    if invitation.is_expired:
        raise InvitationExpiredError(
            "Invitation has expired."
        )

    existing_email = user_repository.get_user_by_email(
        session=session,
        user_email=data.email
    )

    if existing_email is not None:
        raise UserEmailAlreadyExistsError(
            "A user with this email already exists."
        )

    full_name = f"{data.first_name} {data.last_name}".strip()
    doctor_id = None
    if invitation.role == UserRole.DOCTOR:
        doctor_id = invitation.doctor_id

        if invitation.doctor_id is None:
            new_doctor = doctor_repository.stage_doctor(
                session=session,
                name=full_name,
                department_id=invitation.department_id,
                team_id=None,
            )
            doctor_id = new_doctor.id

    account_data = UserAccountCreate(
        email=data.email,
        full_name=full_name,
        password=data.password,
        role=invitation.role,
        department_id=invitation.department_id,
        doctor_id=doctor_id
    )

    user = user_services.stage_user_account(
        session=session,
        account_data=account_data
    )
    auth_repository.mark_invitation_as_used(
        session=session,
        invitation=invitation,
    )

    session.commit()
    session.refresh(user)
    return user


def issue_refresh_session(
    session: Session,
    user_id: uuid.UUID,
) -> tuple[RefreshSessionModel, str, str]:
    """Issues a refresh session."""
    raw_refresh = security.generate_refresh_token()
    raw_csrf = security.generate_csrf_token()

    refresh_hash = security.hash_refresh_token(raw_refresh)
    csrf_hash = security.hash_csrf_token(raw_csrf)

    now = datetime.datetime.now(datetime.timezone.utc)
    expires_at = now + datetime.timedelta(days=14)

    refresh_session = auth_repository.create_refresh_session(
        session=session,
        user_id=user_id,
        refresh_token_hash=refresh_hash,
        expires_at=expires_at,
        csrf_token_hash=csrf_hash,
    )

    return (refresh_session, raw_refresh, raw_csrf)


def rotate_refresh_token(
    session: Session,
    raw_refresh_token: str,
) -> tuple[RefreshSessionModel, str, str]:
    """Rotates a refresh token."""
    incoming_hash = security.hash_refresh_token(raw_refresh_token)

    refresh_session = auth_repository.get_refresh_session_by_token_hash(
        session=session,
        token_hash=incoming_hash
    )

    if not refresh_session:
        raise InvalidRefreshTokenError(
            "Refresh session not found."
        )

    if refresh_session.replaced_by_session_id is not None:
        auth_repository.revoke_session_family(
            session=session,
            session_family=refresh_session.session_family,
            reason="reuse_detected",
        )
        raise RefreshTokenReuseDetectedError(
            "Refresh token reuse detected."
        )

    if refresh_session.revoked_at is not None:
        raise RevokedRefreshTokenError(
            "Refresh token has been revoked."
        )

    if refresh_session.is_expired:
        raise ExpiredRefreshTokenError(
            "Refresh token expired."
        )

    user = user_repository.get_user_by_id(
        session=session,
        user_id=refresh_session.user_id,
    )
    if not user or user.is_deleted:
        auth_repository.revoke_session_family(
            session=session,
            session_family=refresh_session.session_family,
            reason="inactive_user",
        )
        raise InvalidRefreshTokenError("User account is inactive or deleted.")

    now = datetime.datetime.now(datetime.timezone.utc)
    new_expires_at = now + datetime.timedelta(days=14)

    new_raw_refresh = security.generate_refresh_token()
    new_refresh_hash = security.hash_refresh_token(new_raw_refresh)

    new_raw_csrf = security.generate_csrf_token()
    new_csrf_hash = security.hash_csrf_token(new_raw_csrf)

    new_refresh_session = auth_repository.rotate_refresh_session(
        session=session,
        current_session=refresh_session,
        new_refresh_token_hash=new_refresh_hash,
        new_expires_at=new_expires_at,
        new_csrf_token_hash=new_csrf_hash,
    )

    return (new_refresh_session, new_raw_refresh, new_raw_csrf)


def terminate_refresh_session(
    session: Session,
    raw_refresh_token: str,
) -> bool:
    """Terminates a refresh session with reason logout."""
    refresh_hash = security.hash_refresh_token(raw_refresh_token)

    refresh_session = auth_repository.get_refresh_session_by_token_hash(
        session=session,
        token_hash=refresh_hash,
    )

    if not refresh_session:
        return False

    revoked_session = auth_repository.revoke_refresh_session(
        session=session,
        refresh_session=refresh_session,
        reason="logout",
    )

    return revoked_session is not None


def revoke_all_user_refresh_sessions(
    session: Session,
    user_id: uuid.UUID,
    reason: str = "logout"
) -> int:
    """Revokes all user refresh sessions."""
    sessions_revoked = auth_repository.revoke_all_sessions_for_user(
        session=session,
        user_id=user_id,
        reason=reason,
    )
    return sessions_revoked
