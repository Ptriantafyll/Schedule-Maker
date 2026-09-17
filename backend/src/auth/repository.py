"""
Database functions for auth components
"""

from __future__ import annotations

import datetime
import uuid
from typing import Optional
from sqlmodel import Session, select, not_, desc
from src.auth.models import Invitation as InvitationModel
from src.auth.models import RefreshSession as RefreshSessionModel
from src.user.models import UserRole


def create_invitation(
    session: Session,
    token_hash: str,
    role: UserRole,
    department_id: uuid.UUID,
    created_by_user_id: uuid.UUID,
    expires_at: datetime.datetime,
    doctor_id: Optional[uuid.UUID] = None,
) -> InvitationModel:
    """Creates an invitation in the db"""
    new_invitation = InvitationModel(
        token_hash=token_hash,
        role=role,
        department_id=department_id,
        doctor_id=doctor_id,
        created_by_user_id=created_by_user_id,
        expires_at=expires_at,
    )

    session.add(new_invitation)
    session.commit()
    session.refresh(new_invitation)

    return new_invitation


def get_invitation_by_token_hash(
    session: Session,
    token_hash: str,
) -> InvitationModel | None:
    """Retrieves an invitation by its token hash"""
    statement = select(InvitationModel).where(
        InvitationModel.token_hash == token_hash,
        not_(InvitationModel.is_deleted)
    )

    return session.exec(statement).first()


def get_invitation_by_id_for_department(
    session: Session,
    invitation_id: uuid.UUID,
    department_id: uuid.UUID,
) -> InvitationModel | None:
    """Retrieves an invitation by its id for a department id"""
    statement = select(InvitationModel).where(
        InvitationModel.id == invitation_id,
        InvitationModel.department_id == department_id,
        not_(InvitationModel.is_deleted),
    )

    return session.exec(statement).first()


def get_invitation_by_id(
    session: Session,
    invitation_id: uuid.UUID,
) -> InvitationModel | None:
    """Retrieves an invitation by its id"""
    statement = select(InvitationModel).where(
        InvitationModel.id == invitation_id,
        not_(InvitationModel.is_deleted),
    )

    return session.exec(statement).first()


def get_active_invitation_for_doctor(
    session: Session,
    doctor_id: uuid.UUID,
) -> InvitationModel | None:
    """Retrieves any active invitation for a doctor id"""
    now = datetime.datetime.now(datetime.timezone.utc)
    statement = select(InvitationModel).where(
        InvitationModel.doctor_id == doctor_id,
        not_(InvitationModel.is_deleted),
        InvitationModel.used_at == None,
        InvitationModel.revoked_at == None,
        InvitationModel.expires_at > now
    )

    return session.exec(statement).first()


def list_invitations_for_department(
    session: Session,
    department_id: uuid.UUID,
) -> list[InvitationModel]:
    """Retrieves all invitations for a department"""
    statement = select(InvitationModel).where(
        InvitationModel.department_id == department_id,
        not_(InvitationModel.is_deleted),
    ).order_by(desc(InvitationModel.created_at))

    return list(session.exec(statement).all())


def list_all_admin_invitations(session: Session) -> list[InvitationModel]:
    """Retrieves all admin invitations"""
    statement = select(InvitationModel).where(
        InvitationModel.role == UserRole.DEPARTMENT_ADMIN,
        not_(InvitationModel.is_deleted),
    ).order_by(desc(InvitationModel.created_at))

    return list(session.exec(statement).all())


def revoke_invitation(
    session: Session,
    invitation: InvitationModel
) -> InvitationModel:
    """Revokes an invitation"""
    now = datetime.datetime.now(datetime.timezone.utc)
    invitation.revoked_at = now

    session.add(invitation)
    session.commit()
    session.refresh(invitation)

    return invitation


def mark_invitation_as_used(
    session: Session,
    invitation: InvitationModel,
) -> InvitationModel:
    """Consumes an invitation."""
    now = datetime.datetime.now(datetime.timezone.utc)
    invitation.used_at = now

    session.add(invitation)

    return invitation


def create_refresh_session(  # pylint: disable=too-many-arguments,too-many-positional-arguments
    session: Session,
    user_id: uuid.UUID,
    refresh_token_hash: str,
    expires_at: datetime.datetime,
    session_family: Optional[uuid.UUID] = None,
    csrf_token_hash: Optional[str] = None,
) -> RefreshSessionModel:
    """Persist a new refresh session in the db."""
    if session_family is None:
        session_family = uuid.uuid4()

    refresh_session = RefreshSessionModel(
        user_id=user_id,
        refresh_token_hash=refresh_token_hash,
        session_family=session_family,
        csrf_token_hash=csrf_token_hash,
        expires_at=expires_at,
    )
    session.add(refresh_session)
    session.commit()
    session.refresh(refresh_session)

    return refresh_session


def get_refresh_session_by_token_hash(
    session: Session,
    token_hash: str,
) -> RefreshSessionModel | None:
    """Retrieves a refresh session by its token hash."""
    statement = select(RefreshSessionModel).where(
        RefreshSessionModel.refresh_token_hash == token_hash,
        not_(RefreshSessionModel.is_deleted)
    )
    return session.exec(statement).first()


def rotate_refresh_session(
    session: Session,
    current_session: RefreshSessionModel,
    new_refresh_token_hash: str,
    new_expires_at: datetime.datetime,
    new_csrf_token_hash: Optional[str] = None,
) -> RefreshSessionModel:
    """Rotates a refresh session."""
    successor_session = RefreshSessionModel(
        user_id=current_session.user_id,
        refresh_token_hash=new_refresh_token_hash,
        session_family=current_session.session_family,
        csrf_token_hash=new_csrf_token_hash,
        expires_at=new_expires_at,
    )
    session.add(successor_session)
    session.flush()

    # Revoke old session
    now = datetime.datetime.now(datetime.timezone.utc)
    current_session.revoked_at = now
    current_session.replaced_by_session_id = successor_session.id
    current_session.revoked_reason = "rotated"
    current_session.last_used_at = now

    session.add(current_session)
    session.commit()
    session.refresh(successor_session)

    return successor_session


def revoke_refresh_session(
    session: Session,
    refresh_session: RefreshSessionModel,
    reason: str = "logout",
) -> RefreshSessionModel:
    """Revoke a refresh session."""
    now = datetime.datetime.now(datetime.timezone.utc)
    refresh_session.revoked_at = now
    refresh_session.revoked_reason = reason
    session.add(refresh_session)
    session.commit()
    session.refresh(refresh_session)

    return refresh_session


def revoke_session_family(
    session: Session,
    session_family: uuid.UUID,
    reason: str = "reuse_detected",
) -> int:
    """Revokes a session family."""
    statement = select(RefreshSessionModel).where(
        RefreshSessionModel.session_family == session_family,
        not_(RefreshSessionModel.is_deleted),
        RefreshSessionModel.revoked_at == None,
    )

    session_family_to_revoke = list(session.exec(statement).all())

    now = datetime.datetime.now(datetime.timezone.utc)
    for refresh_session in session_family_to_revoke:
        refresh_session.revoked_at = now
        refresh_session.revoked_reason = reason

        session.add(refresh_session)

    session.commit()
    return len(session_family_to_revoke)


def revoke_all_sessions_for_user(
    session: Session,
    user_id: uuid.UUID,
    reason: str = "user_locked",
) -> int:
    """Revokes all sessions of a user."""
    statement = select(RefreshSessionModel).where(
        RefreshSessionModel.user_id == user_id,
        not_(RefreshSessionModel.is_deleted),
        RefreshSessionModel.revoked_at == None,
    )

    sessions_to_revoke = list(session.exec(statement).all())

    now = datetime.datetime.now(datetime.timezone.utc)
    for refresh_session in sessions_to_revoke:
        refresh_session.revoked_at = now
        refresh_session.revoked_reason = reason
        session.add(refresh_session)

    session.commit()

    return len(sessions_to_revoke)


def list_active_sessions_for_user(
    session: Session,
    user_id: uuid.UUID
) -> list[RefreshSessionModel]:
    """Lists the active refresh sessions for a user."""
    now = datetime.datetime.now(datetime.timezone.utc)
    statement = select(RefreshSessionModel).where(
        RefreshSessionModel.user_id == user_id,
        not_(RefreshSessionModel.is_deleted),
        RefreshSessionModel.revoked_at == None,
        RefreshSessionModel.replaced_by_session_id == None,
        RefreshSessionModel.expires_at > now,
    ).order_by(desc(RefreshSessionModel.created_at))

    return list(session.exec(statement).all())
