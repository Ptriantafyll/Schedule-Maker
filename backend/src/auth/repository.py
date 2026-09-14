"""
Database functions for auth components
"""

from __future__ import annotations

import datetime
import uuid
from typing import Optional
from sqlmodel import Session, select, not_, desc
from src.auth.models import Invitation as InvitationModel
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
