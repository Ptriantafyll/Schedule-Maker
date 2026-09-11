"""
Unit tests for the Invitation ORM model.
"""

import datetime
import uuid
import pytest
from sqlalchemy.exc import IntegrityError
from sqlmodel import Session

from src.auth.models import Invitation
from src.user.models import UserRole


def test_invitation_model_creation_defaults(session: Session, department_factory, user_factory):
    """Tests creating an invitation directly via SQLModel and verifies DB defaults."""
    dept = department_factory()
    admin = user_factory(role=UserRole.DEPARTMENT_ADMIN, department_id=dept.id)
    expires_at = datetime.datetime.now(datetime.timezone.utc) + datetime.timedelta(days=7)

    invitation = Invitation(
        token_hash="a" * 64,
        role=UserRole.DOCTOR,
        department_id=dept.id,
        created_by_user_id=admin.id,
        expires_at=expires_at,
    )
    session.add(invitation)
    session.commit()
    session.refresh(invitation)

    assert isinstance(invitation.id, uuid.UUID)
    assert invitation.token_hash == "a" * 64
    assert invitation.role == UserRole.DOCTOR
    assert invitation.department_id == dept.id
    assert invitation.created_by_user_id == admin.id
    assert invitation.doctor_id is None
    assert invitation.used_at is None
    assert invitation.revoked_at is None
    assert invitation.is_deleted is False
    assert invitation.sync_status is False
    assert isinstance(invitation.created_at, datetime.datetime)
    assert isinstance(invitation.updated_at, datetime.datetime)


def test_invitation_model_with_linked_doctor(invitation_factory, doctor_factory):
    """Tests creating an invitation linked to an existing doctor."""
    doctor = doctor_factory()
    invitation = invitation_factory(
        department_id=doctor.department_id,
        doctor_id=doctor.id,
        role=UserRole.DOCTOR,
    )

    assert invitation.doctor_id == doctor.id
    assert invitation.role == UserRole.DOCTOR


def test_invitation_model_with_viewer_role(invitation_factory):
    """Tests creating an invitation for the VIEWER role."""
    invitation = invitation_factory(role=UserRole.VIEWER)

    assert invitation.role == UserRole.VIEWER
    assert invitation.doctor_id is None


def test_invitation_model_with_department_admin_role(invitation_factory, department_factory, user_factory):
    """Tests creating an invitation for the DEPARTMENT_ADMIN role."""
    dept = department_factory()
    super_admin = user_factory(role=UserRole.SUPER_ADMIN, department_id=None)
    invitation = invitation_factory(
        department_id=dept.id,
        created_by_user_id=super_admin.id,
        role=UserRole.DEPARTMENT_ADMIN,
    )

    assert invitation.role == UserRole.DEPARTMENT_ADMIN
    assert invitation.department_id == dept.id
    assert invitation.created_by_user_id == super_admin.id


def test_invitation_token_hash_unique_constraint(session: Session, invitation_factory):
    """Tests that duplicate token_hash values are rejected by the database unique constraint."""
    duplicate_hash = "e" * 64
    invitation_factory(token_hash=duplicate_hash)

    with pytest.raises(IntegrityError):
        invitation_factory(token_hash=duplicate_hash)

    session.rollback()


def test_invitation_is_expired_property():
    """Tests the is_expired helper property."""
    now = datetime.datetime.now(datetime.timezone.utc)

    future_invitation = Invitation(
        token_hash="f" * 64,
        role=UserRole.DOCTOR,
        department_id=uuid.uuid4(),
        created_by_user_id=uuid.uuid4(),
        expires_at=now + datetime.timedelta(hours=1),
    )
    assert future_invitation.is_expired is False

    past_invitation = Invitation(
        token_hash="1" * 64,
        role=UserRole.DOCTOR,
        department_id=uuid.uuid4(),
        created_by_user_id=uuid.uuid4(),
        expires_at=now - datetime.timedelta(minutes=5),
    )
    assert past_invitation.is_expired is True


def test_invitation_is_active_property():
    """Tests the is_active helper property across lifecycles."""
    now = datetime.datetime.now(datetime.timezone.utc)

    active_invitation = Invitation(
        token_hash="2" * 64,
        role=UserRole.DOCTOR,
        department_id=uuid.uuid4(),
        created_by_user_id=uuid.uuid4(),
        expires_at=now + datetime.timedelta(days=1),
    )
    assert active_invitation.is_active is True

    used_invitation = Invitation(
        token_hash="3" * 64,
        role=UserRole.DOCTOR,
        department_id=uuid.uuid4(),
        created_by_user_id=uuid.uuid4(),
        expires_at=now + datetime.timedelta(days=1),
        used_at=now,
    )
    assert used_invitation.is_active is False

    revoked_invitation = Invitation(
        token_hash="4" * 64,
        role=UserRole.DOCTOR,
        department_id=uuid.uuid4(),
        created_by_user_id=uuid.uuid4(),
        expires_at=now + datetime.timedelta(days=1),
        revoked_at=now,
    )
    assert revoked_invitation.is_active is False

    expired_invitation = Invitation(
        token_hash="5" * 64,
        role=UserRole.DOCTOR,
        department_id=uuid.uuid4(),
        created_by_user_id=uuid.uuid4(),
        expires_at=now - datetime.timedelta(seconds=10),
    )
    assert expired_invitation.is_active is False

    deleted_invitation = Invitation(
        token_hash="6" * 64,
        role=UserRole.DOCTOR,
        department_id=uuid.uuid4(),
        created_by_user_id=uuid.uuid4(),
        expires_at=now + datetime.timedelta(days=1),
        is_deleted=True,
    )
    assert deleted_invitation.is_active is False
