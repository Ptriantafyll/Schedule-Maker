"""
Unit and query tests for the Invitation database repository.
"""

import uuid
import datetime
from sqlmodel import Session

from src.user.models import UserRole
from src.auth import repository as auth_repository


def test_create_invitation(session: Session, department_factory, user_factory):
    """Tests persisting a new invitation via create_invitation."""
    dept = department_factory()
    admin = user_factory(role=UserRole.DEPARTMENT_ADMIN, department_id=dept.id)
    expires_at = datetime.datetime.now(datetime.timezone.utc) + datetime.timedelta(days=7)
    token_hash = "a" * 64

    invitation = auth_repository.create_invitation(
        session=session,
        token_hash=token_hash,
        role=UserRole.DOCTOR,
        department_id=dept.id,
        created_by_user_id=admin.id,
        expires_at=expires_at,
    )

    assert isinstance(invitation.id, uuid.UUID)
    assert invitation.token_hash == token_hash
    assert invitation.role == UserRole.DOCTOR
    assert invitation.department_id == dept.id
    assert invitation.created_by_user_id == admin.id
    assert invitation.doctor_id is None
    assert invitation.used_at is None
    assert invitation.revoked_at is None
    assert invitation.is_deleted is False


def test_create_invitation_with_doctor_id(session: Session, department_factory, user_factory, doctor_factory):
    """Tests persisting an invitation with an explicit doctor_id."""
    dept = department_factory()
    admin = user_factory(role=UserRole.DEPARTMENT_ADMIN, department_id=dept.id)
    doc = doctor_factory(department_id=dept.id)
    expires_at = datetime.datetime.now(datetime.timezone.utc) + datetime.timedelta(days=7)

    invitation = auth_repository.create_invitation(
        session=session,
        token_hash="b" * 64,
        role=UserRole.DOCTOR,
        department_id=dept.id,
        created_by_user_id=admin.id,
        expires_at=expires_at,
        doctor_id=doc.id,
    )

    assert invitation.doctor_id == doc.id


def test_get_invitation_by_token_hash(session: Session, invitation_factory):
    """Tests retrieving an invitation by its token hash."""
    target_hash = "c" * 64
    invitation = invitation_factory(token_hash=target_hash)

    # Found
    retrieved = auth_repository.get_invitation_by_token_hash(session, target_hash)
    assert retrieved is not None
    assert retrieved.id == invitation.id

    # Not found
    assert auth_repository.get_invitation_by_token_hash(session, "nonexistent_hash") is None

    # Soft deleted is hidden
    deleted_hash = "d" * 64
    invitation_factory(token_hash=deleted_hash, is_deleted=True)
    assert auth_repository.get_invitation_by_token_hash(session, deleted_hash) is None


def test_get_invitation_by_id_for_department(session: Session, invitation_factory, department_factory):
    """Tests retrieving an invitation scoped to a department."""
    dept_a = department_factory()
    dept_b = department_factory()
    invitation = invitation_factory(department_id=dept_a.id)

    # Found within own department
    retrieved = auth_repository.get_invitation_by_id_for_department(
        session=session,
        invitation_id=invitation.id,
        department_id=dept_a.id,
    )
    assert retrieved is not None
    assert retrieved.id == invitation.id

    # Hidden from foreign department
    assert auth_repository.get_invitation_by_id_for_department(
        session=session,
        invitation_id=invitation.id,
        department_id=dept_b.id,
    ) is None

    # Hidden when soft deleted
    deleted_invitation = invitation_factory(department_id=dept_a.id, is_deleted=True)
    assert auth_repository.get_invitation_by_id_for_department(
        session=session,
        invitation_id=deleted_invitation.id,
        department_id=dept_a.id,
    ) is None


def test_get_invitation_by_id(session: Session, invitation_factory):
    """Tests global lookup by invitation ID."""
    invitation = invitation_factory()

    retrieved = auth_repository.get_invitation_by_id(session, invitation.id)
    assert retrieved is not None
    assert retrieved.id == invitation.id

    # Soft deleted is hidden
    deleted_invitation = invitation_factory(is_deleted=True)
    assert auth_repository.get_invitation_by_id(session, deleted_invitation.id) is None

    # Nonexistent ID
    assert auth_repository.get_invitation_by_id(session, uuid.uuid4()) is None


def test_get_active_invitation_for_doctor(session: Session, invitation_factory, doctor_factory):
    """Tests retrieving active pending invitations for a specific doctor."""
    doctor = doctor_factory()
    now = datetime.datetime.now(datetime.timezone.utc)

    # Active invitation found
    active_inv = invitation_factory(
        department_id=doctor.department_id,
        doctor_id=doctor.id,
        expires_at=now + datetime.timedelta(days=3),
    )
    retrieved = auth_repository.get_active_invitation_for_doctor(session, doctor.id)
    assert retrieved is not None
    assert retrieved.id == active_inv.id

    # Another doctor has no active invitation
    other_doctor = doctor_factory()
    assert auth_repository.get_active_invitation_for_doctor(session, other_doctor.id) is None


def test_get_active_invitation_for_doctor_ignores_inactive(session: Session, invitation_factory, doctor_factory):
    """Tests that consumed, revoked, expired, or deleted doctor invitations are ignored."""
    doctor = doctor_factory()
    now = datetime.datetime.now(datetime.timezone.utc)

    # Consumed invitation
    invitation_factory(
        department_id=doctor.department_id,
        doctor_id=doctor.id,
        used_at=now,
    )
    assert auth_repository.get_active_invitation_for_doctor(session, doctor.id) is None

    # Revoked invitation
    invitation_factory(
        department_id=doctor.department_id,
        doctor_id=doctor.id,
        revoked_at=now,
    )
    assert auth_repository.get_active_invitation_for_doctor(session, doctor.id) is None

    # Expired invitation
    invitation_factory(
        department_id=doctor.department_id,
        doctor_id=doctor.id,
        expires_at=now - datetime.timedelta(days=1),
    )
    assert auth_repository.get_active_invitation_for_doctor(session, doctor.id) is None


def test_list_invitations_for_department(session: Session, invitation_factory, department_factory):
    """Tests listing all active invitations belonging to a specific department."""
    dept_a = department_factory()
    dept_b = department_factory()

    inv_a1 = invitation_factory(department_id=dept_a.id)
    inv_a2 = invitation_factory(department_id=dept_a.id)
    inv_b = invitation_factory(department_id=dept_b.id)
    inv_a_deleted = invitation_factory(department_id=dept_a.id, is_deleted=True)

    results = auth_repository.list_invitations_for_department(session, dept_a.id)
    result_ids = {r.id for r in results}

    assert inv_a1.id in result_ids
    assert inv_a2.id in result_ids
    assert inv_b.id not in result_ids
    assert inv_a_deleted.id not in result_ids


def test_list_all_admin_invitations(session: Session, invitation_factory):
    """Tests listing all department administrator invitations."""
    admin_inv1 = invitation_factory(role=UserRole.DEPARTMENT_ADMIN)
    admin_inv2 = invitation_factory(role=UserRole.DEPARTMENT_ADMIN)
    doctor_inv = invitation_factory(role=UserRole.DOCTOR)
    viewer_inv = invitation_factory(role=UserRole.VIEWER)
    deleted_admin_inv = invitation_factory(role=UserRole.DEPARTMENT_ADMIN, is_deleted=True)

    results = auth_repository.list_all_admin_invitations(session)
    result_ids = {r.id for r in results}

    assert admin_inv1.id in result_ids
    assert admin_inv2.id in result_ids
    assert doctor_inv.id not in result_ids
    assert viewer_inv.id not in result_ids
    assert deleted_admin_inv.id not in result_ids


def test_revoke_invitation(session: Session, invitation_factory):
    """Tests revoking an invitation sets revoked_at and marks it inactive."""
    invitation = invitation_factory()
    assert invitation.revoked_at is None
    assert invitation.is_active is True

    revoked = auth_repository.revoke_invitation(session, invitation)

    assert revoked.id == invitation.id
    assert revoked.revoked_at is not None
    assert revoked.is_active is False
