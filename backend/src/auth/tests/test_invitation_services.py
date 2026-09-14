"""
Unit and business rule tests for department admin staff invitation services.
"""

import uuid
import datetime
import pytest
from sqlmodel import Session

from src.user.models import UserRole
from src.auth.schemas import StaffInvitationCreate, InvitationCreatedResponse
from src.auth.security import hash_invitation_token
from src.auth import repository as auth_repository
from src.auth.services import (
    create_staff_invitation,
    UnauthorizedInvitationActionError,
    InvalidInvitationDoctorError,
    DoctorAlreadyLinkedError,
    DoctorInvitationAlreadyPendingError,
    InvalidInvitationRoleError,
)


def test_create_staff_invitation_doctor_unlinked(session: Session, department_factory, user_factory):
    """Tests department admin creating a generic DOCTOR invitation without pre-assigning doctor_id."""
    dept = department_factory()
    admin = user_factory(role=UserRole.DEPARTMENT_ADMIN, department_id=dept.id)

    data = StaffInvitationCreate(role=UserRole.DOCTOR, doctor_id=None)
    response = create_staff_invitation(
        session=session, current_user=admin, data=data)

    assert isinstance(response, InvitationCreatedResponse)
    assert isinstance(response.id, uuid.UUID)
    assert response.role == UserRole.DOCTOR
    assert response.department_id == dept.id
    assert response.doctor_id is None
    assert response.raw_token is not None
    assert len(response.raw_token) > 20
    assert response.expires_at > datetime.datetime.now(datetime.timezone.utc)

    # Verify database persistence with SHA-256 hash
    expected_hash = hash_invitation_token(response.raw_token)
    persisted = auth_repository.get_invitation_by_token_hash(
        session, expected_hash)
    assert persisted is not None
    assert persisted.id == response.id
    assert persisted.created_by_user_id == admin.id
    assert persisted.department_id == dept.id
    assert persisted.role == UserRole.DOCTOR
    assert persisted.doctor_id is None
    assert persisted.used_at is None
    assert persisted.revoked_at is None


def test_create_staff_invitation_doctor_with_existing_doctor(
    session: Session, department_factory, user_factory, doctor_factory
):
    """Tests department admin creating a DOCTOR invitation linked to an unlinked existing doctor."""
    dept = department_factory()
    admin = user_factory(role=UserRole.DEPARTMENT_ADMIN, department_id=dept.id)
    doc = doctor_factory(department_id=dept.id)

    data = StaffInvitationCreate(role=UserRole.DOCTOR, doctor_id=doc.id)
    response = create_staff_invitation(
        session=session, current_user=admin, data=data)

    assert response.doctor_id == doc.id
    assert response.department_id == dept.id
    assert response.role == UserRole.DOCTOR

    expected_hash = hash_invitation_token(response.raw_token)
    persisted = auth_repository.get_invitation_by_token_hash(
        session, expected_hash)
    assert persisted is not None
    assert persisted.doctor_id == doc.id


def test_create_staff_invitation_viewer(session: Session, department_factory, user_factory):
    """Tests department admin creating a VIEWER invitation."""
    dept = department_factory()
    admin = user_factory(role=UserRole.DEPARTMENT_ADMIN, department_id=dept.id)

    data = StaffInvitationCreate(role=UserRole.VIEWER, doctor_id=None)
    response = create_staff_invitation(
        session=session, current_user=admin, data=data)

    assert response.role == UserRole.VIEWER
    assert response.department_id == dept.id
    assert response.doctor_id is None


def test_create_staff_invitation_rejects_non_admin_caller(
    session: Session, department_factory, user_factory, doctor_factory
):
    """Tests that only DEPARTMENT_ADMIN users can call create_staff_invitation."""
    dept = department_factory()
    doc = doctor_factory(department_id=dept.id)
    doctor_user = user_factory(
        role=UserRole.DOCTOR, department_id=dept.id, doctor_id=doc.id)
    viewer_user = user_factory(role=UserRole.VIEWER, department_id=dept.id)
    super_admin = user_factory(role=UserRole.SUPER_ADMIN, department_id=None)

    data = StaffInvitationCreate(role=UserRole.DOCTOR)

    for unauthorized_user in (doctor_user, viewer_user, super_admin):
        with pytest.raises(UnauthorizedInvitationActionError):
            create_staff_invitation(
                session=session, current_user=unauthorized_user, data=data)


def test_create_staff_invitation_rejects_doctor_from_another_department(
    session: Session, department_factory, user_factory, doctor_factory
):
    """Tests that department admins cannot link a doctor belonging to a different department."""
    dept_a = department_factory()
    dept_b = department_factory()
    admin_a = user_factory(role=UserRole.DEPARTMENT_ADMIN,
                           department_id=dept_a.id)
    doctor_b = doctor_factory(department_id=dept_b.id)

    data = StaffInvitationCreate(role=UserRole.DOCTOR, doctor_id=doctor_b.id)

    with pytest.raises(InvalidInvitationDoctorError):
        create_staff_invitation(
            session=session, current_user=admin_a, data=data)


def test_create_staff_invitation_rejects_nonexistent_or_deleted_doctor(
    session: Session, department_factory, user_factory, doctor_factory
):
    """Tests that referencing nonexistent or soft-deleted doctors is rejected."""
    dept = department_factory()
    admin = user_factory(role=UserRole.DEPARTMENT_ADMIN, department_id=dept.id)

    # Nonexistent doctor ID
    data_nonexistent = StaffInvitationCreate(
        role=UserRole.DOCTOR, doctor_id=uuid.uuid4())
    with pytest.raises(InvalidInvitationDoctorError):
        create_staff_invitation(
            session=session, current_user=admin, data=data_nonexistent)

    # Soft-deleted doctor
    deleted_doc = doctor_factory(department_id=dept.id)
    deleted_doc.is_deleted = True
    session.add(deleted_doc)
    session.commit()
    data_deleted = StaffInvitationCreate(
        role=UserRole.DOCTOR, doctor_id=deleted_doc.id)
    with pytest.raises(InvalidInvitationDoctorError):
        create_staff_invitation(
            session=session, current_user=admin, data=data_deleted)


def test_create_staff_invitation_rejects_already_linked_doctor(
    session: Session, department_factory, user_factory, doctor_factory
):
    """Tests that a doctor already linked to an active user account cannot receive an invitation."""
    dept = department_factory()
    admin = user_factory(role=UserRole.DEPARTMENT_ADMIN, department_id=dept.id)
    doc = doctor_factory(department_id=dept.id)
    # Link doc to a user account
    user_factory(role=UserRole.DOCTOR, department_id=dept.id, doctor_id=doc.id)

    data = StaffInvitationCreate(role=UserRole.DOCTOR, doctor_id=doc.id)
    with pytest.raises(DoctorAlreadyLinkedError):
        create_staff_invitation(session=session, current_user=admin, data=data)


def test_create_staff_invitation_rejects_doctor_with_active_pending_invitation(
    session: Session, department_factory, user_factory, doctor_factory, invitation_factory
):
    """Tests that a doctor with an active pending invitation cannot have a duplicate active invitation."""
    dept = department_factory()
    admin = user_factory(role=UserRole.DEPARTMENT_ADMIN, department_id=dept.id)
    doc = doctor_factory(department_id=dept.id)

    # Create an active pending invitation for this doctor
    invitation_factory(
        department_id=dept.id,
        doctor_id=doc.id,
        role=UserRole.DOCTOR,
    )

    data = StaffInvitationCreate(role=UserRole.DOCTOR, doctor_id=doc.id)
    with pytest.raises(DoctorInvitationAlreadyPendingError):
        create_staff_invitation(session=session, current_user=admin, data=data)


def test_create_staff_invitation_allows_doctor_if_previous_invitation_inactive(
    session: Session, department_factory, user_factory, doctor_factory, invitation_factory
):
    """Tests that a doctor can receive a new invitation if previous invitations are expired, revoked, or used."""
    dept = department_factory()
    admin = user_factory(role=UserRole.DEPARTMENT_ADMIN, department_id=dept.id)
    doc = doctor_factory(department_id=dept.id)
    now = datetime.datetime.now(datetime.timezone.utc)

    # Inactive invitation 1: expired
    invitation_factory(
        department_id=dept.id,
        doctor_id=doc.id,
        expires_at=now - datetime.timedelta(days=1),
    )
    # Inactive invitation 2: revoked
    invitation_factory(
        department_id=dept.id,
        doctor_id=doc.id,
        revoked_at=now,
    )

    data = StaffInvitationCreate(role=UserRole.DOCTOR, doctor_id=doc.id)
    response = create_staff_invitation(
        session=session, current_user=admin, data=data)
    assert response.doctor_id == doc.id


def test_create_staff_invitation_rejects_viewer_with_doctor_id(
    session: Session, department_factory, user_factory, doctor_factory
):
    """Tests that a VIEWER invitation cannot specify a doctor_id."""
    dept = department_factory()
    admin = user_factory(role=UserRole.DEPARTMENT_ADMIN, department_id=dept.id)
    doc = doctor_factory(department_id=dept.id)

    data = StaffInvitationCreate(role=UserRole.VIEWER, doctor_id=doc.id)
    with pytest.raises(InvalidInvitationDoctorError):
        create_staff_invitation(session=session, current_user=admin, data=data)
