"""
Unit and business rule tests for super-admin department provisioning and admin invitation services.
"""

import uuid
import datetime
import pytest
from sqlmodel import Session

from src.user.models import UserRole
from src.department.models import Department as DepartmentModel
from src.auth.schemas import (
    DepartmentAdminProvisioningCreate,
    DepartmentAdminInvitationCreate,
    InvitationCreatedResponse,
)
from src.auth.security import hash_invitation_token
from src.auth import repository as auth_repository
from src.auth.services import (
    provision_department_with_admin,
    create_department_admin_invitation,
    UnauthorizedInvitationActionError,
    DepartmentAlreadyExistsError,
    DepartmentNotFoundError,
)


def test_provision_department_with_admin_success(session: Session, user_factory):
    """Tests super-admin provisioning a new department and its initial admin invitation atomically."""
    super_admin = user_factory(role=UserRole.SUPER_ADMIN, department_id=None)
    data = DepartmentAdminProvisioningCreate(
        department_name="Neurology",
        department_code="NEURO",
    )

    dept, response = provision_department_with_admin(
        session=session,
        current_user=super_admin,
        data=data,
    )

    # Validate created department
    assert isinstance(dept, DepartmentModel)
    assert isinstance(dept.id, uuid.UUID)
    assert dept.name == "Neurology"
    assert dept.code == "NEURO"
    assert dept.is_deleted is False

    # Validate returned invitation response
    assert isinstance(response, InvitationCreatedResponse)
    assert response.role == UserRole.DEPARTMENT_ADMIN
    assert response.department_id == dept.id
    assert response.doctor_id is None
    assert response.raw_token is not None
    assert len(response.raw_token) > 20
    assert response.expires_at > datetime.datetime.now(datetime.timezone.utc)

    # Validate database persistence of invitation
    expected_hash = hash_invitation_token(response.raw_token)
    persisted = auth_repository.get_invitation_by_token_hash(session, expected_hash)
    assert persisted is not None
    assert persisted.id == response.id
    assert persisted.department_id == dept.id
    assert persisted.role == UserRole.DEPARTMENT_ADMIN
    assert persisted.created_by_user_id == super_admin.id
    assert persisted.doctor_id is None


def test_provision_department_with_admin_rejects_non_super_admin(
    session: Session, department_factory, user_factory, doctor_factory
):
    """Tests that only SUPER_ADMIN users can provision a new department with an admin."""
    dept = department_factory()
    doc = doctor_factory(department_id=dept.id)
    dept_admin = user_factory(role=UserRole.DEPARTMENT_ADMIN, department_id=dept.id)
    doctor_user = user_factory(role=UserRole.DOCTOR, department_id=dept.id, doctor_id=doc.id)
    viewer_user = user_factory(role=UserRole.VIEWER, department_id=dept.id)

    data = DepartmentAdminProvisioningCreate(
        department_name="Oncology",
        department_code="ONCO",
    )

    for unauthorized_user in (dept_admin, doctor_user, viewer_user):
        with pytest.raises(UnauthorizedInvitationActionError):
            provision_department_with_admin(
                session=session,
                current_user=unauthorized_user,
                data=data,
            )


def test_provision_department_with_admin_rejects_duplicate_department_name(
    session: Session, department_factory, user_factory
):
    """Tests that attempting to provision a department with an already existing name is rejected."""
    existing_dept = department_factory(name="Existing Dept")
    super_admin = user_factory(role=UserRole.SUPER_ADMIN, department_id=None)

    data = DepartmentAdminProvisioningCreate(
        department_name=existing_dept.name,
        department_code="NEWCODE",
    )

    with pytest.raises(DepartmentAlreadyExistsError):
        provision_department_with_admin(
            session=session,
            current_user=super_admin,
            data=data,
        )


def test_create_department_admin_invitation_existing_department_success(
    session: Session, department_factory, user_factory
):
    """Tests super-admin generating an admin invitation for an existing department."""
    dept = department_factory()
    super_admin = user_factory(role=UserRole.SUPER_ADMIN, department_id=None)

    data = DepartmentAdminInvitationCreate(department_id=dept.id)
    response = create_department_admin_invitation(
        session=session,
        current_user=super_admin,
        data=data,
    )

    assert isinstance(response, InvitationCreatedResponse)
    assert response.role == UserRole.DEPARTMENT_ADMIN
    assert response.department_id == dept.id
    assert response.doctor_id is None
    assert response.raw_token is not None

    expected_hash = hash_invitation_token(response.raw_token)
    persisted = auth_repository.get_invitation_by_token_hash(session, expected_hash)
    assert persisted is not None
    assert persisted.id == response.id
    assert persisted.department_id == dept.id
    assert persisted.role == UserRole.DEPARTMENT_ADMIN
    assert persisted.created_by_user_id == super_admin.id


def test_create_department_admin_invitation_rejects_non_super_admin(
    session: Session, department_factory, user_factory, doctor_factory
):
    """Tests that only SUPER_ADMIN users can generate admin invitations for existing departments."""
    dept = department_factory()
    doc = doctor_factory(department_id=dept.id)
    dept_admin = user_factory(role=UserRole.DEPARTMENT_ADMIN, department_id=dept.id)
    doctor_user = user_factory(role=UserRole.DOCTOR, department_id=dept.id, doctor_id=doc.id)
    viewer_user = user_factory(role=UserRole.VIEWER, department_id=dept.id)

    data = DepartmentAdminInvitationCreate(department_id=dept.id)

    for unauthorized_user in (dept_admin, doctor_user, viewer_user):
        with pytest.raises(UnauthorizedInvitationActionError):
            create_department_admin_invitation(
                session=session,
                current_user=unauthorized_user,
                data=data,
            )


def test_create_department_admin_invitation_rejects_nonexistent_or_deleted_department(
    session: Session, department_factory, user_factory
):
    """Tests that generating an admin invitation for a nonexistent or soft-deleted department is rejected."""
    super_admin = user_factory(role=UserRole.SUPER_ADMIN, department_id=None)

    # Nonexistent department
    data_nonexistent = DepartmentAdminInvitationCreate(department_id=uuid.uuid4())
    with pytest.raises(DepartmentNotFoundError):
        create_department_admin_invitation(
            session=session,
            current_user=super_admin,
            data=data_nonexistent,
        )

    # Soft-deleted department
    deleted_dept = department_factory()
    deleted_dept.is_deleted = True
    session.add(deleted_dept)
    session.commit()

    data_deleted = DepartmentAdminInvitationCreate(department_id=deleted_dept.id)
    with pytest.raises(DepartmentNotFoundError):
        create_department_admin_invitation(
            session=session,
            current_user=super_admin,
            data=data_deleted,
        )
