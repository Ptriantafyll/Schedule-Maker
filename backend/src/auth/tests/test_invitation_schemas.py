"""
Unit and validation tests for Invitation Pydantic schemas (DTOs).
"""

import uuid
import datetime
import pytest
from pydantic import ValidationError

from src.user.models import UserRole
from src.auth.schemas import (
    StaffInvitationCreate,
    DepartmentAdminProvisioningCreate,
    DepartmentAdminInvitationCreate,
    InvitationSignupRequest,
    InvitationCreatedResponse,
    InvitationRead,
)


# ============================================================================
# Group 1: StaffInvitationCreate (Department Admin invites staff)
# ============================================================================


def test_staff_invitation_create_defaults():
    """Tests creating StaffInvitationCreate with defaults (DOCTOR role, no doctor_id)."""
    data = StaffInvitationCreate()
    assert data.role == UserRole.DOCTOR
    assert data.doctor_id is None


def test_staff_invitation_create_viewer():
    """Tests creating StaffInvitationCreate for the VIEWER role."""
    data = StaffInvitationCreate(role=UserRole.VIEWER)
    assert data.role == UserRole.VIEWER
    assert data.doctor_id is None


def test_staff_invitation_create_with_doctor_id():
    """Tests creating StaffInvitationCreate linked to an existing doctor ID."""
    doc_id = uuid.uuid4()
    data = StaffInvitationCreate(role=UserRole.DOCTOR, doctor_id=doc_id)
    assert data.role == UserRole.DOCTOR
    assert data.doctor_id == doc_id


@pytest.mark.parametrize(
    "forbidden_role",
    [
        UserRole.DEPARTMENT_ADMIN,
        UserRole.SUPER_ADMIN,
    ],
)
def test_staff_invitation_create_rejects_admin_or_super_admin_roles(forbidden_role):
    """Tests that department admins cannot invite admins or super admins."""
    with pytest.raises(ValidationError) as exc_info:
        StaffInvitationCreate(role=forbidden_role)

    assert "role" in str(exc_info.value).lower()


def test_staff_invitation_create_rejects_extra_fields():
    """Tests that client-supplied tenant, user, or identity fields are forbidden."""
    with pytest.raises(ValidationError):
        StaffInvitationCreate(department_id=uuid.uuid4())

    with pytest.raises(ValidationError):
        StaffInvitationCreate(email="attacker@hospital.org")

    with pytest.raises(ValidationError):
        StaffInvitationCreate(created_by_user_id=uuid.uuid4())


# ============================================================================
# Group 2: DepartmentAdminProvisioningCreate (Super-Admin provisions department)
# ============================================================================


def test_department_admin_provisioning_create_valid():
    """Tests creating DepartmentAdminProvisioningCreate with valid names and codes."""
    data = DepartmentAdminProvisioningCreate(
        department_name="Cardiology",
        department_code="CARD",
    )
    assert data.department_name == "Cardiology"
    assert data.department_code == "CARD"


def test_department_admin_provisioning_create_normalizes_whitespace():
    """Tests that department names and codes are stripped of whitespace."""
    data = DepartmentAdminProvisioningCreate(
        department_name="  Cardiology  ",
        department_code=" CARD ",
    )
    assert data.department_name == "Cardiology"
    assert data.department_code == "CARD"


@pytest.mark.parametrize(
    "name,code",
    [
        ("", "CARD"),
        ("   ", "CARD"),
        ("Cardiology", ""),
        ("Cardiology", "   "),
    ],
)
def test_department_admin_provisioning_create_rejects_empty_fields(name, code):
    """Tests that empty or whitespace-only department names and codes are rejected."""
    with pytest.raises(ValidationError):
        DepartmentAdminProvisioningCreate(department_name=name, department_code=code)


def test_department_admin_provisioning_create_rejects_extra_fields():
    """Tests that unexpected fields are forbidden on department provisioning."""
    with pytest.raises(ValidationError):
        DepartmentAdminProvisioningCreate(
            department_name="Cardiology",
            department_code="CARD",
            role="super_admin",
        )


# ============================================================================
# Group 3: DepartmentAdminInvitationCreate (Super-Admin invites admin to existing dept)
# ============================================================================


def test_department_admin_invitation_create_valid():
    """Tests creating DepartmentAdminInvitationCreate with a valid department_id."""
    dept_id = uuid.uuid4()
    data = DepartmentAdminInvitationCreate(department_id=dept_id)
    assert data.department_id == dept_id


def test_department_admin_invitation_create_rejects_extra_fields():
    """Tests that extra fields are forbidden on existing department admin invite."""
    with pytest.raises(ValidationError):
        DepartmentAdminInvitationCreate(
            department_id=uuid.uuid4(),
            role="super_admin",
        )


# ============================================================================
# Group 4: InvitationSignupRequest (Public front-door user registration)
# ============================================================================


def test_invitation_signup_request_valid():
    """Tests valid signup request parsing and types."""
    data = InvitationSignupRequest(
        invitation_token="inv_secret_token_123",
        first_name="Gregory",
        last_name="House",
        email="house@hospital.org",
        password="ValidPassword123!",
    )
    assert data.invitation_token == "inv_secret_token_123"
    assert data.first_name == "Gregory"
    assert data.last_name == "House"
    assert data.email == "house@hospital.org"
    assert data.password == "ValidPassword123!"


def test_invitation_signup_request_normalizes_whitespace():
    """Tests that names and token have leading/trailing whitespace stripped."""
    data = InvitationSignupRequest(
        invitation_token="  inv_secret_token_123  ",
        first_name="  Gregory  ",
        last_name="  House  ",
        email="house@hospital.org",
        password="ValidPassword123!",
    )
    assert data.invitation_token == "inv_secret_token_123"
    assert data.first_name == "Gregory"
    assert data.last_name == "House"


def test_invitation_signup_request_rejects_invalid_email():
    """Tests that an invalid email address raises a validation error."""
    with pytest.raises(ValidationError):
        InvitationSignupRequest(
            invitation_token="token_123",
            first_name="Gregory",
            last_name="House",
            email="not-a-valid-email",
            password="ValidPassword123!",
        )


@pytest.mark.parametrize(
    "token,first,last,password",
    [
        ("", "Gregory", "House", "ValidPassword123!"),
        ("   ", "Gregory", "House", "ValidPassword123!"),
        ("token_123", "", "House", "ValidPassword123!"),
        ("token_123", "   ", "House", "ValidPassword123!"),
        ("token_123", "Gregory", "", "ValidPassword123!"),
        ("token_123", "Gregory", "   ", "ValidPassword123!"),
        ("token_123", "Gregory", "House", ""),
    ],
)
def test_invitation_signup_request_rejects_empty_fields(token, first, last, password):
    """Tests that empty or whitespace-only inputs are rejected."""
    with pytest.raises(ValidationError):
        InvitationSignupRequest(
            invitation_token=token,
            first_name=first,
            last_name=last,
            email="house@hospital.org",
            password=password,
        )


def test_invitation_signup_request_rejects_privilege_escalation_fields():
    """Tests that callers cannot supply role, department_id, or doctor_id during signup."""
    # Attempting to supply role
    with pytest.raises(ValidationError):
        InvitationSignupRequest(
            invitation_token="token_123",
            first_name="Gregory",
            last_name="House",
            email="house@hospital.org",
            password="ValidPassword123!",
            role="super_admin",
        )

    # Attempting to supply department_id
    with pytest.raises(ValidationError):
        InvitationSignupRequest(
            invitation_token="token_123",
            first_name="Gregory",
            last_name="House",
            email="house@hospital.org",
            password="ValidPassword123!",
            department_id=uuid.uuid4(),
        )

    # Attempting to supply doctor_id
    with pytest.raises(ValidationError):
        InvitationSignupRequest(
            invitation_token="token_123",
            first_name="Gregory",
            last_name="House",
            email="house@hospital.org",
            password="ValidPassword123!",
            doctor_id=uuid.uuid4(),
        )


# ============================================================================
# Group 5: Response DTOs (InvitationCreatedResponse & InvitationRead)
# ============================================================================


def test_invitation_created_response_includes_raw_token():
    """Tests that InvitationCreatedResponse exposes the raw token upon generation."""
    inv_id = uuid.uuid4()
    dept_id = uuid.uuid4()
    expires_at = datetime.datetime.now(datetime.timezone.utc) + datetime.timedelta(days=7)

    response = InvitationCreatedResponse(
        id=inv_id,
        role=UserRole.DOCTOR,
        department_id=dept_id,
        doctor_id=None,
        expires_at=expires_at,
        raw_token="raw_invitation_secret_xyz",
    )

    assert response.id == inv_id
    assert response.role == UserRole.DOCTOR
    assert response.department_id == dept_id
    assert response.doctor_id is None
    assert response.expires_at == expires_at
    assert response.raw_token == "raw_invitation_secret_xyz"


def test_invitation_read_omits_secrets(invitation_factory):
    """Tests that InvitationRead serializes from ORM models without exposing secrets."""
    invitation = invitation_factory()
    read_dto = InvitationRead.model_validate(invitation)

    assert read_dto.id == invitation.id
    assert read_dto.role == invitation.role
    assert read_dto.department_id == invitation.department_id
    assert read_dto.doctor_id == invitation.doctor_id
    assert read_dto.created_by_user_id == invitation.created_by_user_id
    assert read_dto.expires_at == invitation.expires_at
    assert read_dto.used_at == invitation.used_at
    assert read_dto.revoked_at == invitation.revoked_at
    assert read_dto.is_deleted == invitation.is_deleted
    assert read_dto.sync_status == invitation.sync_status

    # Crucial security guarantee: token_hash and raw_token MUST NOT exist in read DTO
    dumped = read_dto.model_dump()
    assert "token_hash" not in dumped
    assert "raw_token" not in dumped
