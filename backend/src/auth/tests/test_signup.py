"""
Integration and route tests for public invitation consumption and account signup (POST /api/v1/auth/signup).
"""

import uuid
import datetime

from sqlmodel import Session, select
from src.user.models import UserRole
from src.auth.security import generate_invitation_token, hash_invitation_token
from src.auth import repository as auth_repository
from src.doctor.models import Doctor as DoctorModel


# ==============================================================================
# Successful Signup Scenarios
# ==============================================================================

def test_signup_doctor_with_auto_provisioning_success(
    client, session: Session, department_factory, invitation_factory
):
    """
    Invited DOCTOR without a pre-assigned doctor_id signs up:
    - Automatically provisions a new Doctor record matching full_name and department_id
    - Links user.doctor_id to the new doctor
    - Marks invitation as used_at = utcnow()
    - Returns 201 Created with UserRead
    """
    dept = department_factory()
    raw_token = generate_invitation_token()
    token_hash = hash_invitation_token(raw_token)

    invitation = invitation_factory(
        department_id=dept.id,
        role=UserRole.DOCTOR,
        doctor_id=None,
        token_hash=token_hash,
    )
    assert invitation.used_at is None

    payload = {
        "invitation_token": raw_token,
        "first_name": "Gregory",
        "last_name": "House",
        "email": "house@princeton-plainsboro.org",
        "password": "SecurePassword123!",
    }

    response = client.post("/api/v1/auth/signup", json=payload)
    assert response.status_code == 201

    data = response.json()
    assert data["role"] == "doctor"
    assert data["email"] == "house@princeton-plainsboro.org"
    assert data["full_name"] == "Gregory House"
    assert data["department_id"] == str(dept.id)
    assert data["doctor_id"] is not None

    # Verify auto-provisioned Doctor in DB
    created_doctor = session.get(DoctorModel, uuid.UUID(data["doctor_id"]))
    assert created_doctor is not None
    assert created_doctor.name == "Gregory House"
    assert created_doctor.department_id == dept.id
    assert created_doctor.team_id is None
    assert created_doctor.is_deleted is False

    # Verify Invitation was marked used
    refreshed_inv = auth_repository.get_invitation_by_id(session, invitation.id)
    assert refreshed_inv.used_at is not None
    assert refreshed_inv.is_active is False


def test_signup_doctor_with_pre_assigned_doctor_success(
    client, session: Session, department_factory, doctor_factory, invitation_factory
):
    """
    Invited DOCTOR with a pre-assigned doctor_id signs up:
    - Reuses existing doctor record (does NOT create duplicate Doctor)
    - Links user.doctor_id to the existing doctor
    - Marks invitation as used
    """
    dept = department_factory()
    existing_doc = doctor_factory(department_id=dept.id, full_name="James Wilson")
    raw_token = generate_invitation_token()
    token_hash = hash_invitation_token(raw_token)

    invitation = invitation_factory(
        department_id=dept.id,
        role=UserRole.DOCTOR,
        doctor_id=existing_doc.id,
        token_hash=token_hash,
    )

    payload = {
        "invitation_token": raw_token,
        "first_name": "James",
        "last_name": "Wilson",
        "email": "wilson@oncology.org",
        "password": "SecurePassword123!",
    }

    response = client.post("/api/v1/auth/signup", json=payload)
    assert response.status_code == 201

    data = response.json()
    assert data["role"] == "doctor"
    assert data["doctor_id"] == str(existing_doc.id)

    # Verify no new doctor was created
    all_docs = session.exec(select(DoctorModel).where(DoctorModel.department_id == dept.id)).all()
    assert len(all_docs) == 1
    assert all_docs[0].id == existing_doc.id

    # Verify Invitation marked used
    refreshed_inv = auth_repository.get_invitation_by_id(session, invitation.id)
    assert refreshed_inv.used_at is not None


def test_signup_viewer_success(
    client, session: Session, department_factory, invitation_factory
):
    """
    Invited VIEWER signs up:
    - User created with role VIEWER and doctor_id = None
    - No doctor record created
    - Marks invitation as used
    """
    dept = department_factory()
    raw_token = generate_invitation_token()
    token_hash = hash_invitation_token(raw_token)

    invitation = invitation_factory(
        department_id=dept.id,
        role=UserRole.VIEWER,
        doctor_id=None,
        token_hash=token_hash,
    )

    payload = {
        "invitation_token": raw_token,
        "first_name": "Audrey",
        "last_name": "Viewer",
        "email": "audrey@hospital.org",
        "password": "SecurePassword123!",
    }

    response = client.post("/api/v1/auth/signup", json=payload)
    assert response.status_code == 201

    data = response.json()
    assert data["role"] == "viewer"
    assert data["doctor_id"] is None
    assert data["department_id"] == str(dept.id)

    # Verify no doctor was created
    doc_count = len(session.exec(select(DoctorModel).where(DoctorModel.department_id == dept.id)).all())
    assert doc_count == 0

    refreshed_inv = auth_repository.get_invitation_by_id(session, invitation.id)
    assert refreshed_inv.used_at is not None


def test_signup_department_admin_success(
    client, session: Session, department_factory, invitation_factory
):
    """
    Invited DEPARTMENT_ADMIN signs up:
    - User created with role DEPARTMENT_ADMIN and doctor_id = None
    - Marks invitation as used
    """
    dept = department_factory()
    raw_token = generate_invitation_token()
    token_hash = hash_invitation_token(raw_token)

    invitation = invitation_factory(
        department_id=dept.id,
        role=UserRole.DEPARTMENT_ADMIN,
        doctor_id=None,
        token_hash=token_hash,
    )

    payload = {
        "invitation_token": raw_token,
        "first_name": "Lisa",
        "last_name": "Cuddy",
        "email": "cuddy@hospital.org",
        "password": "SecurePassword123!",
    }

    response = client.post("/api/v1/auth/signup", json=payload)
    assert response.status_code == 201

    data = response.json()
    assert data["role"] == "department_admin"
    assert data["department_id"] == str(dept.id)
    assert data["doctor_id"] is None

    refreshed_inv = auth_repository.get_invitation_by_id(session, invitation.id)
    assert refreshed_inv.used_at is not None


# ==============================================================================
# Token Invalidation & Security Rejection Scenarios
# ==============================================================================

def test_signup_rejects_already_used_invitation(client, invitation_factory):
    """Attempting to use an invitation that has already been consumed returns HTTP 400."""
    raw_token = generate_invitation_token()
    token_hash = hash_invitation_token(raw_token)

    invitation_factory(
        token_hash=token_hash,
        used_at=datetime.datetime.now(datetime.timezone.utc),
    )

    payload = {
        "invitation_token": raw_token,
        "first_name": "Robert",
        "last_name": "Chase",
        "email": "chase@hospital.org",
        "password": "SecurePassword123!",
    }

    response = client.post("/api/v1/auth/signup", json=payload)
    assert response.status_code == 400
    assert "used" in response.json()["detail"].lower()


def test_signup_rejects_revoked_invitation(client, invitation_factory):
    """Attempting to use an invitation that was revoked returns HTTP 400."""
    raw_token = generate_invitation_token()
    token_hash = hash_invitation_token(raw_token)

    invitation_factory(
        token_hash=token_hash,
        revoked_at=datetime.datetime.now(datetime.timezone.utc),
    )

    payload = {
        "invitation_token": raw_token,
        "first_name": "Allison",
        "last_name": "Cameron",
        "email": "cameron@hospital.org",
        "password": "SecurePassword123!",
    }

    response = client.post("/api/v1/auth/signup", json=payload)
    assert response.status_code == 400
    assert "revoked" in response.json()["detail"].lower()


def test_signup_rejects_expired_invitation(client, invitation_factory):
    """Attempting to use an invitation past its expiration timestamp returns HTTP 400."""
    raw_token = generate_invitation_token()
    token_hash = hash_invitation_token(raw_token)

    invitation_factory(
        token_hash=token_hash,
        expires_at=datetime.datetime.now(datetime.timezone.utc) - datetime.timedelta(days=1),
    )

    payload = {
        "invitation_token": raw_token,
        "first_name": "Eric",
        "last_name": "Foreman",
        "email": "foreman@hospital.org",
        "password": "SecurePassword123!",
    }

    response = client.post("/api/v1/auth/signup", json=payload)
    assert response.status_code == 400
    assert "expired" in response.json()["detail"].lower()


def test_signup_rejects_nonexistent_invitation_token(client):
    """Submitting a token that does not exist in the system returns HTTP 404."""
    payload = {
        "invitation_token": "nonexistent_fake_bearer_token",
        "first_name": "Chris",
        "last_name": "Taub",
        "email": "taub@hospital.org",
        "password": "SecurePassword123!",
    }

    response = client.post("/api/v1/auth/signup", json=payload)
    assert response.status_code == 404


def test_signup_rejects_soft_deleted_invitation(client, invitation_factory):
    """Soft-deleted invitations cannot be redeemed and return HTTP 404."""
    raw_token = generate_invitation_token()
    token_hash = hash_invitation_token(raw_token)

    invitation_factory(
        token_hash=token_hash,
        is_deleted=True,
    )

    payload = {
        "invitation_token": raw_token,
        "first_name": "Remy",
        "last_name": "Hadley",
        "email": "thirteen@hospital.org",
        "password": "SecurePassword123!",
    }

    response = client.post("/api/v1/auth/signup", json=payload)
    assert response.status_code == 404


def test_signup_rejects_duplicate_user_email_and_preserves_invitation(
    client, session: Session, user_factory, invitation_factory
):
    """
    If the invitee's email already belongs to an active user:
    - Returns HTTP 409 Conflict
    - Atomically rolls back: invitation remains unconsumed (used_at is None)
    """
    existing_user = user_factory(
        role=UserRole.SUPER_ADMIN,
        email="existing@hospital.org",
    )
    raw_token = generate_invitation_token()
    token_hash = hash_invitation_token(raw_token)

    invitation = invitation_factory(token_hash=token_hash)
    assert invitation.used_at is None

    payload = {
        "invitation_token": raw_token,
        "first_name": "Lawrence",
        "last_name": "Kutner",
        "email": "existing@hospital.org",  # Collision!
        "password": "SecurePassword123!",
    }

    response = client.post("/api/v1/auth/signup", json=payload)
    assert response.status_code == 409

    # Verify invitation is NOT consumed
    refreshed_inv = auth_repository.get_invitation_by_id(session, invitation.id)
    assert refreshed_inv.used_at is None
    assert refreshed_inv.is_active is True


def test_signup_validates_request_body(client):
    """Missing or empty fields return HTTP 422 validation error."""
    payload = {
        "invitation_token": "",  # Empty
        "first_name": "Amber",
        "last_name": "Volakis",
        "email": "not-an-email",  # Invalid email
        "password": "",  # Empty password
    }

    response = client.post("/api/v1/auth/signup", json=payload)
    assert response.status_code == 422
