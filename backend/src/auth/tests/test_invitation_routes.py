"""
Integration and route tests for invitation management and issuing endpoints.
"""

import uuid
import datetime
import pytest

from src.user.models import UserRole
from src.auth import repository as auth_repository


# ==============================================================================
# POST /api/v1/auth/invitations/staff (Staff Invitation)
# ==============================================================================

def test_create_staff_invitation_endpoint_success(
    client, department_factory, user_factory, auth_headers_factory
):
    """Department admin successfully issues an invitation for a doctor."""
    dept = department_factory()
    admin = user_factory(role=UserRole.DEPARTMENT_ADMIN, department_id=dept.id)
    headers = auth_headers_factory(admin)

    payload = {
        "role": "doctor",
        "doctor_id": None,
    }

    response = client.post("/api/v1/auth/invitations/staff", json=payload, headers=headers)
    assert response.status_code == 201

    data = response.json()
    assert data["role"] == "doctor"
    assert data["department_id"] == str(dept.id)
    assert data["doctor_id"] is None
    assert "raw_token" in data
    assert len(data["raw_token"]) > 20
    assert "id" in data


def test_create_staff_invitation_endpoint_viewer(
    client, department_factory, user_factory, auth_headers_factory
):
    """Department admin successfully issues an invitation for a viewer."""
    dept = department_factory()
    admin = user_factory(role=UserRole.DEPARTMENT_ADMIN, department_id=dept.id)
    headers = auth_headers_factory(admin)

    payload = {"role": "viewer"}
    response = client.post("/api/v1/auth/invitations/staff", json=payload, headers=headers)
    assert response.status_code == 201

    data = response.json()
    assert data["role"] == "viewer"
    assert data["department_id"] == str(dept.id)


def test_create_staff_invitation_endpoint_forbidden_roles(
    client, department_factory, user_factory, doctor_factory, auth_headers_factory
):
    """Doctors, viewers, and super admins cannot issue staff invitations via this endpoint."""
    dept = department_factory()
    doc = doctor_factory(department_id=dept.id)
    doctor_user = user_factory(role=UserRole.DOCTOR, department_id=dept.id, doctor_id=doc.id)
    viewer_user = user_factory(role=UserRole.VIEWER, department_id=dept.id)
    super_admin = user_factory(role=UserRole.SUPER_ADMIN, department_id=None)

    payload = {"role": "doctor"}

    for unauthorized_user in (doctor_user, viewer_user, super_admin):
        headers = auth_headers_factory(unauthorized_user)
        response = client.post("/api/v1/auth/invitations/staff", json=payload, headers=headers)
        assert response.status_code == 403


# ==============================================================================
# POST /api/v1/auth/invitations/provision (Atomic Provisioning)
# ==============================================================================

def test_provision_department_endpoint_success(
    client, user_factory, auth_headers_factory
):
    """Super admin successfully provisions a new department and receives the admin invitation."""
    super_admin = user_factory(role=UserRole.SUPER_ADMIN, department_id=None)
    headers = auth_headers_factory(super_admin)

    payload = {
        "department_name": "Emergency Medicine",
        "department_code": "ER",
    }

    response = client.post("/api/v1/auth/invitations/provision", json=payload, headers=headers)
    assert response.status_code == 201

    data = response.json()
    assert "department" in data
    assert data["department"]["name"] == "Emergency Medicine"
    assert data["department"]["code"] == "ER"

    assert "invitation" in data
    assert data["invitation"]["role"] == "department_admin"
    assert "raw_token" in data["invitation"]
    assert data["invitation"]["department_id"] == data["department"]["id"]


def test_provision_department_endpoint_forbidden_for_non_super_admin(
    client, department_factory, user_factory, auth_headers_factory
):
    """Department admins cannot provision new departments."""
    dept = department_factory()
    admin = user_factory(role=UserRole.DEPARTMENT_ADMIN, department_id=dept.id)
    headers = auth_headers_factory(admin)

    payload = {
        "department_name": "Radiology",
        "department_code": "RAD",
    }

    response = client.post("/api/v1/auth/invitations/provision", json=payload, headers=headers)
    assert response.status_code == 403


def test_provision_department_endpoint_duplicate_name_rejected(
    client, department_factory, user_factory, auth_headers_factory
):
    """Provisioning with an existing department name returns HTTP 400."""
    dept = department_factory(name="Dermatology")
    super_admin = user_factory(role=UserRole.SUPER_ADMIN, department_id=None)
    headers = auth_headers_factory(super_admin)

    payload = {
        "department_name": dept.name,
        "department_code": "DERM2",
    }

    response = client.post("/api/v1/auth/invitations/provision", json=payload, headers=headers)
    assert response.status_code == 400


# ==============================================================================
# POST /api/v1/auth/invitations/admin (Admin Invitation for Existing Department)
# ==============================================================================

def test_create_department_admin_invitation_endpoint_success(
    client, department_factory, user_factory, auth_headers_factory
):
    """Super admin successfully generates an admin invitation for an existing department."""
    dept = department_factory()
    super_admin = user_factory(role=UserRole.SUPER_ADMIN, department_id=None)
    headers = auth_headers_factory(super_admin)

    payload = {"department_id": str(dept.id)}

    response = client.post("/api/v1/auth/invitations/admin", json=payload, headers=headers)
    assert response.status_code == 201

    data = response.json()
    assert data["role"] == "department_admin"
    assert data["department_id"] == str(dept.id)
    assert "raw_token" in data


def test_create_department_admin_invitation_endpoint_nonexistent_department(
    client, user_factory, auth_headers_factory
):
    """Super admin targeting a nonexistent department receives HTTP 404."""
    super_admin = user_factory(role=UserRole.SUPER_ADMIN, department_id=None)
    headers = auth_headers_factory(super_admin)

    payload = {"department_id": str(uuid.uuid4())}
    response = client.post("/api/v1/auth/invitations/admin", json=payload, headers=headers)
    assert response.status_code == 404


# ==============================================================================
# GET /api/v1/auth/invitations (List Invitations)
# ==============================================================================

def test_list_invitations_department_admin_scoped(
    client, department_factory, user_factory, invitation_factory, auth_headers_factory
):
    """Department admin sees only their own department's active invitations."""
    dept_a = department_factory()
    dept_b = department_factory()

    admin_a = user_factory(role=UserRole.DEPARTMENT_ADMIN, department_id=dept_a.id)
    headers_a = auth_headers_factory(admin_a)

    inv_a1 = invitation_factory(department_id=dept_a.id, role=UserRole.DOCTOR)
    inv_a2 = invitation_factory(department_id=dept_a.id, role=UserRole.VIEWER)
    inv_b = invitation_factory(department_id=dept_b.id, role=UserRole.DOCTOR)
    inv_deleted = invitation_factory(department_id=dept_a.id, is_deleted=True)

    response = client.get("/api/v1/auth/invitations", headers=headers_a)
    assert response.status_code == 200

    items = response.json()
    item_ids = {item["id"] for item in items}

    assert str(inv_a1.id) in item_ids
    assert str(inv_a2.id) in item_ids
    assert str(inv_b.id) not in item_ids
    assert str(inv_deleted.id) not in item_ids

    # Security check: secret credentials must NEVER be in list response
    for item in items:
        assert "raw_token" not in item
        assert "token_hash" not in item


def test_list_invitations_super_admin(
    client, department_factory, user_factory, invitation_factory, auth_headers_factory
):
    """Super admin sees all department admin invitations."""
    super_admin = user_factory(role=UserRole.SUPER_ADMIN, department_id=None)
    headers = auth_headers_factory(super_admin)

    admin_inv1 = invitation_factory(role=UserRole.DEPARTMENT_ADMIN)
    admin_inv2 = invitation_factory(role=UserRole.DEPARTMENT_ADMIN)
    doctor_inv = invitation_factory(role=UserRole.DOCTOR)

    response = client.get("/api/v1/auth/invitations", headers=headers)
    assert response.status_code == 200

    item_ids = {item["id"] for item in response.json()}
    assert str(admin_inv1.id) in item_ids
    assert str(admin_inv2.id) in item_ids
    assert str(doctor_inv.id) not in item_ids


def test_list_invitations_forbidden_for_doctor_and_viewer(
    client, department_factory, user_factory, doctor_factory, auth_headers_factory
):
    """Doctors and viewers cannot access the invitation listing endpoint."""
    dept = department_factory()
    doc = doctor_factory(department_id=dept.id)
    doctor_user = user_factory(role=UserRole.DOCTOR, department_id=dept.id, doctor_id=doc.id)
    viewer_user = user_factory(role=UserRole.VIEWER, department_id=dept.id)

    for unauthorized_user in (doctor_user, viewer_user):
        headers = auth_headers_factory(unauthorized_user)
        response = client.get("/api/v1/auth/invitations", headers=headers)
        assert response.status_code == 403


def test_list_invitations_unauthenticated(client):
    """Unauthenticated access to invitations returns HTTP 401."""
    response = client.get("/api/v1/auth/invitations")
    assert response.status_code == 401


# ==============================================================================
# POST /api/v1/auth/invitations/{invitation_id}/revoke (Revoke Invitation)
# ==============================================================================

def test_revoke_invitation_department_admin_success(
    client, session, department_factory, user_factory, invitation_factory, auth_headers_factory
):
    """Department admin successfully revokes an invitation in their department."""
    dept = department_factory()
    admin = user_factory(role=UserRole.DEPARTMENT_ADMIN, department_id=dept.id)
    headers = auth_headers_factory(admin)

    invitation = invitation_factory(department_id=dept.id)
    assert invitation.revoked_at is None

    response = client.post(f"/api/v1/auth/invitations/{invitation.id}/revoke", headers=headers)
    assert response.status_code == 200

    data = response.json()
    assert data["id"] == str(invitation.id)
    assert data["revoked_at"] is not None

    # Confirm in DB
    refreshed = auth_repository.get_invitation_by_id(session, invitation.id)
    assert refreshed.revoked_at is not None
    assert refreshed.is_active is False


def test_revoke_invitation_cross_department_not_found(
    client, department_factory, user_factory, invitation_factory, auth_headers_factory
):
    """Department admin cannot revoke an invitation belonging to another department (returns 404)."""
    dept_a = department_factory()
    dept_b = department_factory()

    admin_a = user_factory(role=UserRole.DEPARTMENT_ADMIN, department_id=dept_a.id)
    headers_a = auth_headers_factory(admin_a)

    invitation_b = invitation_factory(department_id=dept_b.id)

    response = client.post(f"/api/v1/auth/invitations/{invitation_b.id}/revoke", headers=headers_a)
    assert response.status_code == 404


def test_revoke_invitation_super_admin_success(
    client, session, department_factory, user_factory, invitation_factory, auth_headers_factory
):
    """Super admin can revoke an invitation across departments."""
    dept = department_factory()
    super_admin = user_factory(role=UserRole.SUPER_ADMIN, department_id=None)
    headers = auth_headers_factory(super_admin)

    invitation = invitation_factory(department_id=dept.id, role=UserRole.DEPARTMENT_ADMIN)

    response = client.post(f"/api/v1/auth/invitations/{invitation.id}/revoke", headers=headers)
    assert response.status_code == 200
    assert response.json()["revoked_at"] is not None


def test_revoke_invitation_nonexistent(
    client, department_factory, user_factory, auth_headers_factory
):
    """Revoking a nonexistent invitation returns HTTP 404."""
    dept = department_factory()
    admin = user_factory(role=UserRole.DEPARTMENT_ADMIN, department_id=dept.id)
    headers = auth_headers_factory(admin)

    response = client.post(f"/api/v1/auth/invitations/{uuid.uuid4()}/revoke", headers=headers)
    assert response.status_code == 404


def test_revoke_invitation_forbidden_for_doctor_and_viewer(
    client, department_factory, user_factory, doctor_factory, invitation_factory, auth_headers_factory
):
    """Doctors and viewers cannot revoke invitations."""
    dept = department_factory()
    doc = doctor_factory(department_id=dept.id)
    doctor_user = user_factory(role=UserRole.DOCTOR, department_id=dept.id, doctor_id=doc.id)
    viewer_user = user_factory(role=UserRole.VIEWER, department_id=dept.id)
    invitation = invitation_factory(department_id=dept.id)

    for unauthorized_user in (doctor_user, viewer_user):
        headers = auth_headers_factory(unauthorized_user)
        response = client.post(f"/api/v1/auth/invitations/{invitation.id}/revoke", headers=headers)
        assert response.status_code == 403


def test_revoke_invitation_unauthenticated(client, invitation_factory):
    """Unauthenticated request to revoke returns HTTP 401."""
    invitation = invitation_factory()
    response = client.post(f"/api/v1/auth/invitations/{invitation.id}/revoke")
    assert response.status_code == 401
