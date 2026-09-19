"""
Integration tests for invitation lifecycle, department provisioning, and bootstrap security audit events.

Validates that:
- Staff invitation creation emits audit events with creator id, target role, and department id.
- Department admin invitation creation emits audit events with creator id and department id.
- Invitation revocation emits audit events with revoking admin id and department id.
- Invitation signup/consumption emits audit events with the new user id, role, and department id.
- Department provisioning emits audit events with superadmin id and new department id.
- Super-admin bootstrap emits audit events with user id and role.
- Zero credential/token leakage: raw invitation tokens and passwords never appear in audit output.
"""

import io
import json
import logging
import pytest
from sqlmodel import Session

from src.utils.logger import JsonFormatter
from src.user.models import UserRole
from src.auth.bootstrap import create_super_admin


@pytest.fixture(name="audit_logs")
def audit_logs_fixture():
    """Captures structured JSON logs from the src.security.audit logger."""
    stream = io.StringIO()
    handler = logging.StreamHandler(stream)
    handler.setFormatter(JsonFormatter())
    handler.setLevel(logging.DEBUG)

    logger = logging.getLogger("src.security.audit")
    logger.setLevel(logging.DEBUG)
    logger.addHandler(handler)
    logger.propagate = False

    try:
        yield stream
    finally:
        logger.removeHandler(handler)


def get_audit_records(stream: io.StringIO) -> list[dict]:
    """Helper to parse all captured JSON lines into dictionaries."""
    lines = [line.strip() for line in stream.getvalue().splitlines() if line.strip()]
    records = []
    for line in lines:
        try:
            records.append(json.loads(line))
        except json.JSONDecodeError:
            pass
    return records


# --------------------------------------------------------------------------
# 1. Staff & Department Admin Invitation Creation
# --------------------------------------------------------------------------

def test_staff_invitation_creation_emits_audit_log(
    client, department_factory, user_factory, auth_headers_factory, audit_logs
):
    """Department admin issuing a staff invitation emits an audit event without leaking token."""
    dept = department_factory()
    admin = user_factory(role=UserRole.DEPARTMENT_ADMIN, department_id=dept.id)
    headers = auth_headers_factory(admin)

    payload = {"role": "doctor", "doctor_id": None}
    resp = client.post("/api/v1/auth/invitations/staff", json=payload, headers=headers)
    assert resp.status_code == 201
    raw_token = resp.json()["raw_token"]

    records = get_audit_records(audit_logs)
    events = [r for r in records if r.get("action") == "invitation.create"]
    assert len(events) == 1, "Expected exactly one invitation.create audit event"

    event = events[0]
    assert event["event"] == "security.audit"
    assert event["outcome"] == "success"
    assert event["user_id"] == str(admin.id)
    assert event["role"] == "doctor"
    assert event["department_id"] == str(dept.id)

    # Zero leakage check: raw token must not appear in any log field
    log_dump = audit_logs.getvalue()
    assert raw_token not in log_dump


def test_department_admin_invitation_creation_emits_audit_log(
    client, department_factory, user_factory, auth_headers_factory, audit_logs
):
    """Super admin issuing a department admin invitation emits an audit event."""
    dept = department_factory()
    super_admin = user_factory(role=UserRole.SUPER_ADMIN, department_id=None)
    headers = auth_headers_factory(super_admin)

    payload = {
        "department_id": str(dept.id),
    }
    resp = client.post(
        "/api/v1/auth/invitations/admin", json=payload, headers=headers
    )
    assert resp.status_code == 201
    raw_token = resp.json()["raw_token"]

    records = get_audit_records(audit_logs)
    events = [r for r in records if r.get("action") == "invitation.create"]
    assert len(events) == 1

    event = events[0]
    assert event["event"] == "security.audit"
    assert event["outcome"] == "success"
    assert event["user_id"] == str(super_admin.id)
    assert event["role"] == "department_admin"
    assert event["department_id"] == str(dept.id)

    assert raw_token not in audit_logs.getvalue()


# --------------------------------------------------------------------------
# 2. Invitation Revocation
# --------------------------------------------------------------------------

def test_invitation_revocation_emits_audit_log(
    client, department_factory, user_factory, auth_headers_factory, audit_logs
):
    """Revoking an active invitation emits an audit event with admin and department id."""
    dept = department_factory()
    admin = user_factory(role=UserRole.DEPARTMENT_ADMIN, department_id=dept.id)
    headers = auth_headers_factory(admin)

    # Create an invitation first
    create_resp = client.post(
        "/api/v1/auth/invitations/staff",
        json={"role": "viewer"},
        headers=headers,
    )
    assert create_resp.status_code == 201
    invitation_id = create_resp.json()["id"]

    # Revoke it
    revoke_resp = client.post(
        f"/api/v1/auth/invitations/{invitation_id}/revoke", headers=headers
    )
    assert revoke_resp.status_code == 200

    records = get_audit_records(audit_logs)
    revoke_events = [r for r in records if r.get("action") == "invitation.revoke"]
    assert len(revoke_events) == 1, "Expected one invitation.revoke audit event"

    event = revoke_events[0]
    assert event["event"] == "security.audit"
    assert event["outcome"] == "success"
    assert event["user_id"] == str(admin.id)
    assert event["department_id"] == str(dept.id)


# --------------------------------------------------------------------------
# 3. Invitation Consumption (Signup)
# --------------------------------------------------------------------------

def test_invitation_signup_emits_audit_log(
    client, department_factory, user_factory, auth_headers_factory, audit_logs
):
    """Consuming an invitation to register emits an audit event without leaking the password."""
    dept = department_factory()
    admin = user_factory(role=UserRole.DEPARTMENT_ADMIN, department_id=dept.id)
    headers = auth_headers_factory(admin)

    create_resp = client.post(
        "/api/v1/auth/invitations/staff",
        json={"role": "viewer"},
        headers=headers,
    )
    assert create_resp.status_code == 201
    raw_token = create_resp.json()["raw_token"]

    signup_password = "SecurePassword123!"
    signup_payload = {
        "invitation_token": raw_token,
        "first_name": "New",
        "last_name": "Viewer",
        "email": "new_viewer@hospital.org",
        "password": signup_password,
    }

    signup_resp = client.post("/api/v1/auth/signup", json=signup_payload)
    assert signup_resp.status_code == 201
    new_user_data = signup_resp.json()

    records = get_audit_records(audit_logs)
    consume_events = [r for r in records if r.get("action") == "invitation.consume"]
    assert len(consume_events) == 1, "Expected one invitation.consume audit event"

    event = consume_events[0]
    assert event["event"] == "security.audit"
    assert event["outcome"] == "success"
    assert event["user_id"] == new_user_data["id"]
    assert event["role"] == "viewer"
    assert event["department_id"] == str(dept.id)

    # Zero leakage checks
    log_dump = audit_logs.getvalue()
    assert signup_password not in log_dump
    assert raw_token not in log_dump


# --------------------------------------------------------------------------
# 4. Department Provisioning
# --------------------------------------------------------------------------

def test_department_provisioning_emits_audit_log(
    client, user_factory, auth_headers_factory, audit_logs
):
    """Super admin provisioning a department emits an audit event."""
    super_admin = user_factory(role=UserRole.SUPER_ADMIN, department_id=None)
    headers = auth_headers_factory(super_admin)

    payload = {
        "department_name": "Radiology Department",
        "department_code": "RAD",
    }
    resp = client.post(
        "/api/v1/auth/invitations/provision", json=payload, headers=headers
    )
    assert resp.status_code == 201
    dept_id = resp.json()["department"]["id"]
    raw_token = resp.json()["invitation"]["raw_token"]

    records = get_audit_records(audit_logs)
    prov_events = [r for r in records if r.get("action") == "department.provision"]
    assert len(prov_events) == 1, "Expected one department.provision audit event"

    event = prov_events[0]
    assert event["event"] == "security.audit"
    assert event["outcome"] == "success"
    assert event["user_id"] == str(super_admin.id)
    assert event["department_id"] == str(dept_id)

    assert raw_token not in audit_logs.getvalue()


# --------------------------------------------------------------------------
# 5. Super-Admin Bootstrap
# --------------------------------------------------------------------------

def test_super_admin_bootstrap_emits_audit_log(session: Session, audit_logs):
    """Bootstrapping a super-admin account emits an audit event."""
    password = "BootstrapPassword123!"
    admin = create_super_admin(
        session=session,
        email="root_superadmin@hospital.org",
        full_name="Root Super Admin",
        password=password,
    )

    records = get_audit_records(audit_logs)
    bootstrap_events = [r for r in records if r.get("action") == "super_admin.bootstrap"]
    assert len(bootstrap_events) == 1, "Expected one super_admin.bootstrap audit event"

    event = bootstrap_events[0]
    assert event["event"] == "security.audit"
    assert event["outcome"] == "success"
    assert event["user_id"] == str(admin.id)
    assert event["role"] == "super_admin"

    assert password not in audit_logs.getvalue()
