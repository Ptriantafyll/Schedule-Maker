"""
Integration tests for authorization and tenant denial security audit events.

Validates that:
- Insufficient role access (RBAC) emits an audit event with action='rbac.denied', outcome='failure', and reason='insufficient_role'.
- Missing or invalid tenant/department scope emits an audit event with action='tenant.denied', outcome='failure', and reason='missing_department_scope'.
- Both route-level requests and direct dependency invocations properly emit denial audit events.
"""

import io
import json
import logging
import pytest
from fastapi import HTTPException

from src.utils.logger import JsonFormatter
from src.user.models import UserRole
from src.auth import dependencies


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
# 1. RBAC Denial Audit Tests (Endpoint Integration)
# --------------------------------------------------------------------------

def test_rbac_denial_emits_audit_log_when_doctor_accesses_admin_route(
    client, department_factory, user_factory, doctor_factory, auth_headers_factory, audit_logs
):
    """A doctor attempting to call an admin-only endpoint triggers an rbac.denied audit event."""
    dept = department_factory()
    doc = doctor_factory(department_id=dept.id)
    doctor_user = user_factory(
        role=UserRole.DOCTOR, department_id=dept.id, doctor_id=doc.id
    )
    headers = auth_headers_factory(doctor_user)

    # Calling an endpoint guarded by require_department_admin
    resp = client.post(
        "/api/v1/auth/invitations/staff",
        json={"role": "doctor"},
        headers=headers,
    )
    assert resp.status_code == 403

    records = get_audit_records(audit_logs)
    rbac_events = [r for r in records if r.get("action") == "rbac.denied"]
    assert len(rbac_events) == 1, "Expected one rbac.denied audit event"

    event = rbac_events[0]
    assert event["event"] == "security.audit"
    assert event["outcome"] == "failure"
    assert event["user_id"] == str(doctor_user.id)
    assert event["role"] == "doctor"
    assert event["department_id"] == str(dept.id)
    assert event["reason"] == "insufficient_role"
    assert event["level"] == "WARNING"


def test_rbac_denial_emits_audit_log_when_admin_accesses_superadmin_route(
    client, department_factory, user_factory, auth_headers_factory, audit_logs
):
    """A department admin attempting to call a superadmin route triggers an rbac.denied audit event."""
    dept = department_factory()
    admin_user = user_factory(
        role=UserRole.DEPARTMENT_ADMIN, department_id=dept.id
    )
    headers = auth_headers_factory(admin_user)

    # Calling an endpoint guarded by require_super_admin
    resp = client.post(
        "/api/v1/auth/invitations/provision",
        json={"department_name": "New Dept", "department_code": "ND"},
        headers=headers,
    )
    assert resp.status_code == 403

    records = get_audit_records(audit_logs)
    rbac_events = [r for r in records if r.get("action") == "rbac.denied"]
    assert len(rbac_events) == 1

    event = rbac_events[0]
    assert event["event"] == "security.audit"
    assert event["outcome"] == "failure"
    assert event["user_id"] == str(admin_user.id)
    assert event["role"] == "department_admin"
    assert event["department_id"] == str(dept.id)
    assert event["reason"] == "insufficient_role"


# --------------------------------------------------------------------------
# 2. RBAC Denial Audit Tests (Direct Guard Unit Tests)
# --------------------------------------------------------------------------

def test_require_role_guard_direct_call_emits_audit_log(
    department_factory, user_factory, audit_logs
):
    """Directly invoking role_guard with unauthorized user raises 403 and audits denial."""
    dept = department_factory()
    viewer_user = user_factory(role=UserRole.VIEWER, department_id=dept.id)
    guard = dependencies.require_role(UserRole.SUPER_ADMIN, UserRole.DEPARTMENT_ADMIN)

    with pytest.raises(HTTPException) as exc_info:
        guard(current_user=viewer_user)

    assert exc_info.value.status_code == 403

    records = get_audit_records(audit_logs)
    rbac_events = [r for r in records if r.get("action") == "rbac.denied"]
    assert len(rbac_events) == 1

    event = rbac_events[0]
    assert event["event"] == "security.audit"
    assert event["outcome"] == "failure"
    assert event["user_id"] == str(viewer_user.id)
    assert event["role"] == "viewer"
    assert event["department_id"] == str(dept.id)
    assert event["reason"] == "insufficient_role"


# --------------------------------------------------------------------------
# 3. Tenant Scope Denial Audit Tests
# --------------------------------------------------------------------------

def test_require_department_scope_emits_audit_log_when_missing(
    user_factory, audit_logs
):
    """Users lacking department_id trigger a tenant.denied audit event when scope is required."""
    super_admin = user_factory(role=UserRole.SUPER_ADMIN, department_id=None)

    with pytest.raises(HTTPException) as exc_info:
        dependencies.require_department_scope(current_user=super_admin)

    assert exc_info.value.status_code == 403

    records = get_audit_records(audit_logs)
    tenant_events = [r for r in records if r.get("action") == "tenant.denied"]
    assert len(tenant_events) == 1, "Expected one tenant.denied audit event"

    event = tenant_events[0]
    assert event["event"] == "security.audit"
    assert event["outcome"] == "failure"
    assert event["user_id"] == str(super_admin.id)
    assert event["role"] == "super_admin"
    assert event["reason"] == "missing_department_scope"
    assert event["level"] == "WARNING"

