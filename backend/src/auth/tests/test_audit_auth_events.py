"""
Integration tests for authentication and session lifecycle security audit events.

Validates that:
- Successful logins emit audit events with user identity and role.
- Failed logins emit audit events with outcome='failure' and reason='invalid_credentials'.
- Token rotations emit audit events with outcome='success'.
- Token reuse detections emit high-severity audit events with level='ERROR' and reason='reuse_detected'.
- Logouts emit audit events with outcome='success'.
- Zero credential leakage: passwords, raw tokens, and hashes never appear in audit output.
"""

import io
import json
import logging
import pytest
from sqlmodel import Session

from src.utils.logger import JsonFormatter
from src.user.models import UserRole

LOGIN_EMAIL = "audit_doctor@hospital.org"
LOGIN_PASSWORD = "Password123!"


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


@pytest.fixture(name="test_user")
def test_user_fixture(user_factory):
    """Creates an active test user for login and session audits."""
    return user_factory(
        role=UserRole.SUPER_ADMIN,
        email=LOGIN_EMAIL,
        password=LOGIN_PASSWORD,
    )


# --------------------------------------------------------------------------
# 1. Login Audit Tests
# --------------------------------------------------------------------------

def test_login_success_emits_audit_log(client, test_user, audit_logs):
    """Successful login emits an audit event with user identity and role."""
    resp = client.post(
        "/api/v1/auth/login",
        data={"username": test_user.email, "password": LOGIN_PASSWORD},
    )
    assert resp.status_code == 200

    records = get_audit_records(audit_logs)
    login_events = [r for r in records if r.get("action") == "auth.login"]
    assert len(login_events) >= 1, "Expected auth.login audit event"

    event = login_events[0]
    assert event["event"] == "security.audit"
    assert event["outcome"] == "success"
    assert event["user_id"] == str(test_user.id)
    assert event["role"] == "super_admin"
    assert event["level"] == "INFO"


def test_login_failure_emits_audit_log(client, test_user, audit_logs):
    """Failed login emits a warning audit event without revealing password."""
    resp = client.post(
        "/api/v1/auth/login",
        data={"username": test_user.email, "password": "WrongPassword999!"},
    )
    assert resp.status_code == 401

    records = get_audit_records(audit_logs)
    failed_events = [
        r for r in records
        if r.get("action") == "auth.login" and r.get("outcome") == "failure"
    ]
    assert len(failed_events) >= 1, "Expected failed auth.login audit event"

    event = failed_events[0]
    assert event["event"] == "security.audit"
    assert event["reason"] == "invalid_credentials"
    assert event["level"] == "WARNING"
    assert "WrongPassword999!" not in audit_logs.getvalue()


# --------------------------------------------------------------------------
# 2. Token Refresh Audit Tests
# --------------------------------------------------------------------------

def test_refresh_success_emits_audit_log(client, test_user, audit_logs):
    """Successful token rotation emits an auth.refresh audit event."""
    login_resp = client.post(
        "/api/v1/auth/login",
        data={"username": test_user.email, "password": LOGIN_PASSWORD},
    )
    raw_refresh = login_resp.json()["refresh_token"]

    # Clear capture stream so we only inspect the refresh call
    audit_logs.truncate(0)
    audit_logs.seek(0)
    client.cookies.clear()

    refresh_resp = client.post(
        "/api/v1/auth/refresh",
        json={"refresh_token": raw_refresh},
    )
    assert refresh_resp.status_code == 200

    records = get_audit_records(audit_logs)
    refresh_events = [r for r in records if r.get("action") == "auth.refresh"]
    assert len(refresh_events) >= 1, "Expected auth.refresh audit event"

    event = refresh_events[0]
    assert event["event"] == "security.audit"
    assert event["outcome"] == "success"
    assert event["user_id"] == str(test_user.id)
    assert event["level"] == "INFO"


def test_refresh_reuse_detection_emits_audit_error(client, test_user, audit_logs):
    """Token reuse alarm emits an ERROR-level security audit event."""
    login_resp = client.post(
        "/api/v1/auth/login",
        data={"username": test_user.email, "password": LOGIN_PASSWORD},
    )
    t1 = login_resp.json()["refresh_token"]
    client.cookies.clear()

    # Legitimate rotation T1 -> T2
    rotate_resp = client.post("/api/v1/auth/refresh", json={"refresh_token": t1})
    assert rotate_resp.status_code == 200

    # Clear logs to isolate the attack simulation
    audit_logs.truncate(0)
    audit_logs.seek(0)

    # Attacker replays stale T1
    replay_resp = client.post("/api/v1/auth/refresh", json={"refresh_token": t1})
    assert replay_resp.status_code == 401

    records = get_audit_records(audit_logs)
    reuse_events = [
        r for r in records
        if r.get("action") == "auth.refresh" and r.get("reason") == "reuse_detected"
    ]
    assert len(reuse_events) >= 1, "Expected reuse_detected audit error event"

    event = reuse_events[0]
    assert event["outcome"] == "failure"
    assert event["level"] == "ERROR"


# --------------------------------------------------------------------------
# 3. Logout Audit Tests
# --------------------------------------------------------------------------

def test_logout_emits_audit_log(client, test_user, audit_logs):
    """User logout emits an auth.logout audit event."""
    login_resp = client.post(
        "/api/v1/auth/login",
        data={"username": test_user.email, "password": LOGIN_PASSWORD},
    )
    t1 = login_resp.json()["refresh_token"]

    audit_logs.truncate(0)
    audit_logs.seek(0)

    logout_resp = client.post("/api/v1/auth/logout", json={"refresh_token": t1})
    assert logout_resp.status_code == 200

    records = get_audit_records(audit_logs)
    logout_events = [r for r in records if r.get("action") == "auth.logout"]
    assert len(logout_events) >= 1, "Expected auth.logout audit event"

    event = logout_events[0]
    assert event["event"] == "security.audit"
    assert event["outcome"] == "success"


# --------------------------------------------------------------------------
# 4. Zero-Leakage Assertion
# --------------------------------------------------------------------------

def test_zero_credential_leakage_in_audit_records(client, test_user, audit_logs):
    """Audit logs must never contain raw passwords, refresh tokens, or hashes."""
    login_resp = client.post(
        "/api/v1/auth/login",
        data={"username": test_user.email, "password": LOGIN_PASSWORD},
    )
    tokens = login_resp.json()
    raw_refresh = tokens["refresh_token"]
    raw_csrf = tokens["csrf_token"]

    # Refresh
    client.cookies.clear()
    refresh_resp = client.post("/api/v1/auth/refresh", json={"refresh_token": raw_refresh})
    assert refresh_resp.status_code == 200
    new_refresh = refresh_resp.json()["refresh_token"]

    # Logout
    client.post("/api/v1/auth/logout", json={"refresh_token": new_refresh})

    all_logs = audit_logs.getvalue()

    # Assert zero leakage of secrets
    assert LOGIN_PASSWORD not in all_logs
    assert raw_refresh not in all_logs
    assert raw_csrf not in all_logs
    assert new_refresh not in all_logs
