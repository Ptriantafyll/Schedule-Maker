"""
Unit tests for security audit logging and JsonFormatter field filtering.

Validates:
- log_audit_event emits structured JSON with approved fields
- UUIDs are converted to strings properly
- Severity levels (INFO, WARNING, ERROR) are respected
- Non-whitelisted fields (passwords, raw tokens, hashes) are strictly stripped by JsonFormatter
- Optional fields when None are omitted from the JSON payload
"""

import io
import json
import logging
import uuid
import pytest

from src.utils.logger import (
    JsonFormatter,
    log_audit_event,
    _EXTRA_FIELDS,
)


@pytest.fixture(name="audit_capture")
def audit_capture_fixture():
    """Captures log records emitted by the src.security.audit logger in memory."""
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


def test_log_audit_event_emits_structured_json(audit_capture):
    """Calling log_audit_event serializes an approved JSON security record."""
    user_id = uuid.uuid4()
    dept_id = uuid.uuid4()

    log_audit_event(
        action="auth.login",
        outcome="success",
        message="User login succeeded",
        user_id=user_id,
        role="doctor",
        department_id=dept_id,
        level=logging.INFO,
    )

    output = audit_capture.getvalue().strip()
    assert output, "No log output captured"
    data = json.loads(output)

    assert data["event"] == "security.audit"
    assert data["action"] == "auth.login"
    assert data["outcome"] == "success"
    assert data["message"] == "User login succeeded"
    assert data["user_id"] == str(user_id)
    assert data["role"] == "doctor"
    assert data["department_id"] == str(dept_id)
    assert data["level"] == "INFO"
    assert data["logger"] == "src.security.audit"
    assert "timestamp" in data


def test_log_audit_event_failure_with_reason(audit_capture):
    """Failure events record reason code and appropriate log level."""
    log_audit_event(
        action="auth.login",
        outcome="failure",
        message="Authentication failed for user",
        reason="invalid_credentials",
        level=logging.WARNING,
    )

    output = audit_capture.getvalue().strip()
    data = json.loads(output)

    assert data["event"] == "security.audit"
    assert data["action"] == "auth.login"
    assert data["outcome"] == "failure"
    assert data["reason"] == "invalid_credentials"
    assert data["level"] == "WARNING"


def test_log_audit_event_omits_none_fields(audit_capture):
    """Optional fields that are None are omitted from the JSON payload."""
    log_audit_event(
        action="auth.logout",
        outcome="success",
        message="Session terminated",
    )

    output = audit_capture.getvalue().strip()
    data = json.loads(output)

    assert data["action"] == "auth.logout"
    assert data["outcome"] == "success"
    assert "user_id" not in data
    assert "role" not in data
    assert "department_id" not in data
    assert "reason" not in data


def test_formatter_strips_unapproved_fields_zero_leakage():
    """JsonFormatter strictly drops any extra field not present in _EXTRA_FIELDS."""
    formatter = JsonFormatter()
    record = logging.LogRecord(
        name="src.security.audit",
        level=logging.INFO,
        pathname=__file__,
        lineno=10,
        msg="Test event",
        args=(),
        exc_info=None,
    )

    # Attach both approved and dangerous unapproved fields
    record.event = "security.audit"
    record.action = "auth.login"
    record.password = "PlaintextSuperSecretPassword123!"
    record.raw_refresh_token = "raw_unhashed_refresh_token_value"
    record.token_hash = "secret_sha256_digest"

    formatted_json = formatter.format(record)
    data = json.loads(formatted_json)

    # Approved fields should be kept
    assert data["event"] == "security.audit"
    assert data["action"] == "auth.login"

    # Dangerous unapproved fields MUST NOT be present
    assert "password" not in data
    assert "raw_refresh_token" not in data
    assert "token_hash" not in data
    assert "PlaintextSuperSecretPassword123!" not in formatted_json
    assert "raw_unhashed_refresh_token_value" not in formatted_json


def test_approved_audit_fields_are_in_extra_fields_allowlist():
    """Confirms all security audit fields are registered in _EXTRA_FIELDS."""
    required_audit_fields = {
        "action",
        "outcome",
        "user_id",
        "role",
        "department_id",
        "reason",
    }
    for field in required_audit_fields:
        assert field in _EXTRA_FIELDS, f"Field '{field}' missing from _EXTRA_FIELDS"
