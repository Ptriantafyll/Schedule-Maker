"""
Unit tests for authentication refresh session request and response schemas.
"""

import datetime
import pytest
from pydantic import ValidationError

from src.auth.schemas import Token, RefreshTokenRequest, RefreshSessionRead


def test_token_schema_backward_compatibility():
    """Token schema continues to work with only access_token provided."""
    token = Token(access_token="test_access_token")
    assert token.access_token == "test_access_token"
    assert token.token_type == "bearer"
    assert token.refresh_token is None
    assert token.csrf_token is None


def test_token_schema_with_refresh_and_csrf():
    """Token schema accepts optional refresh_token and csrf_token."""
    token = Token(
        access_token="access_123",
        refresh_token="refresh_456",
        csrf_token="csrf_789",
    )
    assert token.access_token == "access_123"
    assert token.token_type == "bearer"
    assert token.refresh_token == "refresh_456"
    assert token.csrf_token == "csrf_789"


def test_refresh_token_request_valid():
    """RefreshTokenRequest accepts valid tokens."""
    req = RefreshTokenRequest(
        refresh_token="ref_abc123",
        csrf_token="csrf_xyz456",
    )
    assert req.refresh_token == "ref_abc123"
    assert req.csrf_token == "csrf_xyz456"


def test_refresh_token_request_optional_fields():
    """RefreshTokenRequest fields default to None for cookie-based clients."""
    req = RefreshTokenRequest()
    assert req.refresh_token is None
    assert req.csrf_token is None


def test_refresh_token_request_strips_whitespace():
    """RefreshTokenRequest trims leading and trailing whitespace."""
    req = RefreshTokenRequest(
        refresh_token="  token_abc  ",
        csrf_token="  csrf_xyz  ",
    )
    assert req.refresh_token == "token_abc"
    assert req.csrf_token == "csrf_xyz"


def test_refresh_token_request_rejects_empty_or_whitespace_tokens():
    """RefreshTokenRequest raises ValidationError on empty or whitespace tokens."""
    with pytest.raises(ValidationError):
        RefreshTokenRequest(refresh_token="")

    with pytest.raises(ValidationError):
        RefreshTokenRequest(refresh_token="   ")

    with pytest.raises(ValidationError):
        RefreshTokenRequest(csrf_token="")

    with pytest.raises(ValidationError):
        RefreshTokenRequest(csrf_token="   ")


def test_refresh_token_request_extra_fields_forbidden():
    """Extra unexpected payload fields are rejected."""
    with pytest.raises(ValidationError):
        RefreshTokenRequest(extra_field="unexpected")  # pylint: disable=unexpected-keyword-arg


def test_refresh_session_read_from_model(refresh_session_factory):
    """RefreshSessionRead maps cleanly from a persisted RefreshSession instance."""
    session_obj = refresh_session_factory()

    read_dto = RefreshSessionRead.model_validate(session_obj)

    assert read_dto.id == session_obj.id
    assert read_dto.user_id == session_obj.user_id
    assert read_dto.session_family == session_obj.session_family
    assert read_dto.expires_at == session_obj.expires_at
    assert read_dto.last_used_at == session_obj.last_used_at
    assert read_dto.revoked_at is None
    assert read_dto.revoked_reason is None
    assert read_dto.replaced_by_session_id is None
    assert read_dto.is_deleted is False
    assert read_dto.is_active is True
    assert read_dto.is_expired is False


def test_refresh_session_read_omits_security_hashes(refresh_session_factory):
    """RefreshSessionRead strictly avoids exposing raw hashes over the wire."""
    session_obj = refresh_session_factory(
        csrf_token_hash="super_secret_csrf_hash",
    )

    read_dto = RefreshSessionRead.model_validate(session_obj)
    dumped = read_dto.model_dump()

    assert "refresh_token_hash" not in dumped
    assert "csrf_token_hash" not in dumped
    assert "refresh_token_hash" not in RefreshSessionRead.model_fields
    assert "csrf_token_hash" not in RefreshSessionRead.model_fields


def test_refresh_session_read_reflects_revoked_state(refresh_session_factory):
    """Revoked session maps is_active as False in RefreshSessionRead."""
    now = datetime.datetime.now(datetime.timezone.utc)
    session_obj = refresh_session_factory(
        revoked_at=now,
        revoked_reason="logout",
    )

    read_dto = RefreshSessionRead.model_validate(session_obj)

    assert read_dto.revoked_at == session_obj.revoked_at
    assert read_dto.revoked_reason == "logout"
    assert read_dto.is_active is False
