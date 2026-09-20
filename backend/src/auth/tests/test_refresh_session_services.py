"""
Unit tests for the RefreshSession service layer and reuse detection logic.
"""

import datetime
import pytest
from sqlmodel import Session

from src.auth.services import (
    issue_refresh_session,
    rotate_refresh_token,
    terminate_refresh_session,
    revoke_all_user_refresh_sessions,
    InvalidRefreshTokenError,
    ExpiredRefreshTokenError,
    RevokedRefreshTokenError,
    RefreshTokenReuseDetectedError,
)
from src.auth.security import hash_refresh_token, hash_csrf_token
from src.user.models import UserRole


def test_issue_refresh_session(session: Session, user_factory):
    """issue_refresh_session generates raw tokens, hashes them, and persists a session."""
    user = user_factory(role=UserRole.SUPER_ADMIN)

    session_obj, raw_refresh, raw_csrf = issue_refresh_session(
        session=session,
        user_id=user.id,
    )

    assert isinstance(raw_refresh, str)
    assert len(raw_refresh) > 30
    assert isinstance(raw_csrf, str)
    assert len(raw_csrf) > 30

    assert session_obj.user_id == user.id
    assert session_obj.refresh_token_hash == hash_refresh_token(raw_refresh)
    assert session_obj.csrf_token_hash == hash_csrf_token(raw_csrf)
    assert session_obj.is_active is True
    assert session_obj.is_expired is False
    assert session_obj.replaced_by_session_id is None


def test_rotate_refresh_token_happy_path(session: Session, user_factory):
    """rotate_refresh_token invalidates predecessor, issues successor, and preserves family."""
    user = user_factory(role=UserRole.SUPER_ADMIN)
    session_1, raw_refresh_1, _ = issue_refresh_session(session=session, user_id=user.id)
    original_family = session_1.session_family

    session_2, raw_refresh_2, raw_csrf_2 = rotate_refresh_token(
        session=session,
        raw_refresh_token=raw_refresh_1,
    )

    session.refresh(session_1)

    # Successor assertions
    assert session_2.id != session_1.id
    assert session_2.session_family == original_family
    assert session_2.user_id == user.id
    assert session_2.refresh_token_hash == hash_refresh_token(raw_refresh_2)
    assert session_2.csrf_token_hash == hash_csrf_token(raw_csrf_2)
    assert session_2.is_active is True
    assert raw_refresh_2 != raw_refresh_1

    # Predecessor assertions
    assert session_1.replaced_by_session_id == session_2.id
    assert session_1.revoked_reason == "rotated"
    assert session_1.revoked_at is not None
    assert session_1.is_active is False


def test_rotate_refresh_token_invalid_raises(session: Session):
    """rotate_refresh_token raises InvalidRefreshTokenError for unknown tokens."""
    with pytest.raises(InvalidRefreshTokenError):
        rotate_refresh_token(
            session=session,
            raw_refresh_token="non_existent_token_string",
        )


def test_rotate_refresh_token_reuse_detection(session: Session, user_factory):
    """Presenting an already-replaced refresh token revokes the entire session family."""
    user = user_factory(role=UserRole.SUPER_ADMIN)
    session_1, raw_refresh_1, _ = issue_refresh_session(session=session, user_id=user.id)

    # Legitimate user rotates session 1 -> session 2
    session_2, raw_refresh_2, _ = rotate_refresh_token(
        session=session,
        raw_refresh_token=raw_refresh_1,
    )
    assert session_2.is_active is True

    # Attacker replays raw_refresh_1 after it was already rotated!
    with pytest.raises(RefreshTokenReuseDetectedError):
        rotate_refresh_token(
            session=session,
            raw_refresh_token=raw_refresh_1,
        )

    # Both session 1 and session 2 are now revoked!
    session.refresh(session_1)
    session.refresh(session_2)

    assert session_1.is_active is False
    assert session_2.is_active is False
    assert session_2.revoked_reason == "reuse_detected"

    # Subsequent attempts to rotate session_2 will also fail
    with pytest.raises(RevokedRefreshTokenError):
        rotate_refresh_token(
            session=session,
            raw_refresh_token=raw_refresh_2,
        )


def test_rotate_refresh_token_revoked_raises(session: Session, user_factory, refresh_session_factory):
    """rotate_refresh_token raises RevokedRefreshTokenError if session was revoked without successor."""
    user = user_factory(role=UserRole.SUPER_ADMIN)
    raw_token = "revoked_raw_token_value_123"
    token_hash = hash_refresh_token(raw_token)

    refresh_session_factory(
        user_id=user.id,
        refresh_token_hash=token_hash,
        revoked_at=datetime.datetime.now(datetime.timezone.utc),
        revoked_reason="logout",
    )

    with pytest.raises(RevokedRefreshTokenError):
        rotate_refresh_token(
            session=session,
            raw_refresh_token=raw_token,
        )


def test_rotate_refresh_token_expired_raises(session: Session, user_factory, refresh_session_factory):
    """rotate_refresh_token raises ExpiredRefreshTokenError if session expiration has passed."""
    user = user_factory(role=UserRole.SUPER_ADMIN)
    raw_token = "expired_raw_token_value_456"
    token_hash = hash_refresh_token(raw_token)
    now = datetime.datetime.now(datetime.timezone.utc)

    refresh_session_factory(
        user_id=user.id,
        refresh_token_hash=token_hash,
        expires_at=now - datetime.timedelta(hours=2),
    )

    with pytest.raises(ExpiredRefreshTokenError):
        rotate_refresh_token(
            session=session,
            raw_refresh_token=raw_token,
        )


def test_terminate_refresh_session_success(session: Session, user_factory):
    """terminate_refresh_session successfully revokes an active session."""
    user = user_factory(role=UserRole.SUPER_ADMIN)
    session_obj, raw_refresh, _ = issue_refresh_session(session=session, user_id=user.id)

    result = terminate_refresh_session(session=session, raw_refresh_token=raw_refresh)

    assert result is True
    session.refresh(session_obj)
    assert session_obj.is_active is False
    assert session_obj.revoked_reason == "logout"
    assert session_obj.revoked_at is not None


def test_terminate_refresh_session_unknown_returns_false(session: Session):
    """terminate_refresh_session returns False for non-existent tokens."""
    result = terminate_refresh_session(
        session=session,
        raw_refresh_token="unknown_refresh_token",
    )
    assert result is False


def test_revoke_all_user_refresh_sessions(session: Session, user_factory):
    """revoke_all_user_refresh_sessions bulk-revokes all active sessions for that user."""
    user_1 = user_factory(role=UserRole.SUPER_ADMIN)
    user_2 = user_factory(role=UserRole.SUPER_ADMIN)

    s1, raw1, _ = issue_refresh_session(session=session, user_id=user_1.id)
    s2, raw2, _ = issue_refresh_session(session=session, user_id=user_1.id)
    s3, _, _ = issue_refresh_session(session=session, user_id=user_2.id)

    count = revoke_all_user_refresh_sessions(
        session=session,
        user_id=user_1.id,
        reason="password_reset",
    )

    assert count == 2

    session.refresh(s1)
    session.refresh(s2)
    session.refresh(s3)

    assert s1.is_active is False
    assert s1.revoked_reason == "password_reset"
    assert s2.is_active is False
    assert s2.revoked_reason == "password_reset"

    assert s3.is_active is True
    assert s3.revoked_at is None
