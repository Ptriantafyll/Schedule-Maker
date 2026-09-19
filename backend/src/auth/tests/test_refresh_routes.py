"""
Integration and security tests for POST /api/v1/auth/refresh and POST /api/v1/auth/logout.
"""

import uuid
import datetime
import pytest
from sqlmodel import Session, select

from src.auth.models import RefreshSession
from src.auth import security
from src.user.models import UserRole

LOGIN_EMAIL = "doctor@test.com"
LOGIN_PASSWORD = "Password123!"


@pytest.fixture(name="auth_user")
def auth_user_fixture(user_factory):
    """Creates a user for route tests."""
    return user_factory(
        role=UserRole.SUPER_ADMIN,
        email=LOGIN_EMAIL,
        password=LOGIN_PASSWORD,
    )


def test_refresh_via_json_body_success(client, auth_user, session: Session):
    """Clients can rotate tokens by sending refresh_token in JSON request body."""
    # 1. Login to establish session
    login_resp = client.post(
        "/api/v1/auth/login",
        data={"username": auth_user.email, "password": LOGIN_PASSWORD},
    )
    assert login_resp.status_code == 200
    login_data = login_resp.json()
    raw_refresh_1 = login_data["refresh_token"]

    # Clear client cookies so this request tests the JSON body path exclusively
    client.cookies.clear()

    # 2. Call refresh via body
    refresh_resp = client.post(
        "/api/v1/auth/refresh",
        json={"refresh_token": raw_refresh_1},
    )
    assert refresh_resp.status_code == 200
    refresh_data = refresh_resp.json()

    assert refresh_data["token_type"] == "bearer"
    assert isinstance(refresh_data["access_token"], str)
    assert isinstance(refresh_data["refresh_token"], str)
    assert refresh_data["refresh_token"] != raw_refresh_1

    # Verify access token claims
    payload = security.decode_access_token(refresh_data["access_token"])
    assert payload["sub"] == str(auth_user.id)

    # Verify database session was rotated
    old_hash = security.hash_refresh_token(raw_refresh_1)
    old_session = session.exec(
        select(RefreshSession).where(RefreshSession.refresh_token_hash == old_hash)
    ).first()
    assert old_session.is_active is False
    assert old_session.revoked_reason == "rotated"
    assert old_session.replaced_by_session_id is not None


def test_refresh_via_cookie_with_csrf_header_success(client, auth_user):
    """Browser clients with cookies succeed when valid X-CSRF-Token header is provided."""
    # 1. Login sets cookies
    login_resp = client.post(
        "/api/v1/auth/login",
        data={"username": auth_user.email, "password": LOGIN_PASSWORD},
    )
    assert login_resp.status_code == 200
    csrf_token = login_resp.json()["csrf_token"]
    old_refresh_cookie = client.cookies.get("refresh_token")

    # 2. Call refresh with cookies + X-CSRF-Token header
    refresh_resp = client.post(
        "/api/v1/auth/refresh",
        headers={"X-CSRF-Token": csrf_token},
    )
    assert refresh_resp.status_code == 200
    refresh_data = refresh_resp.json()

    assert isinstance(refresh_data["access_token"], str)
    assert isinstance(refresh_data["refresh_token"], str)

    # Verify new cookies are updated on the client
    new_refresh_cookie = client.cookies.get("refresh_token")
    assert new_refresh_cookie is not None
    assert new_refresh_cookie != old_refresh_cookie


def test_refresh_via_cookie_without_csrf_header_forbidden(client, auth_user):
    """Cookie-based refresh requests without X-CSRF-Token are rejected with 403."""
    # 1. Login sets cookies
    login_resp = client.post(
        "/api/v1/auth/login",
        data={"username": auth_user.email, "password": LOGIN_PASSWORD},
    )
    assert login_resp.status_code == 200

    # 2. Call refresh with cookies but NO CSRF header
    refresh_resp = client.post("/api/v1/auth/refresh")
    assert refresh_resp.status_code == 403
    assert "CSRF" in refresh_resp.json()["detail"]


def test_refresh_via_cookie_with_invalid_csrf_header_forbidden(client, auth_user):
    """Cookie-based refresh requests with mismatched X-CSRF-Token are rejected with 403."""
    login_resp = client.post(
        "/api/v1/auth/login",
        data={"username": auth_user.email, "password": LOGIN_PASSWORD},
    )
    assert login_resp.status_code == 200

    refresh_resp = client.post(
        "/api/v1/auth/refresh",
        headers={"X-CSRF-Token": "invalid_forged_csrf_token"},
    )
    assert refresh_resp.status_code == 403
    assert "CSRF" in refresh_resp.json()["detail"]


def test_refresh_missing_token_unauthorized(client):
    """Calling refresh with neither body nor cookie raises 401."""
    client.cookies.clear()
    resp = client.post("/api/v1/auth/refresh")
    assert resp.status_code == 401
    assert "Refresh token" in resp.json()["detail"]


def test_refresh_invalid_token_unauthorized(client):
    """Calling refresh with an unknown token string raises 401."""
    client.cookies.clear()
    resp = client.post(
        "/api/v1/auth/refresh",
        json={"refresh_token": "unknown_token_value_abc"},
    )
    assert resp.status_code == 401
    assert "Invalid refresh token" in resp.json()["detail"]


def test_refresh_token_reuse_revokes_entire_family(client, auth_user, session: Session):
    """Replaying an already-rotated token revokes the entire family and returns 401."""
    login_resp = client.post(
        "/api/v1/auth/login",
        data={"username": auth_user.email, "password": LOGIN_PASSWORD},
    )
    raw_token_1 = login_resp.json()["refresh_token"]

    client.cookies.clear()

    # Rotate token 1 -> token 2
    rotate_resp = client.post(
        "/api/v1/auth/refresh",
        json={"refresh_token": raw_token_1},
    )
    assert rotate_resp.status_code == 200
    raw_token_2 = rotate_resp.json()["refresh_token"]

    # Attacker attempts to replay token 1!
    replay_resp = client.post(
        "/api/v1/auth/refresh",
        json={"refresh_token": raw_token_1},
    )
    assert replay_resp.status_code == 401
    assert "reuse detected" in replay_resp.json()["detail"].lower()

    # Both tokens are now dead; token 2 is revoked
    dead_resp = client.post(
        "/api/v1/auth/refresh",
        json={"refresh_token": raw_token_2},
    )
    assert dead_resp.status_code == 401


def test_refresh_expired_token_unauthorized(client, auth_user, refresh_session_factory):
    """Submitting an expired refresh token returns 401."""
    client.cookies.clear()
    raw_token = "expired_token_test_123"
    token_hash = security.hash_refresh_token(raw_token)
    now = datetime.datetime.now(datetime.timezone.utc)

    refresh_session_factory(
        user_id=auth_user.id,
        refresh_token_hash=token_hash,
        expires_at=now - datetime.timedelta(hours=1),
    )

    resp = client.post(
        "/api/v1/auth/refresh",
        json={"refresh_token": raw_token},
    )
    assert resp.status_code == 401
    assert "expired" in resp.json()["detail"].lower()


def test_logout_via_body_terminates_session(client, auth_user, session: Session):
    """POST /auth/logout with body revokes the active session."""
    login_resp = client.post(
        "/api/v1/auth/login",
        data={"username": auth_user.email, "password": LOGIN_PASSWORD},
    )
    raw_refresh = login_resp.json()["refresh_token"]

    client.cookies.clear()

    logout_resp = client.post(
        "/api/v1/auth/logout",
        json={"refresh_token": raw_refresh},
    )
    assert logout_resp.status_code == 200
    assert logout_resp.json() == {"detail": "Successfully logged out"}

    # Verify session in DB is revoked
    token_hash = security.hash_refresh_token(raw_refresh)
    db_session = session.exec(
        select(RefreshSession).where(RefreshSession.refresh_token_hash == token_hash)
    ).first()
    assert db_session.is_active is False
    assert db_session.revoked_reason == "logout"


def test_logout_via_cookies_clears_cookies(client, auth_user):
    """POST /auth/logout with cookies revokes session and expires the cookies."""
    login_resp = client.post(
        "/api/v1/auth/login",
        data={"username": auth_user.email, "password": LOGIN_PASSWORD},
    )
    assert login_resp.status_code == 200
    assert client.cookies.get("refresh_token") is not None

    logout_resp = client.post("/api/v1/auth/logout")
    assert logout_resp.status_code == 200
    assert logout_resp.json() == {"detail": "Successfully logged out"}

    # Cookies should be cleared/expired on client
    # In starlette TestClient, deleted cookies have empty values or max-age=0
    refresh_val = client.cookies.get("refresh_token")
    assert refresh_val is None or refresh_val == '""' or refresh_val == ""


def test_logout_idempotent_when_no_session(client):
    """POST /auth/logout is safe and returns 200 even if no active session exists."""
    client.cookies.clear()
    resp = client.post("/api/v1/auth/logout")
    assert resp.status_code == 200
    assert resp.json() == {"detail": "Successfully logged out"}
