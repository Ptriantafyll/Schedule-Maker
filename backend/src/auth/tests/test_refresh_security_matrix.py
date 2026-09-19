"""
Security Matrix & Attack Simulation Tests for Refresh Session Subsystem.

Validates RFC 6819 threat mitigations:
- Multi-hop token rotation and historical token reuse detection (family invalidation)
- Independent session family isolation across multiple user devices
- Post-logout token replay rejection
- Expired token rejection at HTTP boundary
- Cross-user session revocation isolation
- Tampered and malformed token resilience
- CSRF double-submit protection in cookie-based transport mode
"""

import uuid
import datetime
import pytest
from sqlmodel import Session, select

from src.auth.models import RefreshSession
from src.auth import security, services
from src.user.models import UserRole

LOGIN_PASSWORD = "Password123!"


@pytest.fixture(name="user_alice")
def user_alice_fixture(user_factory):
    """Creates first test user."""
    return user_factory(
        role=UserRole.SUPER_ADMIN,
        email="alice@test.com",
        password=LOGIN_PASSWORD,
    )


@pytest.fixture(name="user_bob")
def user_bob_fixture(user_factory):
    """Creates second test user."""
    return user_factory(
        role=UserRole.SUPER_ADMIN,
        email="bob@test.com",
        password=LOGIN_PASSWORD,
    )



# --------------------------------------------------------------------------
# 1. Multi-Hop Lineage & Historical Replay Attack Simulation
# --------------------------------------------------------------------------

def test_multi_hop_rotation_and_historical_replay_neutralizes_family(
    client, user_alice, session: Session
):
    """
    Scenario:
    1. Client logs in: receives T_1.
    2. Client rotates T_1 -> T_2 -> T_3 -> T_4.
    3. Attacker steals and replays T_1 (or T_2).
    4. Server detects reuse, rejects attacker (401), and revokes all sessions in family.
    5. Legitimate client subsequently presents latest token T_4 -> rejected (401).
    """
    # 1. Initial Login
    login_resp = client.post(
        "/api/v1/auth/login",
        data={"username": user_alice.email, "password": LOGIN_PASSWORD},
    )
    assert login_resp.status_code == 200
    t1 = login_resp.json()["refresh_token"]

    client.cookies.clear()

    # 2. Multi-Hop Rotations: T1 -> T2 -> T3 -> T4
    resp_2 = client.post("/api/v1/auth/refresh", json={"refresh_token": t1})
    assert resp_2.status_code == 200
    t2 = resp_2.json()["refresh_token"]

    resp_3 = client.post("/api/v1/auth/refresh", json={"refresh_token": t2})
    assert resp_3.status_code == 200
    t3 = resp_3.json()["refresh_token"]

    resp_4 = client.post("/api/v1/auth/refresh", json={"refresh_token": t3})
    assert resp_4.status_code == 200
    t4 = resp_4.json()["refresh_token"]

    # Verify all 4 tokens are distinct
    assert len({t1, t2, t3, t4}) == 4

    # 3. Attacker replays historical token T1
    replay_resp = client.post("/api/v1/auth/refresh", json={"refresh_token": t1})
    assert replay_resp.status_code == 401
    assert "reuse detected" in replay_resp.json()["detail"].lower()

    # 4. Attacker also tries another historical token T2 -> rejected
    replay_t2_resp = client.post("/api/v1/auth/refresh", json={"refresh_token": t2})
    assert replay_t2_resp.status_code == 401

    # 5. Legitimate client attempts to use latest token T4 -> rejected because family was nuked
    client_t4_resp = client.post("/api/v1/auth/refresh", json={"refresh_token": t4})
    assert client_t4_resp.status_code == 401

    # 6. Verify database lineage state
    h1 = security.hash_refresh_token(t1)
    s1 = session.exec(select(RefreshSession).where(RefreshSession.refresh_token_hash == h1)).first()
    family_id = s1.session_family

    all_family_sessions = session.exec(
        select(RefreshSession).where(RefreshSession.session_family == family_id)
    ).all()
    assert len(all_family_sessions) == 4
    for s in all_family_sessions:
        assert s.is_active is False

    # The active tip of the family (T4) was revoked with reuse_detected
    h4 = security.hash_refresh_token(t4)
    s4 = session.exec(select(RefreshSession).where(RefreshSession.refresh_token_hash == h4)).first()
    assert s4.revoked_reason == "reuse_detected"



# --------------------------------------------------------------------------
# 2. Independent Session Family Isolation Across Devices
# --------------------------------------------------------------------------

def test_multiple_devices_independent_family_isolation(
    client, user_alice, session: Session
):
    """
    Scenario:
    1. User logs in on Device A (Family A: T_A1) and Device B (Family B: T_B1).
    2. Device A rotates T_A1 -> T_A2.
    3. Attacker replays T_A1 -> Family A is revoked.
    4. Device B rotates T_B1 -> T_B2 -> SUCCEEDS (Family B is isolated and unaffected).
    """
    # Login on Device A
    resp_a = client.post(
        "/api/v1/auth/login",
        data={"username": user_alice.email, "password": LOGIN_PASSWORD},
    )
    assert resp_a.status_code == 200
    t_a1 = resp_a.json()["refresh_token"]

    # Login on Device B
    resp_b = client.post(
        "/api/v1/auth/login",
        data={"username": user_alice.email, "password": LOGIN_PASSWORD},
    )
    assert resp_b.status_code == 200
    t_b1 = resp_b.json()["refresh_token"]

    client.cookies.clear()

    # Device A rotates T_A1 -> T_A2
    resp_a2 = client.post("/api/v1/auth/refresh", json={"refresh_token": t_a1})
    assert resp_a2.status_code == 200
    t_a2 = resp_a2.json()["refresh_token"]

    # Attacker replays T_A1 -> Family A compromised & revoked
    replay_a = client.post("/api/v1/auth/refresh", json={"refresh_token": t_a1})
    assert replay_a.status_code == 401
    assert "reuse detected" in replay_a.json()["detail"].lower()

    # Device A's latest token T_A2 is now rejected
    resp_a2_replay = client.post("/api/v1/auth/refresh", json={"refresh_token": t_a2})
    assert resp_a2_replay.status_code == 401

    # Device B's session T_B1 must still work normally and rotate to T_B2!
    resp_b2 = client.post("/api/v1/auth/refresh", json={"refresh_token": t_b1})
    assert resp_b2.status_code == 200
    t_b2 = resp_b2.json()["refresh_token"]
    assert isinstance(t_b2, str)
    assert t_b2 != t_b1


# --------------------------------------------------------------------------
# 3. Post-Logout Replay Vector
# --------------------------------------------------------------------------

def test_post_logout_replay_attack_rejected(client, user_alice, session: Session):
    """
    Scenario:
    1. User logs in: receives T_1.
    2. User explicitly logs out with T_1.
    3. Attacker attempts to use T_1 -> rejected (401: revoked).
    """
    login_resp = client.post(
        "/api/v1/auth/login",
        data={"username": user_alice.email, "password": LOGIN_PASSWORD},
    )
    t1 = login_resp.json()["refresh_token"]
    client.cookies.clear()

    # Logout
    logout_resp = client.post("/api/v1/auth/logout", json={"refresh_token": t1})
    assert logout_resp.status_code == 200

    # Attacker tries to refresh with T1
    replay_resp = client.post("/api/v1/auth/refresh", json={"refresh_token": t1})
    assert replay_resp.status_code == 401
    assert "revoked" in replay_resp.json()["detail"].lower()


# --------------------------------------------------------------------------
# 4. Expired Token Matrix
# --------------------------------------------------------------------------

def test_expired_token_rejected_at_endpoint(client, user_alice, session: Session):
    """
    Scenario:
    1. A refresh session in DB whose expires_at is in the past.
    2. Client presents raw token -> rejected with 401 ('Refresh token expired').
    """
    raw_token = security.generate_refresh_token()
    raw_csrf = security.generate_csrf_token()
    token_hash = security.hash_refresh_token(raw_token)
    csrf_hash = security.hash_csrf_token(raw_csrf)

    past_time = datetime.datetime.now(datetime.timezone.utc) - datetime.timedelta(days=1)

    expired_session = RefreshSession(
        user_id=user_alice.id,
        refresh_token_hash=token_hash,
        csrf_token_hash=csrf_hash,
        session_family=uuid.uuid4(),
        expires_at=past_time,
    )
    session.add(expired_session)
    session.commit()

    client.cookies.clear()
    resp = client.post("/api/v1/auth/refresh", json={"refresh_token": raw_token})
    assert resp.status_code == 401
    assert "expired" in resp.json()["detail"].lower()


# --------------------------------------------------------------------------
# 5. Cross-User Session Revocation Isolation
# --------------------------------------------------------------------------

def test_cross_user_session_revocation_isolation(
    client, user_alice, user_bob, session: Session
):
    """
    Scenario:
    1. Alice and Bob both log in.
    2. Admin/emergency revokes all sessions for Alice.
    3. Alice's refresh token is rejected (401).
    4. Bob's refresh token continues to rotate successfully (200).
    """
    resp_a = client.post(
        "/api/v1/auth/login",
        data={"username": user_alice.email, "password": LOGIN_PASSWORD},
    )
    t_alice = resp_a.json()["refresh_token"]

    resp_b = client.post(
        "/api/v1/auth/login",
        data={"username": user_bob.email, "password": LOGIN_PASSWORD},
    )
    t_bob = resp_b.json()["refresh_token"]

    client.cookies.clear()

    # Emergency revoke for Alice only
    services.revoke_all_user_refresh_sessions(
        session=session,
        user_id=user_alice.id,
        reason="admin_lockout",
    )

    # Alice fails
    resp_alice_refresh = client.post("/api/v1/auth/refresh", json={"refresh_token": t_alice})
    assert resp_alice_refresh.status_code == 401

    # Bob succeeds
    resp_bob_refresh = client.post("/api/v1/auth/refresh", json={"refresh_token": t_bob})
    assert resp_bob_refresh.status_code == 200
    assert resp_bob_refresh.json()["refresh_token"] != t_bob


# --------------------------------------------------------------------------
# 6. Malformed and Random Token Resilience
# --------------------------------------------------------------------------

@pytest.mark.parametrize(
    "bad_token",
    [
        "completely_random_nonexistent_token_string",
        "a" * 64,
        "undefined",
        "null",
    ],
)
def test_malformed_tokens_fail_safely(client, bad_token):
    """Unknown/garbage tokens return 401 Unauthorized without 500 errors."""
    client.cookies.clear()
    resp = client.post("/api/v1/auth/refresh", json={"refresh_token": bad_token})
    assert resp.status_code == 401
    assert "invalid" in resp.json()["detail"].lower()


# --------------------------------------------------------------------------
# 7. Cookie Mode CSRF Tampering Protection
# --------------------------------------------------------------------------

def test_cookie_mode_csrf_tampering_rejected(client, user_alice):
    """
    When using cookie-based transport:
    - Missing X-CSRF-Token header -> 403 Forbidden.
    - Tampered/invalid X-CSRF-Token header -> 403 Forbidden.
    - Legitimate X-CSRF-Token header -> 200 OK.
    """
    # 1. Login sets cookies in client
    login_resp = client.post(
        "/api/v1/auth/login",
        data={"username": user_alice.email, "password": LOGIN_PASSWORD},
    )
    assert login_resp.status_code == 200
    real_csrf = client.cookies.get("csrf_token")
    assert real_csrf is not None

    # A: Missing CSRF header
    resp_missing = client.post("/api/v1/auth/refresh")
    assert resp_missing.status_code == 403
    assert "csrf" in resp_missing.json()["detail"].lower()

    # B: Tampered CSRF header
    resp_tampered = client.post(
        "/api/v1/auth/refresh",
        headers={"X-CSRF-Token": "tampered_fake_csrf_token_value"},
    )
    assert resp_tampered.status_code == 403
    assert "csrf" in resp_tampered.json()["detail"].lower()

    # C: Legitimate CSRF header -> 200 OK
    resp_valid = client.post(
        "/api/v1/auth/refresh",
        headers={"X-CSRF-Token": real_csrf},
    )
    assert resp_valid.status_code == 200
    assert "access_token" in resp_valid.json()


# --------------------------------------------------------------------------
# 8. Inactive/Deleted User Session Revocation on Refresh
# --------------------------------------------------------------------------

def test_refresh_fails_and_revokes_when_user_is_deleted(
    client, user_alice, session: Session
):
    """
    Scenario:
    1. Alice logs in and receives refresh token T_1.
    2. Admin soft-deletes Alice (user.is_deleted = True).
    3. Alice attempts to refresh credentials using T_1.
    4. Backend detects inactive user -> rejects with 401 and revokes session family.
    5. Database verifies session family is marked inactive with reason 'inactive_user'.
    """
    login_resp = client.post(
        "/api/v1/auth/login",
        data={"username": user_alice.email, "password": LOGIN_PASSWORD},
    )
    assert login_resp.status_code == 200
    t1 = login_resp.json()["refresh_token"]

    # Soft-delete the user
    user_alice.is_deleted = True
    session.add(user_alice)
    session.commit()

    client.cookies.clear()
    refresh_resp = client.post("/api/v1/auth/refresh", json={"refresh_token": t1})
    assert refresh_resp.status_code == 401
    assert "invalid" in refresh_resp.json()["detail"].lower()


    # Verify session family was revoked in DB
    token_hash = security.hash_refresh_token(t1)
    db_session = session.exec(
        select(RefreshSession).where(RefreshSession.refresh_token_hash == token_hash)
    ).first()
    assert db_session.is_active is False
    assert db_session.revoked_reason == "inactive_user"

