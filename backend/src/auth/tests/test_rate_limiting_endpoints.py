"""
Integration tests for Phase 11 Step 3: Endpoint Rate Limiting.

Validates that:
- High-frequency requests to /api/v1/auth/login trigger HTTP 429 Too Many Requests.
- High-frequency requests to /api/v1/auth/signup trigger HTTP 429 Too Many Requests.
- High-frequency requests to /api/v1/auth/refresh trigger HTTP 429 Too Many Requests.
- The standard 'Retry-After' response header is included in 429 responses.
- Rate limits are tracked independently per client IP (using X-Forwarded-For or client host).
"""

import pytest
from src.utils.rate_limiter import auth_rate_limiter


@pytest.fixture(autouse=True)
def reset_limiter():
    """Ensures the auth rate limiter starts with a clean slate before each test."""
    auth_rate_limiter.reset()
    yield
    auth_rate_limiter.reset()


def test_login_endpoint_rate_limiting(client):
    """Exceeding the login rate limit threshold returns HTTP 429 with Retry-After header."""
    ip_headers = {"X-Forwarded-For": "203.0.113.195"}
    payload = {"username": "doctor@hospital.org", "password": "WrongPassword123!"}

    # Consume the allowed requests for the login window (e.g. 5 requests)
    for _ in range(5):
        client.post("/api/v1/auth/login", data=payload, headers=ip_headers)

    # 6th attempt must be throttled with 429
    resp = client.post("/api/v1/auth/login", data=payload, headers=ip_headers)
    assert resp.status_code == 429
    assert "Retry-After" in resp.headers
    assert int(resp.headers["Retry-After"]) > 0
    assert "Too many requests" in resp.json()["detail"]


def test_signup_endpoint_rate_limiting(client):
    """Exceeding the signup rate limit threshold returns HTTP 429 with Retry-After header."""
    ip_headers = {"X-Forwarded-For": "203.0.113.196"}
    payload = {
        "invitation_token": "some-random-token",
        "first_name": "Test",
        "last_name": "User",
        "email": "test@hospital.org",
        "password": "SecurePassword123!",
    }

    # Consume allowed attempts for signup
    for _ in range(5):
        client.post("/api/v1/auth/signup", json=payload, headers=ip_headers)

    # Subsequent attempt must be blocked
    resp = client.post("/api/v1/auth/signup", json=payload, headers=ip_headers)
    assert resp.status_code == 429
    assert "Retry-After" in resp.headers
    assert int(resp.headers["Retry-After"]) > 0
    assert "Too many requests" in resp.json()["detail"]


def test_refresh_endpoint_rate_limiting(client):
    """Exceeding the refresh rate limit threshold returns HTTP 429 with Retry-After header."""
    ip_headers = {"X-Forwarded-For": "203.0.113.197"}

    # Consume allowed attempts for refresh
    for _ in range(10):
        client.post("/api/v1/auth/refresh", headers=ip_headers)

    # Next attempt must be blocked
    resp = client.post("/api/v1/auth/refresh", headers=ip_headers)
    assert resp.status_code == 429
    assert "Retry-After" in resp.headers
    assert int(resp.headers["Retry-After"]) > 0
    assert "Too many requests" in resp.json()["detail"]


def test_rate_limiting_is_isolated_per_client_ip(client):
    """A client being rate-limited does not affect requests from a different client IP."""
    attacker_headers = {"X-Forwarded-For": "198.51.100.1"}
    legitimate_headers = {"X-Forwarded-For": "198.51.100.2"}

    payload = {"username": "admin@hospital.org", "password": "WrongPassword123!"}

    # Exhaust the limit for the attacker IP
    for _ in range(5):
        client.post("/api/v1/auth/login", data=payload, headers=attacker_headers)

    # Attacker is now blocked with 429
    blocked_resp = client.post("/api/v1/auth/login", data=payload, headers=attacker_headers)
    assert blocked_resp.status_code == 429

    # Legitimate user from different IP must NOT be blocked (will get 401 invalid credentials, not 429)
    legit_resp = client.post("/api/v1/auth/login", data=payload, headers=legitimate_headers)
    assert legit_resp.status_code == 401
