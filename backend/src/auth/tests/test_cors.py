"""
CORS Hardening and Origin Whitelisting Tests.

Validates:
- Trusted origins receive Access-Control-Allow-Origin and Access-Control-Allow-Credentials: true
- Untrusted origins are denied CORS headers
- Allowed headers are restricted to approved list (Authorization, Content-Type, X-CSRF-Token)
- Wildcard origins and wildcard headers are forbidden
- Configurable CORS origins via environment variable helper
"""

from src.main import get_cors_origins, ALLOWED_CORS_HEADERS, ALLOWED_CORS_METHODS


def test_cors_origins_helper_defaults():
    """When no env var is set, returns default local development origins."""
    origins = get_cors_origins(raw_origins="")
    assert "http://localhost:3000" in origins
    assert "http://127.0.0.1:3000" in origins
    assert "http://localhost:8000" in origins


def test_cors_origins_helper_parses_comma_separated_env(monkeypatch):
    """Parses, trims, and cleans comma-separated origins from environment."""
    raw = " https://app.hospital.com, http://localhost:4000 , , https://admin.hospital.com "
    origins = get_cors_origins(raw_origins=raw)
    assert origins == [
        "https://app.hospital.com",
        "http://localhost:4000",
        "https://admin.hospital.com",
    ]


def test_trusted_origin_receives_cors_credentials_headers(client):
    """A trusted origin receives matching allow-origin and credentials header."""
    response = client.options(
        "/api/v1/auth/login",
        headers={
            "Origin": "http://localhost:3000",
            "Access-Control-Request-Method": "POST",
            "Access-Control-Request-Headers": "authorization, content-type, x-csrf-token",
        },
    )
    assert response.status_code == 200
    assert response.headers.get("access-control-allow-origin") == "http://localhost:3000"
    assert response.headers.get("access-control-allow-credentials") == "true"


def test_untrusted_origin_denied_cors_headers(client):
    """An untrusted origin does NOT receive access-control-allow-origin header."""
    response = client.options(
        "/api/v1/auth/login",
        headers={
            "Origin": "http://malicious-attacker-site.com",
            "Access-Control-Request-Method": "POST",
        },
    )
    assert "access-control-allow-origin" not in response.headers


def test_untrusted_origin_simple_request_omits_allow_origin(client):
    """A GET/POST request from an untrusted origin does not receive allow-origin."""
    response = client.get(
        "/health",
        headers={
            "Origin": "http://evil-tracker.org",
        },
    )
    assert response.status_code == 200
    assert "access-control-allow-origin" not in response.headers


def test_disallowed_headers_rejected_on_preflight(client):
    """Untrusted/unapproved headers are not permitted in CORS preflight."""
    response = client.options(
        "/api/v1/auth/login",
        headers={
            "Origin": "http://localhost:3000",
            "Access-Control-Request-Method": "POST",
            "Access-Control-Request-Headers": "x-malicious-exploit-header",
        },
    )
    # The header must not be approved in Access-Control-Allow-Headers
    allowed_headers = response.headers.get("access-control-allow-headers", "").lower()
    assert "x-malicious-exploit-header" not in allowed_headers


def test_cors_constants_are_restricted():
    """CORS configuration does not use wildcards for methods or headers."""
    assert "*" not in ALLOWED_CORS_METHODS
    assert "*" not in ALLOWED_CORS_HEADERS
    for required_header in ["authorization", "content-type", "x-csrf-token"]:
        assert any(required_header == h.lower() for h in ALLOWED_CORS_HEADERS)
