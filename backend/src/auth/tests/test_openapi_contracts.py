"""
OpenAPI Schema & Contract Verification Tests.

Validates RFC and security contract requirements:
- Only explicit public routes omit security requirements.
- All domain routes require OAuth2PasswordBearer authentication.
- Obsolete routes are absent from the published schema.
- Public signup schema cannot accept role or department injection.
"""

import pytest
from src.main import app

PUBLIC_ROUTES = {
    ("/health", "get"),
    ("/api/v1/auth/login", "post"),
    ("/api/v1/auth/signup", "post"),
    ("/api/v1/auth/refresh", "post"),
    ("/api/v1/auth/logout", "post"),
}

OBSOLETE_ROUTES = [
    "/api/v1/users/login",
    "/api/v1/users/signup",
    "/api/v1/users/register",
    "/auth/login",
    "/users/login",
]


@pytest.fixture(name="openapi_schema", scope="module")
def openapi_schema_fixture():
    """Returns the freshly generated OpenAPI specification dictionary."""
    return app.openapi()


def test_openapi_declares_oauth2_security_scheme(openapi_schema):
    """OpenAPI schema must declare OAuth2PasswordBearer security scheme."""
    security_schemes = openapi_schema.get("components", {}).get("securitySchemes", {})
    assert "OAuth2PasswordBearer" in security_schemes

    scheme = security_schemes["OAuth2PasswordBearer"]
    assert scheme["type"] == "oauth2"
    assert "password" in scheme["flows"]
    assert scheme["flows"]["password"]["tokenUrl"] == "/api/v1/auth/login"


def test_openapi_public_routes_strictly_bounded(openapi_schema):
    """Only approved public routes are allowed to have no security requirements."""
    paths = openapi_schema.get("paths", {})
    actual_public_routes = set()

    for path, methods in paths.items():
        for method, operation in methods.items():
            if method.lower() in {"get", "post", "put", "patch", "delete"}:
                security = operation.get("security")
                if not security:
                    actual_public_routes.add((path, method.lower()))

    assert actual_public_routes == PUBLIC_ROUTES, (
        f"Public routes do not match expected whitelist. "
        f"Unexpected public routes: {actual_public_routes - PUBLIC_ROUTES}, "
        f"Missing public routes: {PUBLIC_ROUTES - actual_public_routes}"
    )


def test_openapi_protected_routes_require_bearer_security(openapi_schema):
    """All domain endpoints must bind OAuth2PasswordBearer security."""
    paths = openapi_schema.get("paths", {})

    for path, methods in paths.items():
        for method, operation in methods.items():
            if method.lower() in {"get", "post", "put", "patch", "delete"}:
                if (path, method.lower()) not in PUBLIC_ROUTES:
                    security = operation.get("security", [])
                    assert any("OAuth2PasswordBearer" in req for req in security), (
                        f"Endpoint {method.upper()} {path} is missing OAuth2PasswordBearer security binding."
                    )


def test_openapi_obsolete_routes_are_absent(openapi_schema):
    """Stale legacy endpoints must never appear in the API schema."""
    paths = openapi_schema.get("paths", {})

    for obsolete_route in OBSOLETE_ROUTES:
        assert obsolete_route not in paths, f"Obsolete route {obsolete_route} found in OpenAPI paths!"


def test_openapi_signup_schema_prevents_privilege_escalation(openapi_schema):
    """InvitationSignupRequest must never accept role, department_id, or doctor_id."""
    schemas = openapi_schema.get("components", {}).get("schemas", {})
    assert "InvitationSignupRequest" in schemas, "InvitationSignupRequest missing from schemas"

    signup_props = set(schemas["InvitationSignupRequest"].get("properties", {}).keys())
    allowed_props = {"invitation_token", "first_name", "last_name", "email", "password"}

    assert signup_props == allowed_props
    assert "role" not in signup_props
    assert "department_id" not in signup_props
    assert "doctor_id" not in signup_props
