"""
Route tests for auth
"""

import pytest
import datetime

from src.auth import security
from src.user.schemas import UserCreate
from src.user.models import UserRole
from src.user import controllers as user_controllers
from src.main import app

LOGIN_EMAIL = "test@test.com"
LOGIN_PASSWORD = "password123"


#####################
# Fixtures
#####################

@pytest.fixture(name="login_user")
def login_user_fixture(session):
    """Creates a reusable login user for tests"""
    user_data = UserCreate(
        email=LOGIN_EMAIL,
        password=LOGIN_PASSWORD,
        full_name="Test super admin",
        role=UserRole.SUPER_ADMIN,
        doctor_id=None,
        department_id=None
    )

    return user_controllers.create_user_controller(user_data, session)


@pytest.fixture(name="login_user_headers")
def login_user_headers_fixture(login_user):
    """Creates reusable login user headers for tests"""
    access_token = security.create_access_token({"sub": str(login_user.id)})

    return {"Authorization": f"Bearer {access_token}"}

#####################
# Tests
#####################


def test_login_returns_access_token(client, login_user):
    """Tests that logging in returns an access token"""
    response = client.post(
        "/api/v1/auth/login",
        data={
            "username": login_user.email,
            "password": LOGIN_PASSWORD
        }
    )

    assert response.status_code == 200
    data = response.json()
    assert data["token_type"] == "bearer"
    assert isinstance(data["access_token"], str)
    assert data["access_token"]

    payload = security.decode_access_token(data["access_token"])

    assert payload["sub"] == str(login_user.id)
    assert payload["token_type"] == security.TOKEN_TYPE


@pytest.mark.parametrize(
    "email, password",
    [
        (LOGIN_EMAIL, "wrong-password"),
        ("unknown@test.com", LOGIN_PASSWORD)
    ],
    ids=[
        "wrong-password",
        "unknown-user"
    ],
)
def test_login_rejects_invalid_credentials(client, login_user, email, password):
    """Tests logging in with invalid credentials"""
    response = client.post(
        "/api/v1/auth/login",
        data={
            "username": email,
            "password": password
        }
    )

    assert response.status_code == 401
    assert response.json() == {"detail": "Username or password is incorrect"}
    assert response.headers.get("WWW-Authenticate") == "Bearer"


def test_login_rejects_deleted_user(session, client, login_user):
    """Tests logging in with credentials of a deleted user"""
    login_user.is_deleted = True
    session.add(login_user)
    session.commit()

    response = client.post(
        "/api/v1/auth/login",
        data={
            "username": login_user.email,
            "password": LOGIN_PASSWORD
        }
    )

    assert response.status_code == 401
    assert response.json() == {"detail": "Username or password is incorrect"}
    assert response.headers.get("WWW-Authenticate") == "Bearer"


def test_legacy_login_route_is_not_exposed():
    """Tests that the old /api/v1/users/login route no longer exists"""

    assert "/api/v1/users/login" not in app.openapi()["paths"]


def test_get_current_user_profile_returns_safe_user(client, login_user, login_user_headers):
    """Tests /api/v1/auth/me"""
    response = client.get(
        "/api/v1/auth/me",
        headers=login_user_headers,
    )

    assert response.status_code == 200
    data = response.json()
    assert data["id"] == str(login_user.id)
    assert data["email"] == login_user.email
    assert data["full_name"] == login_user.full_name
    assert data["role"] == UserRole.SUPER_ADMIN.value
    assert "hashed_password" not in data


def test_get_current_user_profile_requires_authentication(client):
    """Tests /api/v1/auth/me without auth"""
    response = client.get(
        "/api/v1/auth/me",
    )

    assert response.status_code == 401
    assert response.json() == {"detail": "Unauthorized"}
    assert response.headers.get("WWW-Authenticate") == "Bearer"


def test_current_user_profile_rejects_invalid_subject(client):
    """Tests get /api/v1/auth/me with invalid subject"""
    access_token = security.create_access_token({"sub": "not-a-uuid"})

    response = client.get(
        "/api/v1/auth/me",
        headers={"Authorization": f"Bearer {access_token}"}
    )

    assert response.status_code == 401
    assert response.json() == {"detail": "Unauthorized"}
    assert response.headers.get("WWW-Authenticate") == "Bearer"


def test_current_user_profile_rejects_malformed_token(client):
    """Tests get /api/v1/auth/me with malformed token"""
    response = client.get(
        "/api/v1/auth/me",
        headers={"Authorization": "Bearer invalid-token"}
    )

    assert response.status_code == 401
    assert response.json() == {"detail": "Unauthorized"}
    assert response.headers.get("WWW-Authenticate") == "Bearer"


def test_current_user_profile_rejects_expired_token(client, login_user):
    """Tests get /api/v1/auth/me with malformed token"""
    access_token = security.create_access_token(
        {"sub": str(login_user.id)},
        expires_delta=datetime.timedelta(seconds=-1),
    )

    response = client.get(
        "/api/v1/auth/me",
        headers={"Authorization": "Bearer invalid-token"}
    )

    assert response.status_code == 401
    assert response.json() == {"detail": "Unauthorized"}
    assert response.headers.get("WWW-Authenticate") == "Bearer"
