"""
Route tests for auth
"""

import pytest

from src.auth import security

from src.user.schemas import UserCreate
from src.user.models import UserRole
from src.user import controllers as user_controllers

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
        "wrong-user"
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

    assert response.status_code == 201
    assert response.json() == {"detail": "Username or password is incorrect"}
    assert response.headers.get("WWW-Authenticate") == "Bearer"
