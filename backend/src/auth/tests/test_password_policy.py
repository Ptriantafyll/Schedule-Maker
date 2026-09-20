"""
Unit and integration tests for Phase 11 Step 1: Password Policy Engine.

Validates that:
- Strong passwords meeting all criteria pass validation.
- Passwords shorter than 10 characters are rejected.
- Passwords exceeding 72 UTF-8 bytes (bcrypt truncation boundary) are rejected.
- Passwords lacking uppercase, lowercase, digits, or special characters are rejected.
- Commonly guessed passwords in the blocklist are rejected.
- Public invitation signup enforces the password policy.
- Super-admin bootstrap enforces the password policy.
"""

import pytest
from sqlmodel import Session

from src.auth.security import (
    validate_password_strength,
    WeakPasswordError,
)
from src.auth.bootstrap import create_super_admin
from src.user.models import UserRole


# --------------------------------------------------------------------------
# 1. Unit Tests for validate_password_strength
# --------------------------------------------------------------------------

def test_strong_password_passes_validation():
    """A high-entropy password meeting all complexity rules passes without exception."""
    strong_passwords = [
        "CorrectHorseBattery99!",
        "Secure#Hospital2026",
        "MedScheduler_Doctor99",
        "P@ssw0rdIsNotSafeButThisIsLong123!",
    ]
    for pwd in strong_passwords:
        validate_password_strength(pwd)  # Should not raise


def test_password_too_short_is_rejected():
    """Passwords with fewer than 10 characters are rejected."""
    short_passwords = [
        "",
        "Ab1!",
        "Short1!",
        "Aa9!bcde",  # 8 chars
        "Aa9!bcdef",  # 9 chars
    ]
    for pwd in short_passwords:
        with pytest.raises(WeakPasswordError, match="at least 10 characters"):
            validate_password_strength(pwd)


def test_password_exceeding_72_bytes_is_rejected():
    """
    Passwords whose UTF-8 byte length exceeds 72 bytes are rejected
    to prevent bcrypt silent truncation vulnerability.
    """
    # 73 ASCII characters = 73 bytes
    too_long_ascii = "A1!" + "a" * 70
    assert len(too_long_ascii.encode("utf-8")) == 73
    with pytest.raises(WeakPasswordError, match="72 bytes"):
        validate_password_strength(too_long_ascii)

    # Multi-byte UTF-8 test: 4-byte emoji repeated
    # 18 emojis = 72 bytes (allowed), 19 emojis = 76 bytes (rejected)
    emoji_pwd_valid = "Aa1!" + "🔒" * 17  # 4 + 17*4 = 72 bytes
    assert len(emoji_pwd_valid.encode("utf-8")) == 72
    validate_password_strength(emoji_pwd_valid)

    emoji_pwd_invalid = "Aa1!" + "🔒" * 18  # 4 + 18*4 = 76 bytes
    assert len(emoji_pwd_invalid.encode("utf-8")) == 76
    with pytest.raises(WeakPasswordError, match="72 bytes"):
        validate_password_strength(emoji_pwd_invalid)


def test_password_missing_lowercase_is_rejected():
    """Passwords without at least one lowercase letter are rejected."""
    with pytest.raises(WeakPasswordError, match="lowercase letter"):
        validate_password_strength("ALLUPPERCASE123!")


def test_password_missing_uppercase_is_rejected():
    """Passwords without at least one uppercase letter are rejected."""
    with pytest.raises(WeakPasswordError, match="uppercase letter"):
        validate_password_strength("alllowercase123!")


def test_password_missing_digit_is_rejected():
    """Passwords without at least one digit are rejected."""
    with pytest.raises(WeakPasswordError, match="digit"):
        validate_password_strength("NoDigitsHereAtAll!")


def test_password_missing_special_character_is_rejected():
    """Passwords without at least one special symbol are rejected."""
    with pytest.raises(WeakPasswordError, match="special character"):
        validate_password_strength("NoSpecialChars1234")


def test_common_passwords_are_rejected():
    """Commonly guessed dictionary passwords are rejected even if they meet syntax rules."""
    blocked_passwords = [
        "Password123!",
        "Admin12345!",
        "Doctor1234!",
        "Welcome123!",
        "Hospital123!",
        "ChangeMe123!",
    ]
    for pwd in blocked_passwords:
        with pytest.raises(WeakPasswordError, match="common or easily guessable"):
            validate_password_strength(pwd)


# --------------------------------------------------------------------------
# 2. Integration: Public Invitation Signup Endpoint Enforces Policy
# --------------------------------------------------------------------------

def test_invitation_signup_rejects_weak_password(
    client, department_factory, user_factory, auth_headers_factory
):
    """Attempting to consume an invitation with a weak password returns 422 validation error."""
    dept = department_factory()
    admin = user_factory(role=UserRole.DEPARTMENT_ADMIN, department_id=dept.id)
    headers = auth_headers_factory(admin)

    # Admin issues an invitation
    inv_resp = client.post(
        "/api/v1/auth/invitations/staff",
        json={"role": "viewer"},
        headers=headers,
    )
    assert inv_resp.status_code == 201
    raw_token = inv_resp.json()["raw_token"]

    # Invitee attempts signup with weak passwords
    weak_payloads = [
        {"password": "short"},  # too short
        {"password": "nospecialchars123"},  # no upper, no special
        {"password": "Password123!"},  # blocked common password
    ]

    for wp in weak_payloads:
        signup_resp = client.post(
            "/api/v1/auth/signup",
            json={
                "invitation_token": raw_token,
                "first_name": "Test",
                "last_name": "Viewer",
                "email": "test_weak_pass@hospital.org",
                "password": wp["password"],
            },
        )
        assert signup_resp.status_code == 422, f"Expected 422 for password {wp['password']}"


# --------------------------------------------------------------------------
# 3. Integration: Super-Admin Bootstrap Enforces Policy
# --------------------------------------------------------------------------

def test_super_admin_bootstrap_rejects_weak_password(session: Session):
    """Bootstrapping a super-admin with a weak password raises WeakPasswordError."""
    with pytest.raises(WeakPasswordError):
        create_super_admin(
            session=session,
            email="root_weak@hospital.org",
            full_name="Root Admin",
            password="weak",
        )

    with pytest.raises(WeakPasswordError):
        create_super_admin(
            session=session,
            email="root_common@hospital.org",
            full_name="Root Admin",
            password="Password123!",
        )
