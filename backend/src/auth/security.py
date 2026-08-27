"""
Security utils for authentication
"""

import os
import datetime
from typing import Optional
import jwt
import bcrypt
import uuid

SECRET_KEY = os.getenv(
    "SECRET_KEY", "dev-secret-key-change-in-production-123456")
ALGORITHM = os.getenv("ALGORITHM", "HS256")
ACCESS_TOKEN_EXPIRE_MINUTES = int(
    os.getenv("ACCESS_TOKEN_EXPIRE_MINUTES", "60"))
ISSUER = "schedule-maker-api"
AUDIENCE = "schedule-maker-clients"
TOKEN_TYPE = "access"
REQUIRED_ACCESS_TOKEN_CLAIMS = (
    "sub",
    "exp",
    "iat",
    "jti",
    "iss",
    "aud",
    "token_type"
)


def hash_password(plain_password: str) -> str:
    """Hash a plaintext password with bcrypt."""
    salt = bcrypt.gensalt()
    return bcrypt.hashpw(plain_password.encode("utf-8"), salt).decode("utf-8")


def verify_password(plain_password: str, hashed_password) -> bool:
    """Verify a plaintext password against a stored bcrypt hash."""
    return bcrypt.checkpw(plain_password.encode("utf-8"), hashed_password.encode("utf-8"))


def create_access_token(data: dict, expires_delta: Optional[datetime.timedelta] = None) -> str:
    """Creates a signed JWT access token."""
    to_encode = data.copy()
    now = datetime.datetime.now(datetime.timezone.utc)
    expire = now + (
        expires_delta
        if expires_delta is not None
        else datetime.timedelta(
            minutes=ACCESS_TOKEN_EXPIRE_MINUTES)
    )

    to_encode.update({
        "exp": expire,
        "iat": now,
        "jti": str(uuid.uuid4()),
        "iss": ISSUER,
        "aud": AUDIENCE,
        "token_type": TOKEN_TYPE
    })

    return jwt.encode(to_encode, SECRET_KEY, algorithm=ALGORITHM)


def decode_access_token(token: str) -> dict:
    """Decode and validate a JWT access token."""
    payload = jwt.decode(
        token,
        SECRET_KEY,
        algorithms=[ALGORITHM],
        issuer=ISSUER,
        audience=AUDIENCE,
        options={"require": REQUIRED_ACCESS_TOKEN_CLAIMS}
    )

    if payload["token_type"] != TOKEN_TYPE:
        raise jwt.InvalidTokenError("Invalid token type")

    return payload
