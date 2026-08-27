"""
Tests for the authentication security
"""

import uuid
import datetime
import jwt
import pytest

from src.auth import security


def test_access_token_contains_required_claims():
    """Tests that an access token contains the required claims"""
    new_uuid = uuid.uuid4()
    access_token = security.create_access_token({"sub": str(new_uuid)})
    decoded_token = security.decode_access_token(access_token)

    assert decoded_token["sub"] == str(new_uuid)
    assert decoded_token["token_type"] == "access"
    assert decoded_token["iss"] == "schedule-maker-api"
    assert decoded_token["aud"] == "schedule-maker-clients"
    jti = decoded_token["jti"]
    assert isinstance(jti, str)
    assert str(uuid.UUID(jti)) == jti
    assert decoded_token["exp"] > decoded_token["iat"]


def test_access_token_security_claims_cannot_be_overridden():
    """Tests that security claims of a token cannot be overridden"""
    access_token = security.create_access_token({
        "sub": str(uuid.uuid4()),
        "token_type": "refresh",
        "iss": "untrusted-issuer",
        "aud": "untrusted-audience",
        "jti": "attacker-controlled"
    })
    decoded_token = security.decode_access_token(access_token)

    assert decoded_token["token_type"] == "access"
    assert decoded_token["iss"] == "schedule-maker-api"
    assert decoded_token["aud"] == "schedule-maker-clients"
    jti = decoded_token["jti"]
    assert isinstance(jti, str)
    assert jti != "attacker-controlled"
    assert str(uuid.UUID(jti)) == jti


def test_decode_access_token_rejects_expired_token():
    """Tests that an expired token is rejected"""
    sub = str(uuid.uuid4())
    access_token = security.create_access_token(
        {"sub": sub},
        expires_delta=datetime.timedelta(seconds=-1)
    )

    with pytest.raises(jwt.ExpiredSignatureError):
        security.decode_access_token(access_token)


@pytest.mark.parametrize(
    "missing_claim",
    [
        "sub",
        "exp",
        "iat",
        "jti",
        "iss",
        "aud",
        "token_type",
    ]
)
def test_decode_access_token_rejects_missing_sub(missing_claim):
    """Tests missing sub"""
    access_token = security.create_access_token(
        {"sub": str(uuid.uuid4())}
    )
    payload = security.decode_access_token(access_token)
    payload.pop(missing_claim)

    token_missing_claim = jwt.encode(
        payload,
        security.SECRET_KEY,
        algorithm=security.ALGORITHM
    )

    with pytest.raises(jwt.MissingRequiredClaimError) as exc_info:
        security.decode_access_token(access_token)

    assert exc_info.value.claim == missing_claim
