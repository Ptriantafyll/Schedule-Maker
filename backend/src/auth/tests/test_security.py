"""
Tests for the authentication security
"""

import uuid

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
        "token_type": "refresh",
        "iss": "untrusted-issuer",
        "aud": "untrusted-audience",
        "jti": "attacker-controlled"
    })
    decoded_token = security.decode_access_token(access_token)

    assert decoded_token["token_type"] == "access"
    jti = decoded_token["jti"]
    assert isinstance(jti, str)
    assert jti != "attacker-controlled"
