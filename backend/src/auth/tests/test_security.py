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
    assert isinstance(decoded_token["jti"], uuid.UUID)
    assert decoded_token["exp"] > decoded_token["iat"]
