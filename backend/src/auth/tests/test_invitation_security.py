"""
Tests for user accoutn invitation security
"""

import re
import pytest

from src.auth import security


def test_generate_invitation_token_properties():
    """Tests that the genarate_invitation_token returns valid token"""
    token = security.generate_invitation_token()
    assert isinstance(token, str)
    assert len(token) >= 40
    assert re.fullmatch(r"^[A-Za-z0-9_-]+$", token) is not None


def test_generate_invitation_token_is_unique():
    """Tests that generating an invitation token creates unique token"""
    tokens = []
    for _ in range(1, 1000):
        tokens.append(security.generate_invitation_token())

    assert len(set(tokens)) == len(tokens)


def test_hash_invitation_token_is_deterministic():
    """Tests that hash_invitation_token is deterministic"""
    sample_token = "inv_test_token_12345"
    hash1 = security.hash_invitation_token(sample_token)
    hash2 = security.hash_invitation_token(sample_token)

    assert hash1 == hash2
    assert len(hash1) == 64
    assert re.fullmatch(r"^[0-9a-f]{64}$", hash1) is not None


def test_hash_invitation_token_different_inputs_produce_different_hashes():
    """Tests that different inputs return different tokens on hash_invitation_token"""
    token_a = "inv_test_token_12345"
    token_b = "inv_test_token_123456"
    hash_a = security.hash_invitation_token(token_a)
    hash_b = security.hash_invitation_token(token_b)

    assert hash_a != hash_b


def test_hash_invitation_token_normalizes_whitespace():
    """Tests that hash_invitation_token normalizes whitespaces"""
    sample_token = " inv_test_token_12345  "
    hash1 = security.hash_invitation_token(sample_token)
    hash2 = security.hash_invitation_token(sample_token.strip())

    assert hash1 == hash2


def test_hash_invitation_token_rejects_empty_token():
    """Tests that hashi_invitation_token rejects empty input"""
    with pytest.raises(ValueError):
        security.hash_invitation_token("")

    with pytest.raises(ValueError):
        security.hash_invitation_token(" ")
