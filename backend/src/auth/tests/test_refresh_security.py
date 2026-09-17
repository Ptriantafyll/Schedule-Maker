"""
Unit tests for refresh token and CSRF token security helpers.
"""

import re
import pytest

from src.auth import security


# ==============================================================================
# Refresh Token Tests
# ==============================================================================

def test_generate_refresh_token_properties():
    """Generates a high-entropy URL-safe base64 string of sufficient length."""
    token = security.generate_refresh_token()
    assert isinstance(token, str)
    assert len(token) >= 40
    assert re.fullmatch(r"^[A-Za-z0-9_-]+$", token) is not None


def test_generate_refresh_token_is_unique():
    """Generates unique refresh tokens across 1,000 samples (zero collisions)."""
    tokens = [security.generate_refresh_token() for _ in range(1000)]
    assert len(set(tokens)) == len(tokens)


def test_hash_refresh_token_is_deterministic():
    """Hashing produces a 64-character lowercase hex string deterministically."""
    sample_token = "refresh_sample_token_xyz_12345"
    hash1 = security.hash_refresh_token(sample_token)
    hash2 = security.hash_refresh_token(sample_token)

    assert hash1 == hash2
    assert len(hash1) == 64
    assert re.fullmatch(r"^[0-9a-f]{64}$", hash1) is not None


def test_hash_refresh_token_different_inputs_produce_different_hashes():
    """Different raw tokens produce distinct SHA-256 digests."""
    token_a = "refresh_token_alpha"
    token_b = "refresh_token_beta"
    assert security.hash_refresh_token(token_a) != security.hash_refresh_token(token_b)


def test_hash_refresh_token_normalizes_whitespace():
    """Leading and trailing whitespace is stripped before hashing."""
    raw = "refresh_token_padded"
    padded = f"   {raw}   \n"
    assert security.hash_refresh_token(padded) == security.hash_refresh_token(raw)


@pytest.mark.parametrize("invalid_input", ["", "   ", "\t\n", None])
def test_hash_refresh_token_rejects_empty_or_whitespace(invalid_input):
    """Empty, whitespace-only, or None inputs raise ValueError."""
    with pytest.raises(ValueError):
        security.hash_refresh_token(invalid_input)


# ==============================================================================
# CSRF Token Tests
# ==============================================================================

def test_generate_csrf_token_properties():
    """Generates a high-entropy URL-safe base64 string of sufficient length."""
    token = security.generate_csrf_token()
    assert isinstance(token, str)
    assert len(token) >= 40
    assert re.fullmatch(r"^[A-Za-z0-9_-]+$", token) is not None


def test_generate_csrf_token_is_unique():
    """Generates unique CSRF tokens across 1,000 samples (zero collisions)."""
    tokens = [security.generate_csrf_token() for _ in range(1000)]
    assert len(set(tokens)) == len(tokens)


def test_hash_csrf_token_is_deterministic():
    """Hashing produces a 64-character lowercase hex string deterministically."""
    sample_csrf = "csrf_sample_token_abc_67890"
    hash1 = security.hash_csrf_token(sample_csrf)
    hash2 = security.hash_csrf_token(sample_csrf)

    assert hash1 == hash2
    assert len(hash1) == 64
    assert re.fullmatch(r"^[0-9a-f]{64}$", hash1) is not None


def test_hash_csrf_token_different_inputs_produce_different_hashes():
    """Different raw CSRF tokens produce distinct SHA-256 digests."""
    csrf_a = "csrf_token_one"
    csrf_b = "csrf_token_two"
    assert security.hash_csrf_token(csrf_a) != security.hash_csrf_token(csrf_b)


def test_hash_csrf_token_normalizes_whitespace():
    """Leading and trailing whitespace is stripped before hashing."""
    raw = "csrf_token_padded"
    padded = f"   {raw}  \t"
    assert security.hash_csrf_token(padded) == security.hash_csrf_token(raw)


@pytest.mark.parametrize("invalid_input", ["", "   ", "\t\n", None])
def test_hash_csrf_token_rejects_empty_or_whitespace(invalid_input):
    """Empty, whitespace-only, or None inputs raise ValueError."""
    with pytest.raises(ValueError):
        security.hash_csrf_token(invalid_input)
