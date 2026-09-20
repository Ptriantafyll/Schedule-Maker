"""
Unit tests for Phase 11 Step 2: In-Memory Sliding Window Rate Limiter.

Validates that:
- Requests under the limit are allowed.
- Requests exceeding max_requests within window_seconds are blocked with a positive retry_after.
- Expired timestamps slide out of the window, allowing subsequent requests.
- Different client keys (e.g. different IP addresses) are tracked independently.
- Reset clears stored history for clean test isolation.
- Prune removes stale keys to bound memory growth.
- Concurrent requests from multiple threads are handled safely without race conditions.
"""

from unittest.mock import patch
import threading
import pytest

from src.utils.rate_limiter import InMemoryRateLimiter


@pytest.fixture(name="limiter")
def limiter_fixture():
    """Provides a fresh rate limiter instance for each test."""
    return InMemoryRateLimiter()


def test_first_request_is_allowed(limiter):
    """The very first request for a key is always allowed."""
    allowed, retry_after = limiter.is_allowed("client-1", max_requests=3, window_seconds=60)
    assert allowed is True
    assert retry_after == 0


def test_requests_up_to_max_are_allowed(limiter):
    """Requests up to max_requests are allowed."""
    key = "client-ip-127.0.0.1"
    for i in range(3):
        allowed, retry_after = limiter.is_allowed(key, max_requests=3, window_seconds=60)
        assert allowed is True, f"Request {i + 1} should have been allowed"
        assert retry_after == 0


def test_request_exceeding_max_is_rejected_with_retry_after(limiter):
    """The (max_requests + 1)th request in the same window is rejected."""
    key = "client-ip-127.0.0.1"
    # Consume 3 allowed requests
    for _ in range(3):
        limiter.is_allowed(key, max_requests=3, window_seconds=60)

    # 4th request must be rejected
    allowed, retry_after = limiter.is_allowed(key, max_requests=3, window_seconds=60)
    assert allowed is False
    assert 0 < retry_after <= 60


def test_window_expiration_slides_out_old_requests(limiter):
    """When time advances beyond the window, old timestamps slide out and allow new requests."""
    key = "sliding-client"
    base_time = 1000.0

    with patch("time.time", return_value=base_time):
        # 2 requests at t=1000
        limiter.is_allowed(key, max_requests=2, window_seconds=10)
        limiter.is_allowed(key, max_requests=2, window_seconds=10)
        # 3rd request rejected at t=1000
        allowed, retry_after = limiter.is_allowed(key, max_requests=2, window_seconds=10)
        assert allowed is False
        assert retry_after == 10

    # Advance time by 6 seconds (t=1006): still inside 10-second window
    with patch("time.time", return_value=base_time + 6):
        allowed, retry_after = limiter.is_allowed(key, max_requests=2, window_seconds=10)
        assert allowed is False
        assert retry_after == 4

    # Advance time by 11 seconds (t=1011): original requests have expired!
    with patch("time.time", return_value=base_time + 11):
        allowed, retry_after = limiter.is_allowed(key, max_requests=2, window_seconds=10)
        assert allowed is True
        assert retry_after == 0


def test_different_keys_are_isolated(limiter):
    """Rate limiting one client key does not impact other client keys."""
    client_a = "192.168.1.10"
    client_b = "192.168.1.20"

    # Exhaust limit for Client A
    for _ in range(2):
        limiter.is_allowed(client_a, max_requests=2, window_seconds=60)

    # Client A should be blocked
    allowed_a, _ = limiter.is_allowed(client_a, max_requests=2, window_seconds=60)
    assert allowed_a is False

    # Client B should still be allowed
    allowed_b, retry_b = limiter.is_allowed(client_b, max_requests=2, window_seconds=60)
    assert allowed_b is True
    assert retry_b == 0


def test_reset_clears_stored_limits(limiter):
    """Calling reset() wipes the recorded rate limits."""
    key = "blocked-client"
    limiter.is_allowed(key, max_requests=1, window_seconds=60)
    assert limiter.is_allowed(key, max_requests=1, window_seconds=60)[0] is False

    limiter.reset()

    # After reset, the client can request again immediately
    allowed, retry_after = limiter.is_allowed(key, max_requests=1, window_seconds=60)
    assert allowed is True
    assert retry_after == 0


def test_prune_stale_cleans_up_memory(limiter):
    """Stale keys whose timestamps are all expired get evicted from memory."""
    base_time = 1000.0

    with patch("time.time", return_value=base_time):
        limiter.is_allowed("stale-1", max_requests=5, window_seconds=30)
        limiter.is_allowed("stale-2", max_requests=5, window_seconds=30)

    # Advance time far into future (t=2000, 1000s later)
    with patch("time.time", return_value=base_time + 1000):
        # Add an active key at t=2000
        limiter.is_allowed("active-key", max_requests=5, window_seconds=30)

        # Prune keys older than 60 seconds
        pruned_count = limiter.prune_stale(older_than_seconds=60)
        assert pruned_count == 2


def test_thread_safety_under_concurrent_access(limiter):
    """Concurrent requests across multiple threads do not cause race conditions or corrupt state."""
    key = "concurrent-client"
    max_requests = 50
    results = []

    def make_request():
        allowed, _ = limiter.is_allowed(key, max_requests=max_requests, window_seconds=10)
        results.append(allowed)

    threads = [threading.Thread(target=make_request) for _ in range(100)]
    for t in threads:
        t.start()
    for t in threads:
        t.join()

    # Exactly 50 requests must be allowed, and 50 must be blocked
    assert results.count(True) == max_requests
    assert results.count(False) == 50
