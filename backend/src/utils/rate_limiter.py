"""
In-memory sliding window rate limiter utility.
"""

import time
import math
import threading
import logging

from fastapi import HTTPException, status, Request

from src.utils import logger


class InMemoryRateLimiter:
    """Rate limiter for keys (client identifiers)."""

    def __init__(self):
        self._records: dict[str, list[float]] = {}
        self._lock = threading.Lock()

    def is_allowed(self, key: str, max_requests: int, window_seconds: int) -> tuple[bool, int]:
        """
        Determines if a request for the given key is permitted within the sliding window.

        Args:
            key: Client identifier (e.g. IP address or user ID).
            max_requests: Maximum number of allowed requests within the time window.
            window_seconds: Duration of the sliding window in seconds.

        Returns:
            A tuple of (allowed: bool, retry_after: int). If allowed, retry_after is 0.
            If blocked, retry_after indicates seconds until a request slot opens up.
        """
        with self._lock:
            now = time.time()
            cutoff = now - window_seconds

            existing_timestamps = self._records.get(key, [])
            valid_timestamps = [t for t in existing_timestamps if t > cutoff]

            if len(valid_timestamps) >= max_requests:
                retry_after = int(
                    math.ceil(valid_timestamps[0] + window_seconds - now)
                )
                return (False, max(1, retry_after))

            valid_timestamps.append(now)
            self._records[key] = valid_timestamps
            return (True, 0)

    def reset(self) -> None:
        """Clears all stored rate limit history."""
        with self._lock:
            self._records = {}

    def prune_stale(self, older_than_seconds: int = 300) -> int:
        """Deletes expired keys and returns the number of keys removed."""
        with self._lock:
            now = time.time()
            cutoff = now - older_than_seconds

            stale_keys = []
            for key, timestamps in self._records.items():
                if not timestamps or max(timestamps) <= cutoff:
                    stale_keys.append(key)

            for key in stale_keys:
                del self._records[key]

            return len(stale_keys)


auth_rate_limiter = InMemoryRateLimiter()


def get_client_ip(request: Request) -> str:
    """Extracts the client ip from headers or connection."""
    forwarded = request.headers.get("X-Forwarded-For")
    if forwarded:
        return forwarded.split(",")[0].strip()
    if request.client and request.client.host:
        return request.client.host
    return "127.0.0.1"


def rate_limit(max_requests: int, window_seconds: int, endpoint_name: str):
    """Factory returning a FastAPI dependency for endpoint rate limiting."""
    def dependency(request: Request = None) -> None:
        client_ip = get_client_ip(request)
        key = f"{client_ip}:{endpoint_name}"
        allowed, retry_after = auth_rate_limiter.is_allowed(
            key=key,
            max_requests=max_requests,
            window_seconds=window_seconds,
        )

        if not allowed:
            logger.log_audit_event(
                action="auth.rate_limited",
                outcome="failure",
                message=f"Rate limit exceeded for endpoint '{endpoint_name}' from IP {client_ip}",
                reason="too_many_requests",
                level=logging.WARNING,
            )

            raise HTTPException(
                status_code=status.HTTP_429_TOO_MANY_REQUESTS,
                detail="Too many requests. Please try again later.",
                headers={"Retry-After": str(retry_after)},
            )

    return dependency
