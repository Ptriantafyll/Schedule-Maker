"""
Module: logging.py

This module sets up the logger that the backend uses
"""

from __future__ import annotations

import json
import os
from contextvars import ContextVar
from datetime import datetime, timezone
from logging import Formatter, LogRecord
from logging.config import dictConfig
from typing import Any
from fastapi import Request

request_id_var: ContextVar[str] = ContextVar("request_id", default="-")

_EXTRA_FIELDS = (
    "event",
    "method",
    "path",
    "status_code",
    "duration_ms",
    "error_count",
)

_LOG_LEVELS = frozenset(
    {
        "DEBUG",
        "INFO",
        "WARNING",
        "ERROR",
        "CRITICAL"
    }
)


class JsonFormatter(Formatter):
    """Format approved log-record fields as one JSON object per line"""

    def format(self, record: LogRecord) -> str:
        payload: dict[str, Any] = {
            "timestamp": (
                datetime.fromtimestamp(record.created, tz=timezone.utc)
                .isoformat(timespec="milliseconds")
                .replace("+00:00", "Z")
            ),
            "level": record.levelname,
            "logger": record.name,
            "message": record.getMessage(),
        }

        request_id = request_id_var.get()
        if request_id != "-":
            payload["request_id"] = request_id

        # Keep the schema explicit. Do not automatically copy arbitrary
        # `extra` values into logs, where secrets could be included by mistake
        for field in _EXTRA_FIELDS:
            value = getattr(record, field, None)
            if value is not None:
                payload[field] = value

        if record.exc_info:
            payload["exception"] = self.formatException(record.exc_info)

        # json.dumps escapes control characters, so an untrusted value cannot
        # turn one JSON event into multiple physical log lines
        return json.dumps(
            payload,
            ensure_ascii=False,
            default=str,
            separators=(",", ":"),
        )


def _configured_log_level() -> str:
    level = os.getenv("LOG_LEVEL", "INFO").upper()
    if level not in _LOG_LEVELS:
        raise ValueError(
            "LOG_LEVEL must be a standard logging level, e.g. DEBUG "
            "INFO, WARNING, ERROR, or CRITICAL"
        )
    return level


def configure_logging() -> None:
    """Configure handlers once for each application process."""

    level = _configured_log_level()

    dictConfig(
        {
            "version": 1,
            "disable_existing_loggers": False,
            "formatters": {
                "json": {
                    "()": JsonFormatter
                }
            },
            "handlers": {
                "stdout": {
                    "class": "logging.StreamHandler",
                    "level": level,
                    "formatter": "json",
                    "stream": "ext://sys.stdout"
                },
                # A real handler prevents logging's last-resort handler from
                # writing suppressed Ubicorn access records to stderr.
                "discard": {
                    "class": "logging.NullHandler",
                },
            },
            # Third-party libraries are noisy at INFO. keep their default at
            # WARNING while allowing the application package to use LOG_LEVEL.
            "root": {
                "level": "WARNING",
                "handlers": ["stdout"],
            },
            "loggers": {
                # This matches the `src.*` module names
                "src": {
                    "level": level,
                    "handlers": ["stdout"],
                    "propagate": False
                },
                "uvicorn": {
                    "level": level,
                    "handlers": ["stdout"],
                    "propagate": False
                },
                "uvicorn.error": {
                    "level": level,
                    "handlers": ["stdout"],
                    "propagate": False
                },
                "uvicorn.access": {
                    "level": "INFO",
                    "handlers": ["discard"],
                    "propagate": False
                },
                "sqlalchemy.engine.Engine": {
                    "level": "WARNING",
                    "handlers": ["discard"],
                    "propagate": False
                },
            },
        }
    )


def request_fields(
        request: Request,
        *,
        event: str,
        status_code: int | None = None,
        duration_ms: float | None = None,
) -> dict[str, object]:
    """Build safe, low-cardinality fields for an HTTP log event"""

    route = request.scope.get("route")
    path = getattr(route, "path", "<unmatched>")
    fields: dict[str, object] = {
        "event": event,
        "method": request.method,
        "path": path
    }

    if status_code is not None:
        fields["status_code"] = status_code
    if duration_ms is not None:
        fields["duration_ms"] = duration_ms

    return fields
