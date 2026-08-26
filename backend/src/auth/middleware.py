"""
Auth middleware
"""

from contextvars import ContextVar
from starlette.middleware.base import BaseHTTPMiddleware
from fastapi import Request
from src.auth.security import decode_access_token

current_user_var: ContextVar[dict| None] = ContextVar("current_user", default=None)

class AuthContextMiddleware(BaseHTTPMiddleware):
    """Auth context middleware"""
    async def dispatch(self, request: Request, call_next):
        auth_header = request.headers.get("Authorization")
        user_data = None

        if auth_header and auth_header.startswith("Bearer "):
            token = auth_header.split(" ", 1)[1]
            try:
                user_data = decode_access_token(token)
            except Exception:
                user_data = None

        request.state.user = user_data
        token_ctx = current_user_var.set(user_data)

        try:
            response = await call_next(request)
            return response
        finally:
            current_user_var.reset(token_ctx)