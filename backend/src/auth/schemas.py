"""
Pydantic schemas for authentication
"""

import datetime
from typing import Optional
from pydantic import BaseModel
from src.user.models import UserRole


class Token(BaseModel):
    """Token class"""
    access_token: str
    token_type: str = "bearer"


class TokenPayload(BaseModel):
    """Token payload class"""
    sub: str  # user.id as string
    email: str
    role: UserRole
    department_id: Optional[str] = None
    exp: datetime.datetime
