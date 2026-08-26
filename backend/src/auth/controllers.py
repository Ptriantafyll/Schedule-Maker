"""
Authentication controller functions handling business logic
"""

from fastapi import HTTPException, status, Depends
from sqlmodel import Session
from src.user import repository as user_repository
from src.user.schemas import UserLogin
from src.auth.schemas import Token
from src.auth.security import verify_password, create_access_token
from src.db.connection import get_session

