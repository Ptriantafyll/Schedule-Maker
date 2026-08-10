"""
Module: connection.py
Description: This module sets up the database connection and session management for the application.
"""

import os
import logging
from collections.abc import Generator
from sqlmodel import Session, SQLModel, create_engine

logger = logging.getLogger(__name__)

# Default to local SQLite for development; swap to Postgres in production via env variables
DATABASE_URL = os.getenv("DATABASE_URL", "sqlite:///hospital_schedule.db")

# connect_args={"check_same_thread": False} is strictly required ONLY for SQLite
connect_args = {"check_same_thread": False} if DATABASE_URL.startswith(
    "sqlite") else {}

engine = create_engine(DATABASE_URL, echo=True, connect_args=connect_args)


def init_db() -> None:
    """
    Physically creates the tables in the target database if they do not exist.
    """
    from src.db.schemas import ScheduleConfig, Position, Shift  # pylint: disable=unused-import
    from src.department.models import Department  # pylint: disable=unused-import
    from src.team.models import Team  # pylint: disable=unused-import
    from src.doctor.models import Doctor  # pylint: disable=unused-import

    SQLModel.metadata.create_all(engine)
    logger.info("Database initialized")


def get_session() -> Generator[Session, None, None]:
    """Dependency provider for FastAPI routes to yield an isolated database session.

    Ensures that sessions are safely closed after a request finishes, preventing
    connection leaks.
    """
    with Session(engine) as session:
        yield session
