"""
Tests for the team module
"""

import uuid
import datetime
import pytest
from sqlalchemy.pool import StaticPool
from sqlmodel import SQLModel, create_engine, Session

from src.team.schemas import TeamCreate
from src.team.repository import create_team
from src.department.schemas import DepartmentCreate
from src.department.repository import create_department


@pytest.fixture(name="session")
def session_fixture():
    """Creates a fresh in-memory database session for each test."""
    engine = create_engine(
        "sqlite:///:memory:",
        connect_args={"check_same_thread": False},
        poolclass=StaticPool,
    )
    SQLModel.metadata.create_all(engine)
    with Session(engine) as session:
        yield session


################################
# Repository tests
################################
def test_create_team(session):
    """Test creating a team and verifying its fields."""

    dept_data = DepartmentCreate(name="Cardiology", code="CARD")
    new_dept = create_department(session, dept_data)

    team_data = TeamCreate(name="ER Team A", department_id=new_dept.id)
    new_team = create_team(session, team_data)

    assert isinstance(new_team.id, uuid.UUID)
    assert new_team.name == "ER Team A"
    assert new_team.department_id == new_dept.id
    assert new_team.is_deleted is False
    assert new_team.sync_status is False
    assert isinstance(new_team.created_at, datetime.datetime)
    assert isinstance(new_team.updated_at, datetime.datetime)
