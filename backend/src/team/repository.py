"""
Team repository functions for handling database operations.
"""

import uuid

from sqlmodel import Session, not_, select
from src.team.models import Team as TeamModel


def get_team_by_name_for_department(
        session: Session,
        name: str,
        department_id: uuid.UUID,
) -> TeamModel:
    """Retrieves a team by its unique name"""
    statement = select(TeamModel).where(
        TeamModel.name == name,
        TeamModel.department_id == department_id,
    )

    return session.exec(statement).first()


def get_team_by_id(session: Session, team_id: str) -> TeamModel:
    """Retrieves a specific team by its UUID"""
    return session.get(TeamModel, team_id)


def get_active_teams(session: Session) -> list[TeamModel]:
    """Retrieves all active (non-deleted) teams"""
    return list(session.exec(select(TeamModel).where(not_(TeamModel.is_deleted))).all())


def create_team(
    session: Session,
    name: str,
    department_id: uuid.UUID,
) -> TeamModel:
    """Creates a new team in the database."""
    new_team = TeamModel(
        name=name,
        department_id=department_id
    )
    session.add(new_team)
    session.commit()
    session.refresh(new_team)
    return new_team
