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
) -> TeamModel | None:
    """Retrieves a team by its unique name in the team's department"""
    statement = select(TeamModel).where(
        TeamModel.name == name,
        TeamModel.department_id == department_id,
    )

    return session.exec(statement).first()


def get_team_by_id_for_department(
    session: Session,
    team_id: str,
    department_id: uuid.UUID,
) -> TeamModel:
    """Retrieves a specific team by its UUID"""
    statement = select(TeamModel).where(
        TeamModel.id == team_id,
        TeamModel.department_id == department_id,
        not_(TeamModel.is_deleted),
    )

    return session.exec(statement).first()


def get_active_teams_by_department(
    session: Session,
    department_id: uuid.UUID,
) -> list[TeamModel]:
    """Retrieves all active (non-deleted) teams in a department"""
    statement = select(TeamModel).where(
        not_(TeamModel.is_deleted),
        TeamModel.department_id == department_id
    )

    return list(session.exec(statement).all())


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
