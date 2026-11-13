"""
Team repository functions for handling database operations.
"""

import uuid
import datetime

from sqlmodel import Session, not_, select
from src.team.models import Team as TeamModel


def _sync_metadata(
    session: Session,
    team: TeamModel,
    *,
    commit: bool = False
) -> TeamModel:
    """Syncs metadata and commits or flushes"""
    team.sync_status = False
    team.updated_at = datetime.datetime.now(datetime.timezone.utc)

    session.add(team)
    session.flush()
    if commit:
        session.commit()
        session.refresh(team)

    return team


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
    team_id: uuid.UUID,
    department_id: uuid.UUID,
) -> TeamModel | None:
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
    name: uuid.UUID,
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


def restore_team(
    session: Session,
    team: TeamModel,
    *,
    commit: bool = False,
) -> TeamModel:
    """Restores a soft-deleted team"""
    team.is_deleted = False
    team = _sync_metadata(session, team, commit=commit)

    return team


def update_team(
    session: Session,
    team: TeamModel,
    name: str,
    *,
    commit: bool = False
) -> TeamModel:
    """Updates a team's name"""
    team.name = name
    team = _sync_metadata(session, team, commit=commit)

    return team


def soft_delete_team(
    session: Session,
    team: TeamModel,
    *,
    commit: bool = False,
) -> TeamModel:
    """Soft-deletes a team"""
    team.is_deleted = True
    team = _sync_metadata(session, team, commit=commit)

    return team
