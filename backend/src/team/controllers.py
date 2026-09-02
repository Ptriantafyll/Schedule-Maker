"""
Team controller functions for handling business logic related to team management.
"""

import uuid
from fastapi import HTTPException, status
from sqlmodel import Session

from src.team import repository
from src.team.schemas import TeamCreate
from src.team.models import Team as TeamModel


def create_team_controller(
    team_data: TeamCreate,
    department_id: uuid.UUID,
    session: Session,
) -> TeamModel:
    """Handles the business logic for creating a new team."""
    existing_team = repository.get_team_by_name_for_department(
        session=session,
        name=team_data.name,
        department_id=department_id,
    )

    if existing_team:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"A team named '{team_data.name}' already exists."
        )

    return repository.create_team(
        session=session,
        name=team_data.name,
        department_id=department_id,
    )


def list_teams_controller(
    session: Session,
    department_id: uuid.UUID,
) -> list[TeamModel]:
    """Handles logic for listing all active teams."""
    return repository.get_active_teams_by_department(
        session=session,
        department_id=department_id,
    )


def get_team_controller(
    team_id: uuid.UUID,
    department_id: uuid.UUID,
    session: Session,
) -> TeamModel:
    """Handles logic for fetching a specific team by its UUID."""
    team = repository.get_team_by_id_for_department(
        session=session,
        team_id=team_id,
        department_id=department_id
    )
    if team is None:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Team not found."
        )
    return team
