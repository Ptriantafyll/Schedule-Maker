"""
Team controller functions for handling business logic related to team management.
"""

import uuid
from fastapi import HTTPException, status
from sqlmodel import Session

from src.team import repository
from src.team.schemas import TeamCreate, TeamUpdate
from src.team.models import Team as TeamModel
from src.doctor import repository as doctor_repository


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
        if not existing_team.is_deleted:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail=f"A team named '{team_data.name}' already exists."
            )

        return repository.restore_team(
            session=session,
            team=existing_team,
            commit=True
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


def update_team_controller(
    team_id: uuid.UUID,
    team_data: TeamUpdate,
    department_id: uuid.UUID,
    session: Session,
) -> TeamModel:
    """Handles logic for updating a specific team."""
    team_to_update = repository.get_team_by_id_for_department(
        session, team_id, department_id
    )

    if team_to_update is None:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Team not found."
        )

    update_data = team_data.model_dump(exclude_unset=True)
    new_name = update_data.get("name")

    if new_name is None or new_name == team_to_update.name:
        return team_to_update

    existing_team = repository.get_team_by_name_for_department(
        session=session,
        name=new_name,
        department_id=department_id,
    )

    if existing_team and existing_team.id != team_to_update.id:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"A team named '{new_name}' already exists.",
        )

    return repository.update_team(
        session=session,
        team=team_to_update,
        name=new_name,
        commit=True
    )


def delete_team_controller(
    team_id: uuid.UUID,
    department_id: uuid.UUID,
    session: Session,
) -> None:
    """Handles logic for soft-deleting a specific team."""
    team_to_delete = repository.get_team_by_id_for_department(
        session, team_id, department_id
    )

    if team_to_delete is None:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Team not found."
        )
    doctor_repository.clear_team_from_doctors_for_department(
        session=session,
        team_id=team_id,
        department_id=department_id,
        commit=False
    )

    repository.soft_delete_team(
        session, team_to_delete, commit=False
    )

    session.commit()

    return None
