"""
Team controller functions for handling business logic related to team management.
"""

from fastapi import HTTPException, status

from src.team import repository


def create_team_controller(team_data, session):
    """Handles the business logic for creating a new team."""
    existing_team = repository.get_team_by_name(session, team_data.name)
    if existing_team:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"A team named '{team_data.name}' already exists."
        )
    return repository.create_team(session, team_data)


def list_teams_controller(session):
    """Handles logic for listing all active teams."""
    return repository.get_active_teams(session)


def get_team_controller(team_id, session):
    """Handles logic for fetching a specific team by its UUID."""
    team = repository.get_team_by_id(session, team_id)
    if not team or team.is_deleted:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Team not found."
        )
    return team
