"""
Team routes for handling API requests related to team management.
"""

import uuid
from fastapi import APIRouter, Depends, status
from sqlmodel import Session
from src.db.connection import get_session

from src.team.schemas import TeamCreate, TeamRead
from src.team import controllers as team_controllers
from src.user.models import User as UserModel
from src.auth.dependencies import require_department_admin, require_department_member

router = APIRouter(
    prefix="/teams",
    tags=["Teams"]
)


@router.post("/", response_model=TeamRead, status_code=status.HTTP_201_CREATED)
def create_team(
    team_data: TeamCreate,
    session: Session = Depends(get_session),
    current_user: UserModel = Depends(require_department_admin),
):
    """Endpoint to create a new team. Placeholder implementation."""
    return team_controllers.create_team_controller(team_data, session)


@router.get("/", response_model=list[TeamRead])
def list_teams(
    session: Session = Depends(get_session),
    current_user: UserModel = Depends(require_department_member),
):
    """Endpoint to list all teams"""
    return team_controllers.list_teams_controller(session)


@router.get("/{team_id}", response_model=TeamRead)
def get_team(
    team_id: uuid.UUID,
    session: Session = Depends(get_session),
    current_user: UserModel = Depends(require_department_member),
):
    """Fetches a specific team by its UUID."""
    return team_controllers.get_team_controller(team_id, session)
