"""
Team routes for handling API requests related to team management.
"""

import uuid
from fastapi import APIRouter, Depends, status
from sqlmodel import Session
from src.db.connection import get_session

from src.team.schemas import TeamCreate, TeamRead
from src.team.controllers import create_team_controller, list_teams_controller, get_team_controller


router = APIRouter(
    prefix="/teams",
    tags=["Teams"]
)


@router.post("/", response_model=TeamRead, status_code=status.HTTP_201_CREATED)
def create_team(team_data: TeamCreate, session: Session = Depends(get_session)):
    """Endpoint to create a new team. Placeholder implementation."""
    return create_team_controller(team_data, session)


@router.get("/", response_model=list[TeamRead])
def list_teams(session: Session = Depends(get_session)):
    """Endpoint to list all teams"""
    return list_teams_controller(session)


@router.get("/{team_id}", response_model=TeamRead)
def get_team(team_id: uuid.UUID, session: Session = Depends(get_session)):
    """Fetches a specific team by its UUID."""
    return get_team_controller(team_id, session)
