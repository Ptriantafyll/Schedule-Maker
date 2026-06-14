"""
Team repository functions for handling database operations.
"""


from sqlmodel import Session
from src.team.schemas import TeamCreate
from src.team.models import Team as TeamModel


def create_team(session: Session, team_data: TeamCreate) -> TeamModel:
    """Creates a new team in the database."""
    new_team = TeamModel(
        name=team_data.name,
        department_id=team_data.department_id
    )
    session.add(new_team)
    session.commit()
    session.refresh(new_team)
    return new_team
