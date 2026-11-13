"""
Position repository functions for handling database operations.
"""

# from typing import Optional
import uuid
import datetime
from sqlmodel import Session, not_, select

from src.position.models import Position as PositionModel


def _sync_metadata(
    session: Session,
    position: PositionModel,
    *,
    commit: bool = False
) -> PositionModel:
    """Syncs metadata and commits or flushes"""
    position.sync_status = False
    position.updated_at = datetime.datetime.now(datetime.timezone.utc)

    session.add(position)
    session.flush()
    if commit:
        session.commit()
        session.refresh(position)

    return position


def create_position(
    session: Session,
    position_name: str,
    duty_days: list[int],
    department_id: uuid.UUID,
) -> PositionModel:
    """Creates a new position in the database"""
    new_position = PositionModel(
        name=position_name,
        department_id=department_id,
        duty_days=duty_days,
    )

    session.add(new_position)
    session.commit()
    session.refresh(new_position)
    return new_position


def stage_position(
    session: Session,
    position_name: str,
    duty_days: list[int],
    department_id: uuid.UUID
) -> PositionModel:
    """Stages a new position, but does not commit in the db"""
    new_position = PositionModel(
        name=position_name,
        department_id=department_id,
        duty_days=duty_days
    )

    session.add(new_position)
    session.flush()
    return new_position


def get_position_by_id_for_department(
    session: Session,
    position_id: uuid.UUID,
    department_id: uuid.UUID,
) -> PositionModel | None:
    """Retrieves a position by its id"""
    statement = select(PositionModel).where(
        PositionModel.id == position_id,
        PositionModel.department_id == department_id,
        not_(PositionModel.is_deleted),
    )
    return session.exec(statement).first()


def get_position_by_name_for_department(
    session: Session,
    position_name: str,
    department_id: uuid.UUID,
) -> PositionModel | None:
    """Retrieves a position by its id"""
    statement = select(PositionModel).where(
        PositionModel.name == position_name,
        PositionModel.department_id == department_id,
    )

    return session.exec(statement).first()


def get_active_positions_for_department(
    session: Session,
    department_id: uuid.UUID,
) -> list[PositionModel]:
    """Retrieves all active (non deleted) positions"""
    statement = select(PositionModel).where(
        not_(PositionModel.is_deleted),
        PositionModel.department_id == department_id
    )

    return session.exec(statement).all()


def restore_position(
    session: Session,
    position: PositionModel,
    duty_days: list[int],
    *,
    commit: bool = False,
) -> PositionModel:
    """Restores a soft-deleted position"""
    position.is_deleted = False
    position.duty_days = duty_days
    position = _sync_metadata(session, position, commit=commit)

    return position


def update_position(
    session: Session,
    position: PositionModel,
    name: str = None,
    duty_days: list[int] = None,
    *,
    commit: bool = False
) -> PositionModel:
    """Updates a position."""
    if name is not None:
        position.name = name

    if duty_days is not None:
        position.duty_days = duty_days

    position = _sync_metadata(session, position, commit=commit)

    return position


def soft_delete_position(
    session: Session,
    position: PositionModel,
    *,
    commit: bool = False,
) -> PositionModel:
    """Soft-deletes a position."""
    position.is_deleted = True
    position = _sync_metadata(session, position, commit=commit)

    return position
