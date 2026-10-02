"""
Schedule repository module for database operations on schedule drafts.
"""

from typing import Optional
import uuid
from sqlmodel import Session, select, not_, desc
from src.schedule.models import ScheduleDraft


def create_schedule_draft(session: Session, draft: ScheduleDraft) -> ScheduleDraft:
    """
    Persists a new ScheduleDraft into the database.
    """
    session.add(draft)
    session.commit()
    session.refresh(draft)
    return draft


def get_schedule_draft_by_id(
    session: Session,
    draft_id: uuid.UUID,
) -> Optional[ScheduleDraft]:
    """
    Retrieves an active (non-deleted) ScheduleDraft by its UUID.
    """
    draft = session.get(ScheduleDraft, draft_id)
    if draft and not draft.is_deleted:
        return draft
    return None


def get_schedule_draft_for_department_month(
    session: Session,
    department_id: uuid.UUID,
    target_month: str,
) -> Optional[ScheduleDraft]:
    """
    Retrieves the most recent active ScheduleDraft for a given department and target month.
    """
    statement = (
        select(ScheduleDraft)
        .where(
            ScheduleDraft.department_id == department_id,
            ScheduleDraft.target_month == target_month,
            not_(ScheduleDraft.is_deleted),
        )
        .order_by(desc(ScheduleDraft.created_at))
    )
    return session.exec(statement).first()


def soft_delete_schedule_draft(
    session: Session,
    draft_id: uuid.UUID,
) -> bool:
    """
    Soft-deletes an active ScheduleDraft by setting is_deleted=True.
    """
    draft = session.get(ScheduleDraft, draft_id)
    if draft and not draft.is_deleted:
        draft.is_deleted = True
        session.add(draft)
        session.commit()
        return True
    return False
