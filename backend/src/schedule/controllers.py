"""Controllers for schedule management and generation.

Orchestrates calls between FastAPI routes, schedule repository, and solver service,
translating domain exceptions into HTTP status codes.
"""

import uuid
from fastapi import HTTPException, status
from sqlmodel import Session

from src.user.models import User as UserModel, UserRole
from src.schedule.models import ScheduleDraft
from src.schedule import repository as schedule_repository
from src.schedule import service as schedule_service


def _resolve_and_verify_department(
    current_user: UserModel | None,
    requested_department_id: uuid.UUID | None,
) -> uuid.UUID:
    """Validate tenant boundaries and return authorized department ID."""
    if current_user is None:
        if requested_department_id is None:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail="Department ID is required.",
            )
        return requested_department_id

    if current_user.role != UserRole.SUPER_ADMIN:
        if current_user.department_id is None:
            raise HTTPException(
                status_code=status.HTTP_403_FORBIDDEN,
                detail="Invalid account scope.",
            )
        if (
            requested_department_id is not None
            and requested_department_id != current_user.department_id
        ):
            raise HTTPException(
                status_code=status.HTTP_403_FORBIDDEN,
                detail="Cannot access schedules for another department.",
            )
        return current_user.department_id

    if requested_department_id is None:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Department ID is required for Super Admin.",
        )
    return requested_department_id


def get_schedule_draft_controller(
    draft_id: uuid.UUID,
    session: Session,
) -> ScheduleDraft:
    """Fetch a schedule draft by ID or raise HTTP 404."""
    draft = schedule_repository.get_schedule_draft_by_id(session, draft_id)
    if draft is None:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Schedule draft not found",
        )
    return draft


def get_schedule_draft_for_department_month_controller(
    session: Session,
    target_month: str,
    department_id: uuid.UUID | None = None,
    current_user: UserModel | None = None,
) -> ScheduleDraft:
    """Fetch active schedule draft for department and month or raise HTTP 404."""
    dept_id = _resolve_and_verify_department(current_user, department_id)

    draft = schedule_repository.get_schedule_draft_for_department_month(
        session=session,
        department_id=dept_id,
        target_month=target_month,
    )
    if draft is None:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Schedule draft not found",
        )
    return draft


def generate_schedule_from_excel_controller(
    session: Session,
    file_bytes: bytes,
    target_month: str,
    department_id: uuid.UUID | None = None,
    current_user: UserModel | None = None,
    filename: str | None = None,
    source_filename: str | None = None,
) -> ScheduleDraft:
    """Generate a schedule draft from Excel file bytes.

    Validates file extension, tenant scope, and translates domain ValueError into HTTP 400.
    """
    actual_filename = filename or source_filename or "schedule.xlsx"
    if not actual_filename.endswith(".xlsx"):
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Only .xlsx files are supported",
        )

    dept_id = _resolve_and_verify_department(current_user, department_id)

    try:
        return schedule_service.generate_schedule_from_excel(
            session=session,
            file_bytes=file_bytes,
            department_id=dept_id,
            target_month=target_month,
            source_filename=actual_filename,
        )
    except ValueError as exc:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=str(exc),
        ) from exc


def export_schedule_draft_controller(
    session: Session,
    draft_id: uuid.UUID,
    department_id: uuid.UUID | None = None,
    current_user: UserModel | None = None,
) -> tuple[bytes, str]:
    """Exports a draft schedule to an excel file fia bytes"""
    draft = schedule_repository.get_schedule_draft_by_id(
        session=session,
        draft_id=draft_id
    )

    if not draft or draft.is_deleted:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Schedule draft not found."
        )

    if department_id is not None and department_id != draft.department_id:
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="Cannot access schedules for another department.",
        )

    _resolve_and_verify_department(
        current_user=current_user,
        requested_department_id=draft.department_id,
    )

    file_bytes = schedule_service.export_schedule_draft_to_excel(draft)
    filename = f"schedule_{draft.target_month}.xlsx"

    return (file_bytes, filename)


def list_schedules_controller(
    session: Session,
    department_id: uuid.UUID | None = None,
    current_user: UserModel | None = None,
) -> list[ScheduleDraft]:
    """Retrieve all schedule summaries for an authorized department."""
    target_dept_id = _resolve_and_verify_department(
        current_user=current_user,
        requested_department_id=department_id,
    )

    return schedule_repository.list_schedules_for_department(
        session=session,
        department_id=target_dept_id,
    )


def get_target_month_controller(
    session: Session,
    department_id: uuid.UUID | None = None,
    current_user: UserModel | None = None,
) -> dict[str, str | None]:
    """Retrieve the target month for the next schedule"""
    target_dept_id = _resolve_and_verify_department(
        current_user=current_user,
        requested_department_id=department_id,
    )
    return schedule_service.resolve_next_target_month(
        session=session,
        department_id=target_dept_id,
    )


def publish_schedule_draft_controller(
    session: Session,
    draft_id: uuid.UUID,
    current_user: UserModel | None = None
) -> ScheduleDraft:
    """Publishes a schedule draft"""
    draft = schedule_repository.get_schedule_draft_by_id(session, draft_id)
    if not draft or draft.is_deleted:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Schedule draft not found.",
        )

    department_id = _resolve_and_verify_department(
        current_user=current_user,
        requested_department_id=draft.department_id
    )

    try:
        return schedule_service.publish_schedule_draft(
            session=session,
            draft_id=draft_id,
            department_id=department_id
        )

    except ValueError as exc:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=str(exc)
        )
