"""Routes for schedule generation and retrieval."""

import uuid
from fastapi import (
    APIRouter,
    Depends,
    File,
    Form,
    Query,
    UploadFile,
    status,
    Response
)
from sqlmodel import Session
from src.db.connection import get_session
from src.schedule.schemas import ScheduleDraftRead, ScheduleSummaryRead, TargetMonthResponse
from src.schedule import controllers as schedule_controllers
from src.user.models import User as UserModel
from src.auth.dependencies import (
    require_department_admin,
    require_department_member,
)

router = APIRouter(
    prefix="/schedules",
    tags=["Schedules"],
)


@router.get(
    "/",
    response_model=list[ScheduleSummaryRead],
    status_code=status.HTTP_200_OK
)
def list_schedules(
    department_id: uuid.UUID | None = Query(None),
    session: Session = Depends(get_session),
    current_user: UserModel = Depends(require_department_member),
):
    """Retrieve all schedule summaries for a department."""
    return schedule_controllers.list_schedules_controller(
        session=session,
        department_id=department_id,
        current_user=current_user,
    )


@router.post(
    "/generate-from-excel",
    response_model=ScheduleDraftRead,
    status_code=status.HTTP_201_CREATED,
)
async def generate_schedule_from_excel(
    file: UploadFile = File(...),
    target_month: str = Form(...),
    department_id: uuid.UUID | None = Form(None),
    session: Session = Depends(get_session),
    current_user: UserModel = Depends(require_department_admin),
):
    """Generate a monthly duty schedule draft from an uploaded Excel roster."""
    file_bytes = await file.read()
    return schedule_controllers.generate_schedule_from_excel_controller(
        session=session,
        file_bytes=file_bytes,
        target_month=target_month,
        department_id=department_id,
        current_user=current_user,
        filename=file.filename,
    )


@router.get(
    "/draft",
    response_model=ScheduleDraftRead,
    status_code=status.HTTP_200_OK,
)
def get_schedule_draft(
    target_month: str = Query(...),
    department_id: uuid.UUID | None = Query(None),
    session: Session = Depends(get_session),
    current_user: UserModel = Depends(require_department_member),
):
    """Fetch the active schedule draft for a department and target month."""
    return schedule_controllers.get_schedule_draft_for_department_month_controller(
        session=session,
        target_month=target_month,
        department_id=department_id,
        current_user=current_user,
    )


@router.get("/export-excel")
def export_schedule_draft(
    draft_id: uuid.UUID = Query(...),
    session: Session = Depends(get_session),
    current_user: UserModel = Depends(require_department_member),
):
    """Export a draft schedule to excel"""
    file_bytes, filename = schedule_controllers.export_schedule_draft_controller(
        session=session,
        draft_id=draft_id,
        current_user=current_user,
    )

    return Response(
        content=file_bytes,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@router.get(
    "/target-month",
    response_model=TargetMonthResponse,
    status_code=status.HTTP_200_OK,
)
def get_target_month(
    department_id: uuid.UUID | None = Query(None),
    session: Session = Depends(get_session),
    current_user: UserModel = Depends(require_department_member),
):
    """Get the target month for the next schedule."""
    return schedule_controllers.get_target_month_controller(
        session=session,
        department_id=department_id,
        current_user=current_user,
    )


@router.post(
    "/{draft_id}/publish",
    response_model=ScheduleDraftRead,
    status_code=status.HTTP_200_OK,
)
def publish_schedule_draft(
    draft_id: uuid.UUID,
    session: Session = Depends(get_session),
    current_user: UserModel = Depends(require_department_admin),
):
    "Publishes a schedule draft."
    return schedule_controllers.publish_schedule_draft_controller(
        session=session,
        draft_id=draft_id,
        current_user=current_user,
    )
