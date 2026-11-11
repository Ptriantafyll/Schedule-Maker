"""Routes for schedule generation and retrieval."""

import json
import uuid
from fastapi import (
    APIRouter,
    Depends,
    File,
    Form,
    HTTPException,
    Query,
    UploadFile,
    status,
    Response
)
from sqlmodel import Session
from src.db.connection import get_session
from src.schedule.schemas import ScheduleDraftRead, ScheduleSummaryRead, TargetMonthResponse, GenerateFromRosterRequest
from src.schedule import controllers as schedule_controllers
from src.schedule.schemas import ExcelRosterDiffResponse
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
    "",
    response_model=list[ScheduleSummaryRead],
    status_code=status.HTTP_200_OK,
    include_in_schema=False,
)
@router.get(
    "/",
    response_model=list[ScheduleSummaryRead],
    status_code=status.HTTP_200_OK,
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
    "/generate-from-roster",
    response_model=ScheduleDraftRead,
    status_code=status.HTTP_201_CREATED,
)
def generate_schedule_from_roster(
    payload: GenerateFromRosterRequest,
    session: Session = Depends(get_session),
    current_user: UserModel = Depends(require_department_admin),
):
    """Generate a monthly duty schedule draft from the department roster."""
    return schedule_controllers.generate_schedule_from_roster_controller(
        session=session,
        target_month=payload.target_month,
        department_id=payload.department_id,
        current_user=current_user,
        country_code=payload.country_code,
    )


@router.post(
    "/generate-from-excel",
    response_model=ScheduleDraftRead,
    status_code=status.HTTP_201_CREATED,
)
def generate_schedule_from_excel(
    file: UploadFile = File(...),
    target_month: str = Form(...),
    department_id: uuid.UUID | None = Form(None),
    country_code: str | None = Form(None),
    sync_roster: bool = Form(False),
    remove_doctor_ids: str | None = Form(None),
    ignored_new_doctors: str | None = Form(None),
    session: Session = Depends(get_session),
    current_user: UserModel = Depends(require_department_admin),
):
    """Generate a monthly duty schedule draft from an uploaded Excel roster."""
    parsed_remove_ids = None
    if remove_doctor_ids:
        try:
            raw_ids = json.loads(remove_doctor_ids)
            if not isinstance(raw_ids, list):
                raise ValueError("Expected list of doctor IDs.")
            parsed_remove_ids = [uuid.UUID(str(i)) for i in raw_ids]
        except (json.JSONDecodeError, ValueError) as exc:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail=f"Invalid remove_doctor_ids format: {exc}",
            )

    parsed_ignored = None
    if ignored_new_doctors:
        try:
            raw_ignored = json.loads(ignored_new_doctors)
            if not isinstance(raw_ignored, list):
                raise ValueError("Expected list of doctor names.")
            parsed_ignored = [str(name) for name in raw_ignored]
        except (json.JSONDecodeError, ValueError) as exc:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail=f"Invalid ignored_new_doctors format: {exc}",
            )

    file_bytes = file.file.read()
    return schedule_controllers.generate_schedule_from_excel_controller(
        session=session,
        file_bytes=file_bytes,
        target_month=target_month,
        department_id=department_id,
        current_user=current_user,
        filename=file.filename,
        country_code=country_code,
        sync_roster=sync_roster,
        remove_doctor_ids=parsed_remove_ids,
        ignored_new_doctors=parsed_ignored,
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
    country_code: str | None = Query(None),
    session: Session = Depends(get_session),
    current_user: UserModel = Depends(require_department_member),
):
    """Export a draft schedule to excel"""
    file_bytes, filename = schedule_controllers.export_schedule_draft_controller(
        session=session,
        draft_id=draft_id,
        current_user=current_user,
        country_code=country_code,
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
    "/preview-excel-roster-diff",
    response_model=ExcelRosterDiffResponse,
    status_code=status.HTTP_200_OK
)
def preview_excel_roster_diff(
    file: UploadFile = File(...),
    target_month: str = Form(...),
    department_id: uuid.UUID | None = Form(None),
    session: Session = Depends(get_session),
    current_user: UserModel = Depends(require_department_admin)
):
    """Preview roster differences between an uploaded Excel spreadsheet and the department database."""
    file_bytes = file.file.read()
    return schedule_controllers.preview_excel_roster_diff_controller(
        session=session,
        file_bytes=file_bytes,
        target_month=target_month,
        department_id=department_id,
        current_user=current_user,
        filename=file.filename
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
