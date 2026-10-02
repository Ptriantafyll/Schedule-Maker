"""Unit tests for the schedule controller layer."""

import io
import uuid
import pytest
from openpyxl import Workbook
from fastapi import HTTPException
from src.user.models import User as UserModel, UserRole
from src.schedule.models import ScheduleDraft
from src.schedule import repository as schedule_repository
from src.schedule.controllers import (
    get_schedule_draft_controller,
    get_schedule_draft_for_department_month_controller,
    generate_schedule_from_excel_controller,
    export_schedule_draft_controller,
)


def _create_test_workbook_bytes(doctor_rows: list[dict]) -> bytes:
    """Helper to generate an in-memory .xlsx file with a Doctors sheet."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Doctors"
    if doctor_rows:
        headers = list(doctor_rows[0].keys())
        ws.append(headers)
        for row in doctor_rows:
            ws.append([row.get(h) for h in headers])
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


def test_get_schedule_draft_controller_success(session, department_factory):
    """Verify controller fetches an existing draft via repository."""
    dept = department_factory()
    draft = ScheduleDraft(
        department_id=dept.id,
        target_month="2026-11",
        source_filename="test.xlsx",
        total_duties=5,
        solver_status="OPTIMAL",
        assignments=[],
    )
    saved = schedule_repository.create_schedule_draft(session, draft)

    result = get_schedule_draft_controller(draft_id=saved.id, session=session)
    assert result.id == saved.id
    assert result.department_id == dept.id


def test_get_schedule_draft_controller_not_found_raises_404(session):
    """Verify controller raises HTTPException(404) when draft is not found."""
    fake_id = uuid.uuid4()
    with pytest.raises(HTTPException) as exc_info:
        get_schedule_draft_controller(draft_id=fake_id, session=session)

    assert exc_info.value.status_code == 404
    assert "not found" in exc_info.value.detail.lower()


def test_get_schedule_draft_for_department_month_controller_success(session, department_factory):
    """Verify controller fetches active draft for department and month."""
    dept = department_factory()
    draft = ScheduleDraft(
        department_id=dept.id,
        target_month="2026-11",
        source_filename="test.xlsx",
        total_duties=5,
        solver_status="OPTIMAL",
        assignments=[],
    )
    saved = schedule_repository.create_schedule_draft(session, draft)

    result = get_schedule_draft_for_department_month_controller(
        department_id=dept.id,
        target_month="2026-11",
        session=session,
    )
    assert result.id == saved.id


def test_get_schedule_draft_for_department_month_not_found_raises_404(session):
    """Verify controller raises 404 when no draft exists for department and month."""
    fake_id = uuid.uuid4()
    with pytest.raises(HTTPException) as exc_info:
        get_schedule_draft_for_department_month_controller(
            department_id=fake_id,
            target_month="2026-11",
            session=session,
        )

    assert exc_info.value.status_code == 404


def test_generate_schedule_controller_translates_value_error_to_400(session):
    """Verify controller translates service ValueError into HTTPException(400)."""
    fake_id = uuid.uuid4()
    file_bytes = _create_test_workbook_bytes([{"Name": "Dr. House", "Email": "house@hospital.org", "Position": "ER"}])

    with pytest.raises(HTTPException) as exc_info:
        generate_schedule_from_excel_controller(
            session=session,
            file_bytes=file_bytes,
            department_id=fake_id,
            target_month="2026-11",
            source_filename="roster.xlsx",
        )

    assert exc_info.value.status_code == 400
    assert "Department not found" in exc_info.value.detail


def test_generate_schedule_controller_success(session, department_factory):
    """Verify controller successfully generates and returns a draft."""
    dept = department_factory()
    doctor_rows = [
        {"Name": "Dr. Gregory House", "Email": "house@hospital.org", "Position": "ER", "Team": "Alpha"},
        {"Name": "Dr. Allison Cameron", "Email": "cameron@hospital.org", "Position": "ER", "Team": "Alpha"},
        {"Name": "Dr. Eric Foreman", "Email": "foreman@hospital.org", "Position": "ER", "Team": "Alpha"},
        {"Name": "Dr. Robert Chase", "Email": "chase@hospital.org", "Position": "ER", "Team": "Beta"},
        {"Name": "Dr. James Wilson", "Email": "wilson@hospital.org", "Position": "ER", "Team": "Beta"},
        {"Name": "Dr. Lisa Cuddy", "Email": "cuddy@hospital.org", "Position": "ER", "Team": "Beta"},
    ]
    file_bytes = _create_test_workbook_bytes(doctor_rows)

    result = generate_schedule_from_excel_controller(
        session=session,
        file_bytes=file_bytes,
        department_id=dept.id,
        target_month="2026-11",
        source_filename="november_roster.xlsx",
    )

    assert result.id is not None
    assert result.department_id == dept.id
    assert result.target_month == "2026-11"
    assert result.source_filename == "november_roster.xlsx"
    assert result.solver_status in ("OPTIMAL", "FEASIBLE")
    assert len(result.assignments) == 30


def test_generate_schedule_controller_rejects_non_excel_filename(session, department_factory):
    """Verify controller rejects filenames not ending with .xlsx."""
    dept = department_factory()
    with pytest.raises(HTTPException) as exc_info:
        generate_schedule_from_excel_controller(
            session=session,
            file_bytes=b"fake-bytes",
            target_month="2026-11",
            department_id=dept.id,
            filename="data.csv",
        )

    assert exc_info.value.status_code == 400
    assert "Only .xlsx files are supported" in exc_info.value.detail


def test_generate_schedule_controller_rejects_cross_department_access(
    session,
    department_factory,
    user_factory,
):
    """Verify controller rejects generating schedule for a foreign department."""
    dept_a = department_factory()
    dept_b = department_factory()
    admin_a = user_factory(role=UserRole.DEPARTMENT_ADMIN, department_id=dept_a.id)

    with pytest.raises(HTTPException) as exc_info:
        generate_schedule_from_excel_controller(
            session=session,
            file_bytes=b"fake-bytes",
            target_month="2026-11",
            department_id=dept_b.id,
            current_user=admin_a,
            filename="roster.xlsx",
        )

    assert exc_info.value.status_code == 403
    assert "Cannot access schedules for another department" in exc_info.value.detail


def test_get_schedule_draft_controller_rejects_cross_department_access(
    session,
    department_factory,
    user_factory,
):
    """Verify controller rejects fetching schedule draft for a foreign department."""
    dept_a = department_factory()
    dept_b = department_factory()
    viewer_a = user_factory(role=UserRole.VIEWER, department_id=dept_a.id)

    with pytest.raises(HTTPException) as exc_info:
        get_schedule_draft_for_department_month_controller(
            session=session,
            target_month="2026-11",
            department_id=dept_b.id,
            current_user=viewer_a,
        )

    assert exc_info.value.status_code == 403
    assert "Cannot access schedules for another department" in exc_info.value.detail


def test_controller_rejects_user_with_missing_department_scope(session):
    """Verify controller rejects non-super admin users without a department_id."""
    admin_no_dept = UserModel(
        id=uuid.uuid4(),
        email="nodept@test.com",
        full_name="No Dept User",
        role=UserRole.DEPARTMENT_ADMIN,
        department_id=None,
    )

    with pytest.raises(HTTPException) as exc_info:
        get_schedule_draft_for_department_month_controller(
            session=session,
            target_month="2026-11",
            current_user=admin_no_dept,
        )

    assert exc_info.value.status_code == 403
    assert "Invalid account scope" in exc_info.value.detail


def test_controller_super_admin_requires_department_id(
    session,
    user_factory,
):
    """Verify Super Admin must provide an explicit department_id."""
    super_admin = user_factory(role=UserRole.SUPER_ADMIN, department_id=None)

    with pytest.raises(HTTPException) as exc_info:
        get_schedule_draft_for_department_month_controller(
            session=session,
            target_month="2026-11",
            current_user=super_admin,
            department_id=None,
        )

    assert exc_info.value.status_code == 400
    assert "Department ID is required for Super Admin" in exc_info.value.detail


def test_export_schedule_draft_controller_success(session, department_factory):
    """Verify controller loads draft and exports Excel bytes with filename."""
    dept = department_factory()
    draft = ScheduleDraft(
        department_id=dept.id,
        target_month="2026-11",
        source_filename="november_roster.xlsx",
        total_duties=0,
        solver_status="OPTIMAL",
        assignments=[],
    )
    saved = schedule_repository.create_schedule_draft(session, draft)

    file_bytes, filename = export_schedule_draft_controller(
        session=session,
        draft_id=saved.id,
        department_id=dept.id,
    )

    assert isinstance(file_bytes, bytes)
    assert len(file_bytes) > 0
    assert filename == "schedule_2026-11.xlsx"


def test_export_schedule_draft_controller_not_found_raises_404(session):
    """Verify controller raises 404 when draft_id does not exist."""
    fake_id = uuid.uuid4()
    with pytest.raises(HTTPException) as exc_info:
        export_schedule_draft_controller(
            session=session,
            draft_id=fake_id,
        )

    assert exc_info.value.status_code == 404
    assert "not found" in exc_info.value.detail.lower()


def test_export_schedule_draft_controller_rejects_cross_department_access(
    session,
    department_factory,
    user_factory,
):
    """Verify controller rejects export when user belongs to a different department."""
    dept_a = department_factory()
    dept_b = department_factory()
    viewer_a = user_factory(role=UserRole.VIEWER, department_id=dept_a.id)

    draft_b = ScheduleDraft(
        department_id=dept_b.id,
        target_month="2026-11",
        source_filename="test.xlsx",
        total_duties=0,
        solver_status="OPTIMAL",
        assignments=[],
    )
    saved_b = schedule_repository.create_schedule_draft(session, draft_b)

    with pytest.raises(HTTPException) as exc_info:
        export_schedule_draft_controller(
            session=session,
            draft_id=saved_b.id,
            current_user=viewer_a,
        )

    assert exc_info.value.status_code == 403
    assert "Cannot access schedules for another department" in exc_info.value.detail


