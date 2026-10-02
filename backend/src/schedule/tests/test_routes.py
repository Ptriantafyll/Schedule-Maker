"""Integration tests for the schedule FastAPI routes."""

import io
import uuid
import openpyxl
import pytest
from openpyxl import Workbook
from src.user.models import UserRole
from src.schedule.models import ScheduleDraft
from src.schedule import repository as schedule_repository


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


@pytest.fixture(name="valid_excel_bytes")
def valid_excel_bytes_fixture() -> bytes:
    """Creates valid in-memory Excel file bytes with 6 doctors."""
    doctor_rows = [
        {"Name": "Dr. Gregory House", "Email": "house@hospital.org", "Position": "ER", "Team": "Alpha"},
        {"Name": "Dr. Allison Cameron", "Email": "cameron@hospital.org", "Position": "ER", "Team": "Alpha"},
        {"Name": "Dr. Eric Foreman", "Email": "foreman@hospital.org", "Position": "ER", "Team": "Alpha"},
        {"Name": "Dr. Robert Chase", "Email": "chase@hospital.org", "Position": "ER", "Team": "Beta"},
        {"Name": "Dr. James Wilson", "Email": "wilson@hospital.org", "Position": "ER", "Team": "Beta"},
        {"Name": "Dr. Lisa Cuddy", "Email": "cuddy@hospital.org", "Position": "ER", "Team": "Beta"},
    ]
    return _create_test_workbook_bytes(doctor_rows)


def test_generate_schedule_endpoint_success(
    client,
    department_factory,
    user_factory,
    auth_headers_factory,
    valid_excel_bytes,
):
    """Verify Department Admin can upload Excel and generate schedule draft."""
    dept = department_factory()
    admin = user_factory(role=UserRole.DEPARTMENT_ADMIN, department_id=dept.id)
    headers = auth_headers_factory(admin)

    response = client.post(
        "/api/v1/schedules/generate-from-excel",
        data={
            "target_month": "2026-11",
            "department_id": str(dept.id),
        },
        files={
            "file": ("november_roster.xlsx", valid_excel_bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        },
        headers=headers,
    )

    assert response.status_code == 201
    data = response.json()
    assert data["target_month"] == "2026-11"
    assert data["department_id"] == str(dept.id)
    assert data["source_filename"] == "november_roster.xlsx"
    assert data["solver_status"] in ("OPTIMAL", "FEASIBLE")
    assert len(data["assignments"]) == 30


def test_generate_schedule_rejects_non_excel_extension(
    client,
    department_factory,
    user_factory,
    auth_headers_factory,
):
    """Verify endpoint rejects files that do not end with .xlsx."""
    dept = department_factory()
    admin = user_factory(role=UserRole.DEPARTMENT_ADMIN, department_id=dept.id)
    headers = auth_headers_factory(admin)

    response = client.post(
        "/api/v1/schedules/generate-from-excel",
        data={"target_month": "2026-11", "department_id": str(dept.id)},
        files={"file": ("roster.csv", b"dummy,csv,content", "text/csv")},
        headers=headers,
    )

    assert response.status_code == 400
    assert "Only .xlsx files are supported" in response.json()["detail"]


def test_generate_schedule_rejects_non_admin_role(
    client,
    department_factory,
    user_factory,
    auth_headers_factory,
    valid_excel_bytes,
):
    """Verify non-admin role (e.g. VIEWER) is forbidden from generating schedules."""
    dept = department_factory()
    viewer = user_factory(role=UserRole.VIEWER, department_id=dept.id)
    headers = auth_headers_factory(viewer)

    response = client.post(
        "/api/v1/schedules/generate-from-excel",
        data={"target_month": "2026-11", "department_id": str(dept.id)},
        files={"file": ("roster.xlsx", valid_excel_bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
        headers=headers,
    )

    assert response.status_code == 403


def test_generate_schedule_rejects_cross_department_access(
    client,
    department_factory,
    user_factory,
    auth_headers_factory,
    valid_excel_bytes,
):
    """Verify Department Admin cannot generate schedules for another department."""
    dept_a = department_factory()
    dept_b = department_factory()
    admin_a = user_factory(role=UserRole.DEPARTMENT_ADMIN, department_id=dept_a.id)
    headers = auth_headers_factory(admin_a)

    response = client.post(
        "/api/v1/schedules/generate-from-excel",
        data={"target_month": "2026-11", "department_id": str(dept_b.id)},
        files={"file": ("roster.xlsx", valid_excel_bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
        headers=headers,
    )

    assert response.status_code == 403


def test_generate_schedule_unauthenticated_returns_401(
    client,
    department_factory,
    valid_excel_bytes,
):
    """Verify unauthenticated request is rejected with 401."""
    dept = department_factory()

    response = client.post(
        "/api/v1/schedules/generate-from-excel",
        data={"target_month": "2026-11", "department_id": str(dept.id)},
        files={"file": ("roster.xlsx", valid_excel_bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )

    assert response.status_code == 401


def test_get_schedule_draft_endpoint_success(
    client,
    session,
    department_factory,
    user_factory,
    auth_headers_factory,
):
    """Verify department member can fetch the active schedule draft."""
    dept = department_factory()
    viewer = user_factory(role=UserRole.VIEWER, department_id=dept.id)
    headers = auth_headers_factory(viewer)

    draft = ScheduleDraft(
        department_id=dept.id,
        target_month="2026-11",
        source_filename="test.xlsx",
        total_duties=5,
        solver_status="OPTIMAL",
        assignments=[],
    )
    saved = schedule_repository.create_schedule_draft(session, draft)

    response = client.get(
        f"/api/v1/schedules/draft?target_month=2026-11&department_id={dept.id}",
        headers=headers,
    )

    assert response.status_code == 200
    data = response.json()
    assert data["id"] == str(saved.id)
    assert data["target_month"] == "2026-11"
    assert data["department_id"] == str(dept.id)


def test_get_schedule_draft_endpoint_not_found_returns_404(
    client,
    department_factory,
    user_factory,
    auth_headers_factory,
):
    """Verify 404 is returned when no draft exists for target month."""
    dept = department_factory()
    viewer = user_factory(role=UserRole.VIEWER, department_id=dept.id)
    headers = auth_headers_factory(viewer)

    response = client.get(
        f"/api/v1/schedules/draft?target_month=2026-12&department_id={dept.id}",
        headers=headers,
    )

    assert response.status_code == 404


def test_get_schedule_draft_rejects_cross_department_access(
    client,
    department_factory,
    user_factory,
    auth_headers_factory,
):
    """Verify department member cannot read drafts of another department."""
    dept_a = department_factory()
    dept_b = department_factory()
    viewer_a = user_factory(role=UserRole.VIEWER, department_id=dept_a.id)
    headers = auth_headers_factory(viewer_a)

    response = client.get(
        f"/api/v1/schedules/draft?target_month=2026-11&department_id={dept_b.id}",
        headers=headers,
    )

    assert response.status_code == 403


def test_export_schedule_endpoint_success(
    client,
    session,
    department_factory,
    user_factory,
    auth_headers_factory,
):
    """Verify department member can export active draft as Excel file."""
    dept = department_factory()
    viewer = user_factory(role=UserRole.VIEWER, department_id=dept.id)
    headers = auth_headers_factory(viewer)

    draft = ScheduleDraft(
        department_id=dept.id,
        target_month="2026-11",
        source_filename="test.xlsx",
        total_duties=5,
        solver_status="OPTIMAL",
        assignments=[],
    )
    saved = schedule_repository.create_schedule_draft(session, draft)

    response = client.get(
        f"/api/v1/schedules/export-excel?draft_id={saved.id}",
        headers=headers,
    )

    assert response.status_code == 200
    assert response.headers["content-type"] == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    assert 'attachment; filename="schedule_2026-11.xlsx"' in response.headers["content-disposition"]

    wb = openpyxl.load_workbook(io.BytesIO(response.content))
    assert wb.active is not None


def test_export_schedule_endpoint_unauthenticated_returns_401(client):
    """Verify unauthenticated export request returns 401."""
    response = client.get(f"/api/v1/schedules/export-excel?draft_id={uuid.uuid4()}")
    assert response.status_code == 401


def test_export_schedule_endpoint_not_found_returns_404(
    client,
    department_factory,
    user_factory,
    auth_headers_factory,
):
    """Verify export returns 404 if draft does not exist."""
    dept = department_factory()
    viewer = user_factory(role=UserRole.VIEWER, department_id=dept.id)
    headers = auth_headers_factory(viewer)

    response = client.get(
        f"/api/v1/schedules/export-excel?draft_id={uuid.uuid4()}",
        headers=headers,
    )
    assert response.status_code == 404


def test_export_schedule_endpoint_rejects_cross_department_access(
    client,
    session,
    department_factory,
    user_factory,
    auth_headers_factory,
):
    """Verify user cannot export drafts belonging to another department."""
    dept_a = department_factory()
    dept_b = department_factory()
    viewer_a = user_factory(role=UserRole.VIEWER, department_id=dept_a.id)
    headers = auth_headers_factory(viewer_a)

    draft_b = ScheduleDraft(
        department_id=dept_b.id,
        target_month="2026-11",
        source_filename="test.xlsx",
        total_duties=0,
        solver_status="OPTIMAL",
        assignments=[],
    )
    saved_b = schedule_repository.create_schedule_draft(session, draft_b)

    response = client.get(
        f"/api/v1/schedules/export-excel?draft_id={saved_b.id}",
        headers=headers,
    )
    assert response.status_code == 403

