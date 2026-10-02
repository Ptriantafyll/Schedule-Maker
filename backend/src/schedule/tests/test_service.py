"""Unit tests for the schedule service layer."""

import io
import uuid
import openpyxl
import pytest
from openpyxl import Workbook
from src.department.models import Department as DepartmentModel
from src.models import ScheduleConfig
from src.schedule.models import ScheduleDraft
from src.schedule.service import (
    generate_schedule_from_excel,
    export_schedule_draft_to_excel,
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


def test_generate_schedule_from_excel_success(session):
    """Verify that a valid Excel workbook generates and persists a schedule draft."""
    # 1. Create Department in DB
    dept = DepartmentModel(name="Emergency Medicine", code="EMERG")
    session.add(dept)
    session.commit()
    session.refresh(dept)

    # 2. 6 doctors for 30 days in November ensures feasibility with standard limits
    doctor_rows = [
        {"Name": "Dr. Gregory House", "Email": "house@hospital.org", "Position": "ER", "Team": "Alpha", "Unavailability": "14"},
        {"Name": "Dr. Allison Cameron", "Email": "cameron@hospital.org", "Position": "ER", "Team": "Alpha"},
        {"Name": "Dr. Eric Foreman", "Email": "foreman@hospital.org", "Position": "ER", "Team": "Alpha"},
        {"Name": "Dr. Robert Chase", "Email": "chase@hospital.org", "Position": "ER", "Team": "Beta"},
        {"Name": "Dr. James Wilson", "Email": "wilson@hospital.org", "Position": "ER", "Team": "Beta"},
        {"Name": "Dr. Lisa Cuddy", "Email": "cuddy@hospital.org", "Position": "ER", "Team": "Beta"},
    ]
    file_bytes = _create_test_workbook_bytes(doctor_rows)

    # 3. Call Service
    draft = generate_schedule_from_excel(
        session=session,
        file_bytes=file_bytes,
        department_id=dept.id,
        target_month="2026-11",
        source_filename="november_roster.xlsx",
    )

    # 4. Assertions
    assert draft.id is not None
    assert draft.department_id == dept.id
    assert draft.target_month == "2026-11"
    assert draft.source_filename == "november_roster.xlsx"
    assert draft.solver_status in ("OPTIMAL", "FEASIBLE")
    assert draft.status == "draft"
    assert draft.total_duties == 30  # 30 days in November
    assert len(draft.assignments) == 30
    assert draft.unavailabilities == {"Dr. Gregory House": [14]}

    # Verify assignment item structure
    first_assignment = draft.assignments[0]
    assert "date" in first_assignment
    assert "day_name" in first_assignment
    assert "doctor_name" in first_assignment
    assert "doctor_email" in first_assignment
    assert "position" in first_assignment
    assert "shift" in first_assignment


def test_generate_schedule_nonexistent_department_raises_error(session):
    """Verify that attempting to generate for a missing department raises ValueError."""
    fake_dept_id = uuid.uuid4()
    doctor_rows = [
        {"Name": "Dr. Gregory House", "Email": "house@hospital.org", "Position": "ER"}
    ]
    file_bytes = _create_test_workbook_bytes(doctor_rows)

    with pytest.raises(ValueError, match="Department not found"):
        generate_schedule_from_excel(
            session=session,
            file_bytes=file_bytes,
            department_id=fake_dept_id,
            target_month="2026-11",
            source_filename="roster.xlsx",
        )


def test_generate_schedule_soft_deleted_department_raises_error(session):
    """Verify that attempting to generate for a soft-deleted department raises ValueError."""
    dept = DepartmentModel(name="Archived Clinic", code="ARCH", is_deleted=True)
    session.add(dept)
    session.commit()
    session.refresh(dept)

    doctor_rows = [
        {"Name": "Dr. Gregory House", "Email": "house@hospital.org", "Position": "ER"}
    ]
    file_bytes = _create_test_workbook_bytes(doctor_rows)

    with pytest.raises(ValueError, match="Department not found"):
        generate_schedule_from_excel(
            session=session,
            file_bytes=file_bytes,
            department_id=dept.id,
            target_month="2026-11",
            source_filename="roster.xlsx",
        )


def test_generate_schedule_infeasible_constraints_raises_error(session):
    """Verify that impossible constraints (e.g. 1 doctor for 30 shifts) raises ValueError."""
    dept = DepartmentModel(name="Understaffed Dept", code="UNDER")
    session.add(dept)
    session.commit()
    session.refresh(dept)

    # 1 single doctor cannot cover 30 days due to max_duties and no-consecutive constraints
    doctor_rows = [
        {"Name": "Dr. Lone Ranger", "Email": "lone@hospital.org", "Position": "ER", "Team": "Solo"}
    ]
    file_bytes = _create_test_workbook_bytes(doctor_rows)

    with pytest.raises(ValueError, match="infeasible"):
        generate_schedule_from_excel(
            session=session,
            file_bytes=file_bytes,
            department_id=dept.id,
            target_month="2026-11",
            source_filename="impossible.xlsx",
        )


def test_generate_schedule_corrupted_excel_file_raises_error(session):
    """Verify that passing corrupted/non-Excel binary bytes raises ValueError."""
    dept = DepartmentModel(name="Emergency", code="EMERG")
    session.add(dept)
    session.commit()
    session.refresh(dept)

    corrupted_bytes = b"This is plain text and definitely not a valid xlsx zip archive"

    with pytest.raises(ValueError, match="Invalid or corrupted Excel file"):
        generate_schedule_from_excel(
            session=session,
            file_bytes=corrupted_bytes,
            department_id=dept.id,
            target_month="2026-11",
            source_filename="corrupted.xlsx",
        )


def test_generate_schedule_with_custom_config_override(session):
    """Verify that custom ScheduleConfig overrides are applied and generate successfully."""
    dept = DepartmentModel(name="Flexible Dept", code="FLEX")
    session.add(dept)
    session.commit()
    session.refresh(dept)

    doctor_rows = [
        {"Name": "Dr. A", "Email": "a@hospital.org", "Position": "ER"},
        {"Name": "Dr. B", "Email": "b@hospital.org", "Position": "ER"},
        {"Name": "Dr. C", "Email": "c@hospital.org", "Position": "ER"},
        {"Name": "Dr. D", "Email": "d@hospital.org", "Position": "ER"},
        {"Name": "Dr. E", "Email": "e@hospital.org", "Position": "ER"},
        {"Name": "Dr. F", "Email": "f@hospital.org", "Position": "ER"},
    ]
    file_bytes = _create_test_workbook_bytes(doctor_rows)

    custom_config = ScheduleConfig(
        solver_time_limit=15,
        max_duties_per_month=7,
        w_every_other_penalty=10,
    )

    draft = generate_schedule_from_excel(
        session=session,
        file_bytes=file_bytes,
        department_id=dept.id,
        target_month="2026-11",
        source_filename="custom_config.xlsx",
        config_override=custom_config,
    )

    assert draft.id is not None
    assert draft.solver_status in ("OPTIMAL", "FEASIBLE")
    assert len(draft.assignments) == 30


def test_export_schedule_draft_to_excel_success():
    """Verify that export_schedule_draft_to_excel produces a correctly formatted Excel grid."""
    draft = ScheduleDraft(
        department_id=uuid.uuid4(),
        target_month="2026-11",
        source_filename="november_roster.xlsx",
        total_duties=3,
        solver_status="OPTIMAL",
        assignments=[
            {
                "date": "2026-11-01",
                "day_name": "Sunday",
                "doctor_name": "Dr. Gregory House",
                "doctor_email": "house@hospital.org",
                "position": "ER",
                "shift": "Duty",
            },
            {
                "date": "2026-11-02",
                "day_name": "Monday",
                "doctor_name": "Dr. Allison Cameron",
                "doctor_email": "cameron@hospital.org",
                "position": "ER",
                "shift": "Duty",
            },
            {
                "date": "2026-11-03",
                "day_name": "Tuesday",
                "doctor_name": "Dr. Gregory House",
                "doctor_email": "house@hospital.org",
                "position": "ER",
                "shift": "Duty",
            },
        ],
        unavailabilities={"Dr. Gregory House": [14]},
    )

    file_bytes = export_schedule_draft_to_excel(draft)
    assert isinstance(file_bytes, bytes)
    assert len(file_bytes) > 0

    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
    ws = wb.active
    assert ws is not None

    # Check header row 1 (weekdays)
    assert ws.cell(row=1, column=2).value == "Sun"
    assert ws.cell(row=1, column=3).value == "Mon"
    assert ws.cell(row=1, column=4).value == "Tue"

    # Check header row 2 (day numbers & labels)
    assert ws.cell(row=2, column=1).value == "Doctor \\ Date"
    assert ws.cell(row=2, column=2).value == 1
    assert ws.cell(row=2, column=3).value == 2
    assert ws.cell(row=2, column=31).value == 30  # November has 30 days
    assert ws.cell(row=2, column=32).value == "Total Duties"

    # Check doctor names are present in column 1
    doctor_names = [ws.cell(row=r, column=1).value for r in range(3, ws.max_row + 1)]
    assert "Dr. Gregory House" in doctor_names
    assert "Dr. Allison Cameron" in doctor_names

    # Locate Dr. House row
    house_row = 3 if ws.cell(row=3, column=1).value == "Dr. Gregory House" else 4
    cameron_row = 4 if house_row == 3 else 3

    # House: Nov 1 = "Duty", Nov 2 = None, Nov 3 = "Duty", Total = 2
    assert ws.cell(row=house_row, column=2).value == "Duty"
    assert ws.cell(row=house_row, column=3).value is None
    assert ws.cell(row=house_row, column=4).value == "Duty"
    assert ws.cell(row=house_row, column=32).value == 2

    # Cameron: Nov 2 = "Duty", Total = 1
    assert ws.cell(row=cameron_row, column=3).value == "Duty"
    assert ws.cell(row=cameron_row, column=32).value == 1

    # Verify weekend column has fill
    weekend_cell = ws.cell(row=1, column=2)  # Sun Nov 1
    assert weekend_cell.fill.fill_type is not None

    # Verify Dr. House unavailable cell (Nov 14 -> Col 15) has black fill
    house_unavail_cell = ws.cell(row=house_row, column=15)
    assert house_unavail_cell.fill.fill_type == "solid"
    assert str(house_unavail_cell.fill.start_color.rgb).endswith("000000")



def test_export_schedule_draft_empty_assignments():
    """Verify that exporting a draft with no assignments produces valid headers and no crashes."""
    draft = ScheduleDraft(
        department_id=uuid.uuid4(),
        target_month="2026-11",
        source_filename="empty.xlsx",
        total_duties=0,
        solver_status="OPTIMAL",
        assignments=[],
    )

    file_bytes = export_schedule_draft_to_excel(draft)
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
    ws = wb.active

    assert ws.cell(row=2, column=1).value == "Doctor \\ Date"
    assert ws.cell(row=2, column=32).value == "Total Duties"
    assert ws.max_row == 2


def test_export_schedule_draft_multiple_shifts_same_day():
    """Verify that multiple shifts for the same doctor on the same day are joined with commas."""
    draft = ScheduleDraft(
        department_id=uuid.uuid4(),
        target_month="2026-11",
        source_filename="multi.xlsx",
        total_duties=2,
        solver_status="OPTIMAL",
        assignments=[
            {
                "date": "2026-11-05",
                "day_name": "Thursday",
                "doctor_name": "Dr. Gregory House",
                "doctor_email": "house@hospital.org",
                "position": "ER",
                "shift": "Morning",
            },
            {
                "date": "2026-11-05",
                "day_name": "Thursday",
                "doctor_name": "Dr. Gregory House",
                "doctor_email": "house@hospital.org",
                "position": "ER",
                "shift": "Night",
            },
        ],
    )

    file_bytes = export_schedule_draft_to_excel(draft)
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
    ws = wb.active

    # Nov 5 is column 6 (Col 1 is Doctor, Col 2 is Day 1 -> Col 6 is Day 5)
    assert ws.cell(row=3, column=6).value == "Morning, Night"
    assert ws.cell(row=3, column=32).value == 2


def test_export_schedule_draft_invalid_target_month_raises_value_error():
    """Verify that an invalid target_month string raises ValueError."""
    draft = ScheduleDraft(
        department_id=uuid.uuid4(),
        target_month="invalid-month",
        source_filename="test.xlsx",
        total_duties=0,
        solver_status="OPTIMAL",
        assignments=[],
    )

    with pytest.raises(ValueError) as exc_info:
        export_schedule_draft_to_excel(draft)

    assert "Invalid target_month format" in str(exc_info.value)

