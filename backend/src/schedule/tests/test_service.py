"""Unit tests for the schedule service layer."""

import io
import uuid
import datetime
import openpyxl
import pytest
from openpyxl import Workbook
from src.department.models import Department as DepartmentModel
from src.department.schemas import DepartmentCreate
from src.department import repository as department_repository
from src.team import repository as team_repository
from src.position import repository as position_repository
from src.shift import repository as shift_repository
from src.shift.schemas import ShiftCreate
from src.doctor import repository as doctor_repository
from src.doctor.schemas import (
    DoctorPositionCreate,
    DoctorUnavailabilityCreate,
    DoctorPreAssignmentCreate,
)
from sqlmodel import select
from src.user.models import User as UserModel, UserRole
from src.doctor.models import Doctor as DoctorModel, DoctorPosition as DoctorPositionModel
from src.team.models import Team as TeamModel
from src.position.models import Position as PositionModel
from src.models import ScheduleConfig
from src.schedule.models import ScheduleDraft
from src.schedule import repository as schedule_repository
from src.schedule.service import (
    generate_schedule_from_excel,
    generate_schedule_from_roster,
    export_schedule_draft_to_excel,
    publish_schedule_draft,
    preview_excel_roster_diff,
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
    assert draft.status == "one_time"
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


def test_resolve_next_target_month_advances_following_month(session):
    """Verify that target month advances to the month following the latest shift assignment."""
    from src.doctor.models import Doctor as DoctorModel
    from src.position.models import Position as PositionModel
    from src.shift.models import Shift as ShiftModel, ShiftAssignment as ShiftAssignmentModel
    from src.schedule.service import resolve_next_target_month

    dept = DepartmentModel(name="Cardiology", code="CARD")
    session.add(dept)
    session.commit()

    pos = PositionModel(name="Cardio Ward", department_id=dept.id)
    session.add(pos)
    session.commit()

    shift = ShiftModel(name="Morning", position_id=pos.id)
    doc = DoctorModel(name="Dr. Adams", department_id=dept.id)
    session.add(shift)
    session.add(doc)
    session.commit()

    # Latest assignment on May 31, 2026
    assignment = ShiftAssignmentModel(
        doctor_id=doc.id,
        shift_id=shift.id,
        date=datetime.date(2026, 5, 31),
    )
    session.add(assignment)
    session.commit()

    result = resolve_next_target_month(session=session, department_id=dept.id)

    assert result["last_published_month"] == "2026-05"
    assert result["next_target_month"] == "2026-06"


def test_resolve_next_target_month_handles_december_year_rollover(session):
    """Verify that target month rolls over from December to January of the following year."""
    from src.doctor.models import Doctor as DoctorModel
    from src.position.models import Position as PositionModel
    from src.shift.models import Shift as ShiftModel, ShiftAssignment as ShiftAssignmentModel
    from src.schedule.service import resolve_next_target_month

    dept = DepartmentModel(name="Neurology", code="NEURO")
    session.add(dept)
    session.commit()

    pos = PositionModel(name="Neuro Ward", department_id=dept.id)
    session.add(pos)
    session.commit()

    shift = ShiftModel(name="Night", position_id=pos.id)
    doc = DoctorModel(name="Dr. Cuddy", department_id=dept.id)
    session.add(shift)
    session.add(doc)
    session.commit()

    # Assignment on Dec 25, 2026 -> should roll to 2027-01
    assignment = ShiftAssignmentModel(
        doctor_id=doc.id,
        shift_id=shift.id,
        date=datetime.date(2026, 12, 25),
    )
    session.add(assignment)
    session.commit()

    result = resolve_next_target_month(session=session, department_id=dept.id)

    assert result["last_published_month"] == "2026-12"
    assert result["next_target_month"] == "2027-01"


def test_resolve_next_target_month_empty_department_falls_back_to_next_calendar_month(session):
    """Verify that an empty department defaults to next calendar month from today with last_published=None."""
    from src.schedule.service import resolve_next_target_month

    dept = DepartmentModel(name="Pediatrics", code="PED")
    session.add(dept)
    session.commit()

    result = resolve_next_target_month(session=session, department_id=dept.id)

    # Compute expected next calendar month relative to date.today()
    today = datetime.date.today()
    expected_year = today.year if today.month < 12 else today.year + 1
    expected_month = today.month + 1 if today.month < 12 else 1
    expected_target = f"{expected_year:04d}-{expected_month:02d}"

    assert result["last_published_month"] is None
    assert result["next_target_month"] == expected_target


def test_resolve_next_target_month_ignores_soft_deleted_assignments(session):
    """Verify that soft-deleted assignments are ignored during target month calculation."""
    from src.doctor.models import Doctor as DoctorModel
    from src.position.models import Position as PositionModel
    from src.shift.models import Shift as ShiftModel, ShiftAssignment as ShiftAssignmentModel
    from src.schedule.service import resolve_next_target_month

    dept = DepartmentModel(name="Orthopedics", code="ORTHO")
    session.add(dept)
    session.commit()

    pos = PositionModel(name="Ortho Ward", department_id=dept.id)
    session.add(pos)
    session.commit()

    shift = ShiftModel(name="Duty", position_id=pos.id)
    doc = DoctorModel(name="Dr. Wilson", department_id=dept.id)
    session.add(shift)
    session.add(doc)
    session.commit()

    deleted_assignment = ShiftAssignmentModel(
        doctor_id=doc.id,
        shift_id=shift.id,
        date=datetime.date(2026, 10, 15),
        is_deleted=True,
    )
    session.add(deleted_assignment)
    session.commit()

    result = resolve_next_target_month(session=session, department_id=dept.id)

    # Soft-deleted assignment should not count as published
    assert result["last_published_month"] is None


def test_publish_schedule_draft_success(session):
    """Verify publishing a draft materializes ShiftAssignments and marks draft published."""
    from src.schedule.service import publish_schedule_draft
    from src.shift.models import ShiftAssignment as ShiftAssignmentModel

    dept = DepartmentModel(name="General Medicine", code="GEN")
    session.add(dept)
    session.commit()

    draft = ScheduleDraft(
        department_id=dept.id,
        target_month="2026-11",
        source_filename="november_roster.xlsx",
        total_duties=2,
        solver_status="OPTIMAL",
        status="draft",
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
        ],
    )
    session.add(draft)
    session.commit()

    published_draft = publish_schedule_draft(
        session=session,
        draft_id=draft.id,
        department_id=dept.id,
    )

    assert published_draft.status == "published"

    # Verify ShiftAssignment rows created in database
    assignments_in_db = session.query(ShiftAssignmentModel).filter(
        ShiftAssignmentModel.is_deleted == False
    ).all()
    assert len(assignments_in_db) == 2
    dates = {str(a.date) for a in assignments_in_db}
    assert dates == {"2026-11-01", "2026-11-02"}


def test_publish_schedule_draft_already_published_raises_error(session):
    """Verify attempting to publish an already published draft raises ValueError."""
    from src.schedule.service import publish_schedule_draft

    dept = DepartmentModel(name="Surgery", code="SURG")
    session.add(dept)
    session.commit()

    draft = ScheduleDraft(
        department_id=dept.id,
        target_month="2026-11",
        source_filename="november_roster.xlsx",
        status="published",
        assignments=[],
    )
    session.add(draft)
    session.commit()

    with pytest.raises(ValueError, match="already published"):
        publish_schedule_draft(
            session=session,
            draft_id=draft.id,
            department_id=dept.id,
        )


def test_publish_schedule_draft_nonexistent_or_foreign_raises_error(session):
    """Verify attempting to publish a nonexistent or foreign draft raises ValueError."""
    from src.schedule.service import publish_schedule_draft

    dept_a = DepartmentModel(name="Dept A", code="DA")
    dept_b = DepartmentModel(name="Dept B", code="DB")
    session.add_all([dept_a, dept_b])
    session.commit()

    draft = ScheduleDraft(
        department_id=dept_a.id,
        target_month="2026-11",
        source_filename="test.xlsx",
        status="draft",
        assignments=[],
    )
    session.add(draft)
    session.commit()

    # Nonexistent draft ID
    with pytest.raises(ValueError, match="not found"):
        publish_schedule_draft(
            session=session,
            draft_id=uuid.uuid4(),
            department_id=dept_a.id,
        )

    # Department mismatch (caller is dept_b, draft belongs to dept_a)
    with pytest.raises(ValueError, match="not found"):
        publish_schedule_draft(
            session=session,
            draft_id=draft.id,
            department_id=dept_b.id,
        )


def test_publish_schedule_draft_idempotency_cleans_prior_assignments(session):
    """Verify publishing replaces existing assignments for that month (BL-015)."""
    from src.schedule.service import publish_schedule_draft
    from src.doctor.models import Doctor as DoctorModel
    from src.position.models import Position as PositionModel
    from src.shift.models import Shift as ShiftModel, ShiftAssignment as ShiftAssignmentModel

    dept = DepartmentModel(name="ICU Dept", code="ICU")
    session.add(dept)
    session.commit()

    pos = PositionModel(name="ICU", department_id=dept.id)
    session.add(pos)
    session.commit()

    shift = ShiftModel(name="Night", position_id=pos.id)
    doc = DoctorModel(name="Dr. Old", email="old@hospital.org", department_id=dept.id)
    session.add(shift)
    session.add(doc)
    session.commit()

    # Existing assignment in November 2026
    old_assignment = ShiftAssignmentModel(
        doctor_id=doc.id,
        shift_id=shift.id,
        date=datetime.date(2026, 11, 15),
    )
    session.add(old_assignment)
    session.commit()

    # Draft with new assignment on Nov 1
    draft = ScheduleDraft(
        department_id=dept.id,
        target_month="2026-11",
        source_filename="new_nov.xlsx",
        status="draft",
        assignments=[
            {
                "date": "2026-11-01",
                "doctor_name": "Dr. New",
                "doctor_email": "new@hospital.org",
                "position": "ICU",
                "shift": "Night",
            }
        ],
    )
    session.add(draft)
    session.commit()

    publish_schedule_draft(session=session, draft_id=draft.id, department_id=dept.id)

    session.refresh(old_assignment)
    assert old_assignment.is_deleted is True

    active_assignments = session.query(ShiftAssignmentModel).filter(
        ShiftAssignmentModel.is_deleted == False
    ).all()
    assert len(active_assignments) == 1
    assert str(active_assignments[0].date) == "2026-11-01"






def test_export_schedule_draft_to_excel_highlights_public_holidays():
    """Verify that public holidays receive the weekend/holiday fill in exported Excel."""
    draft = ScheduleDraft(
        department_id=uuid.uuid4(),
        target_month="2026-10",
        source_filename="october_roster.xlsx",
        total_duties=1,
        solver_status="OPTIMAL",
        assignments=[
            {
                "date": "2026-10-28",
                "day_name": "Wednesday",
                "doctor_name": "Dr. Gregory House",
                "doctor_email": "house@hospital.org",
                "position": "ER",
                "shift": "Night",
            },
        ],
    )

    file_bytes = export_schedule_draft_to_excel(draft, country_code="GR")
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
    ws = wb.active

    # Oct 28 is column 29 (column 1 is Doctor \ Date, col 2 is Oct 1, ..., col 29 is Oct 28)
    holiday_header_cell = ws.cell(row=1, column=29)
    assert holiday_header_cell.fill.fill_type is not None, "Oct 28 holiday in GR should have holiday/weekend fill"

    # Tuesday Oct 27 (column 28) is a normal weekday and should not have fill
    weekday_header_cell = ws.cell(row=1, column=28)
    assert weekday_header_cell.fill.fill_type is None


def test_export_schedule_draft_to_excel_includes_holiday_duties_column():
    """Verify that export_schedule_draft_to_excel adds a 'Total Holiday Duties' column next to 'Total Duties'."""
    draft = ScheduleDraft(
        department_id=uuid.uuid4(),
        target_month="2026-10",
        source_filename="october_roster.xlsx",
        total_duties=4,
        solver_status="OPTIMAL",
        assignments=[
            # Dr. House:
            # 1. Oct 1 (Thu) - Weekday duty
            {
                "date": "2026-10-01",
                "day_name": "Thursday",
                "doctor_name": "Dr. Gregory House",
                "doctor_email": "house@hospital.org",
                "position": "ER",
                "shift": "Night",
            },
            # 2. Oct 3 (Sat) - Weekend duty (counts as holiday duty)
            {
                "date": "2026-10-03",
                "day_name": "Saturday",
                "doctor_name": "Dr. Gregory House",
                "doctor_email": "house@hospital.org",
                "position": "ER",
                "shift": "Night",
            },
            # 3. Oct 28 (Wed) - Public Holiday duty (Ohi Day in Greece)
            {
                "date": "2026-10-28",
                "day_name": "Wednesday",
                "doctor_name": "Dr. Gregory House",
                "doctor_email": "house@hospital.org",
                "position": "ER",
                "shift": "Night",
            },
            # Dr. Wilson:
            # 4. Oct 2 (Fri) - Weekday duty
            {
                "date": "2026-10-02",
                "day_name": "Friday",
                "doctor_name": "Dr. James Wilson",
                "doctor_email": "wilson@hospital.org",
                "position": "Oncology",
                "shift": "Day",
            },
        ],
    )

    file_bytes = export_schedule_draft_to_excel(draft, country_code="GR")
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
    ws = wb.active

    # October has 31 days:
    # Col 1: Doctor \ Date
    # Col 2-32: Days 1 to 31
    # Col 33: Total Duties (days_in_month + 2)
    # Col 34: Total Holiday Duties (days_in_month + 3)
    total_duties_header = ws.cell(row=2, column=33)
    assert total_duties_header.value == "Total Duties"
    assert total_duties_header.font.bold is True

    holiday_duties_header = ws.cell(row=2, column=34)
    assert holiday_duties_header.value == "Total Holiday Duties"
    assert holiday_duties_header.font.bold is True

    # Row 3: Dr. Gregory House (alphabetical)
    # Total Duties = 3, Total Holiday Duties = 2 (Oct 3 Sat + Oct 28 Holiday)
    assert ws.cell(row=3, column=1).value == "Dr. Gregory House"
    assert ws.cell(row=3, column=33).value == 3
    assert ws.cell(row=3, column=34).value == 2

    # Row 4: Dr. James Wilson
    # Total Duties = 1, Total Holiday Duties = 0
    assert ws.cell(row=4, column=1).value == "Dr. James Wilson"
    assert ws.cell(row=4, column=33).value == 1
    assert ws.cell(row=4, column=34).value == 0


@pytest.fixture(name="realistic_roster")
def realistic_roster_fixture(session, user_factory):
    """Pytest fixture to seed a feasible department roster with multiple positions and shifts."""
    dept = department_repository.create_department(
        session, DepartmentCreate(name="Critical Care", code="CCU")
    )
    team_a = team_repository.create_team(session, name="Team Alpha", department_id=dept.id)

    # 2 Positions: Emergency (Day, Night) and ICU (Duty) across all days
    pos_er = position_repository.create_position(
        session, position_name="Emergency", department_id=dept.id, duty_days=[0, 1, 2, 3, 4, 5, 6]
    )
    pos_icu = position_repository.create_position(
        session, position_name="ICU", department_id=dept.id, duty_days=[0, 1, 2, 3, 4, 5, 6]
    )

    shift_er_day = shift_repository.create_shift(
        session, ShiftCreate(name="Day", position_id=pos_er.id, doctors_per_shift=1, grants_day_off=True)
    )
    shift_er_night = shift_repository.create_shift(
        session, ShiftCreate(name="Night", position_id=pos_er.id, doctors_per_shift=1, grants_day_off=True)
    )
    shift_icu_duty = shift_repository.create_shift(
        session, ShiftCreate(name="Duty", position_id=pos_icu.id, doctors_per_shift=1, grants_day_off=False)
    )

    # 15 doctors (teamed + teamless)
    doctors = []
    for i in range(1, 16):
        team_id = team_a.id if i <= 7 else None
        doc = doctor_repository.create_doctor(
            session, name=f"Dr. Doctor {i}", department_id=dept.id, team_id=team_id
        )
        user_factory(
            email=f"doctor{i}@hospital.org",
            role=UserRole.DOCTOR,
            department_id=dept.id,
            doctor_id=doc.id,
        )
        doctors.append(doc)

    # Doctors 1..5: Emergency only
    for doc in doctors[:5]:
        doctor_repository.create_doctor_position(
            session, doc.id, DoctorPositionCreate(position_id=pos_er.id)
        )

    # Doctors 6..10: ICU only
    for doc in doctors[5:10]:
        doctor_repository.create_doctor_position(
            session, doc.id, DoctorPositionCreate(position_id=pos_icu.id)
        )

    # Doctors 11..15: Multi-position (Emergency AND ICU)
    for doc in doctors[10:15]:
        doctor_repository.create_doctor_position(
            session, doc.id, DoctorPositionCreate(position_id=pos_er.id)
        )
        doctor_repository.create_doctor_position(
            session, doc.id, DoctorPositionCreate(position_id=pos_icu.id)
        )

    # Doctor 1 has unavailability on Nov 10 and pre-assignment on Nov 15
    doctor_repository.create_doctor_unavailability(
        session, doctors[0].id, DoctorUnavailabilityCreate(date=datetime.date(2026, 11, 10))
    )
    doctor_repository.create_doctor_pre_assignment(
        session,
        doctors[0].id,
        DoctorPreAssignmentCreate(shift_id=shift_er_day.id, date=datetime.date(2026, 11, 15)),
    )

    return {
        "dept": dept,
        "pos_er": pos_er,
        "pos_icu": pos_icu,
        "doctors": doctors,
    }


def test_generate_schedule_from_roster_success(session, realistic_roster):
    """Verify solving and persisting a ScheduleDraft directly from the database roster."""
    dept = realistic_roster["dept"]

    draft = generate_schedule_from_roster(
        session=session,
        department_id=dept.id,
        target_month="2026-11",
        country_code="GR",
    )

    assert isinstance(draft, ScheduleDraft)
    assert draft.department_id == dept.id
    assert draft.target_month == "2026-11"
    assert draft.source_filename == "Current Department Roster"
    assert draft.status == "draft"
    assert draft.solver_status in ("OPTIMAL", "FEASIBLE")
    assert draft.total_duties == len(draft.assignments)
    assert len(draft.assignments) > 0

    # Ensure persisted in database
    persisted = session.get(ScheduleDraft, draft.id)
    assert persisted is not None
    assert persisted.total_duties == draft.total_duties

    # Verify assignment structure
    first_assignment = draft.assignments[0]
    assert "date" in first_assignment
    assert "day_name" in first_assignment
    assert "doctor_name" in first_assignment
    assert "doctor_email" in first_assignment
    assert "position" in first_assignment
    assert "shift" in first_assignment


def test_generate_schedule_from_roster_missing_department_raises(session):
    """Verify that a nonexistent department raises ValueError."""
    fake_id = uuid.uuid4()
    with pytest.raises(ValueError, match=f"Department {fake_id} not found"):
        generate_schedule_from_roster(
            session=session,
            department_id=fake_id,
            target_month="2026-11",
        )


def test_generate_schedule_from_roster_validation_error_propagates(session):
    """Verify that roster assembler validation errors propagate up from the service."""
    dept = department_repository.create_department(
        session, DepartmentCreate(name="Empty Dept", code="EMPTY")
    )
    with pytest.raises(ValueError, match="No active positions found"):
        generate_schedule_from_roster(
            session=session,
            department_id=dept.id,
            target_month="2026-11",
        )



def test_generate_schedule_from_excel_defaults_to_one_time_status(session):
    """Verify generate_schedule_from_excel sets status='one_time' when sync_roster is False (default)."""
    dept = DepartmentModel(name="Emergency", code="EMERG")
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

    draft = generate_schedule_from_excel(
        session=session,
        file_bytes=file_bytes,
        department_id=dept.id,
        target_month="2026-11",
        source_filename="november_roster.xlsx",
        sync_roster=False,
    )

    assert draft.status == "one_time"


def test_generate_schedule_from_excel_sync_roster_true_sets_draft_status(session):
    """Verify generate_schedule_from_excel sets status='draft' when sync_roster is True."""
    dept = DepartmentModel(name="Emergency", code="EMERG")
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

    draft = generate_schedule_from_excel(
        session=session,
        file_bytes=file_bytes,
        department_id=dept.id,
        target_month="2026-11",
        source_filename="november_roster.xlsx",
        sync_roster=True,
    )

    assert draft.status == "draft"


def test_publish_schedule_draft_rejects_one_time_status(session):
    """Verify publishing a one_time schedule draft raises ValueError."""
    dept = DepartmentModel(name="Emergency", code="EMERG")
    session.add(dept)
    session.commit()
    session.refresh(dept)

    one_time_draft = ScheduleDraft(
        department_id=dept.id,
        target_month="2026-11",
        source_filename="one_time.xlsx",
        total_duties=1,
        solver_status="OPTIMAL",
        status="one_time",
        assignments=[
            {
                "date": "2026-11-01",
                "doctor_name": "Dr. House",
                "position": "ER",
                "shift": "Duty",
            }
        ],
    )
    saved_draft = schedule_repository.create_schedule_draft(session, one_time_draft)

    with pytest.raises(ValueError, match="One-time Excel schedules are for export only"):
        publish_schedule_draft(
            session=session,
            draft_id=saved_draft.id,
            department_id=dept.id,
        )


def test_repository_queries_exclude_one_time_drafts(session):
    """Verify get_for_department_month and list_schedules exclude one_time drafts, but get_by_id allows them."""
    dept = DepartmentModel(name="Pediatrics", code="PED")
    session.add(dept)
    session.commit()
    session.refresh(dept)

    one_time_draft = ScheduleDraft(
        department_id=dept.id,
        target_month="2026-11",
        source_filename="ad_hoc.xlsx",
        status="one_time",
    )
    regular_draft = ScheduleDraft(
        department_id=dept.id,
        target_month="2026-11",
        source_filename="saved_roster.xlsx",
        status="draft",
    )
    schedule_repository.create_schedule_draft(session, one_time_draft)
    schedule_repository.create_schedule_draft(session, regular_draft)

    # 1. get_schedule_draft_for_department_month ignores one_time and returns regular draft
    active = schedule_repository.get_schedule_draft_for_department_month(
        session=session,
        department_id=dept.id,
        target_month="2026-11",
    )
    assert active is not None
    assert active.id == regular_draft.id

    # 2. list_schedules_for_department excludes one_time
    history = schedule_repository.list_schedules_for_department(
        session=session,
        department_id=dept.id,
    )
    assert len(history) == 1
    assert history[0].id == regular_draft.id

    # 3. get_schedule_draft_by_id still fetches one_time draft for export
    resolved_by_id = schedule_repository.get_schedule_draft_by_id(
        session=session,
        draft_id=one_time_draft.id,
    )
    assert resolved_by_id is not None
    assert resolved_by_id.id == one_time_draft.id


def test_preview_excel_roster_diff_with_differences(session):
    """Verify preview_excel_roster_diff detects missing DB doctors and new Excel doctors."""
    dept = department_repository.create_department(
        session, DepartmentCreate(name="Emergency Medicine", code="EMERG")
    )
    team_alpha = team_repository.create_team(session, name="Alpha", department_id=dept.id)

    # DB has Dr. House (Team Alpha) and Dr. Cuddy (teamless)
    doc_house = doctor_repository.create_doctor(
        session, name="Dr. Gregory House", department_id=dept.id, team_id=team_alpha.id
    )
    doc_cuddy = doctor_repository.create_doctor(
        session, name="Dr. Lisa Cuddy", department_id=dept.id, team_id=None
    )

    # Excel has Dr. House (matching) and Dr. Cameron (new)
    doctor_rows = [
        {"Name": "Dr. Gregory House", "Email": "house@hospital.org", "Position": "ER", "Team": "Alpha"},
        {"Name": "Dr. Allison Cameron", "Email": "cameron@hospital.org", "Position": "ER", "Team": "Alpha"},
    ]
    file_bytes = _create_test_workbook_bytes(doctor_rows)

    diff = preview_excel_roster_diff(
        session=session,
        file_bytes=file_bytes,
        department_id=dept.id,
        target_month="2026-11",
    )

    assert diff.has_existing_roster is True
    # Missing from Excel: Dr. Cuddy
    assert len(diff.missing_from_excel) == 1
    assert diff.missing_from_excel[0].id == doc_cuddy.id
    assert diff.missing_from_excel[0].name == "Dr. Lisa Cuddy"
    assert diff.missing_from_excel[0].team_name is None

    # New in Excel: Dr. Cameron
    assert len(diff.new_in_excel) == 1
    assert diff.new_in_excel[0].name == "Dr. Allison Cameron"
    assert diff.new_in_excel[0].team_name == "Alpha"
    assert "ER" in diff.new_in_excel[0].positions


def test_preview_excel_roster_diff_empty_department(session):
    """Verify preview_excel_roster_diff reports has_existing_roster=False when department has no doctors."""
    dept = department_repository.create_department(
        session, DepartmentCreate(name="Neurology", code="NEURO")
    )

    doctor_rows = [
        {"Name": "Dr. Eric Foreman", "Email": "foreman@hospital.org", "Position": "Neuro", "Team": "Beta"},
    ]
    file_bytes = _create_test_workbook_bytes(doctor_rows)

    diff = preview_excel_roster_diff(
        session=session,
        file_bytes=file_bytes,
        department_id=dept.id,
        target_month="2026-11",
    )

    assert diff.has_existing_roster is False
    assert len(diff.missing_from_excel) == 0
    assert len(diff.new_in_excel) == 1
    assert diff.new_in_excel[0].name == "Dr. Eric Foreman"
    assert diff.new_in_excel[0].team_name == "Beta"


def test_preview_excel_roster_diff_case_and_whitespace_insensitive(session):
    """Verify name comparison normalizes case and surrounding whitespace."""
    dept = department_repository.create_department(
        session, DepartmentCreate(name="Cardiology", code="CARD")
    )
    doctor_repository.create_doctor(
        session, name="Dr. Robert Chase", department_id=dept.id, team_id=None
    )

    # Excel row has irregular casing and whitespace
    doctor_rows = [
        {"Name": "  dr. robert chase  ", "Email": "chase@hospital.org", "Position": "Cardio", "Team": ""},
    ]
    file_bytes = _create_test_workbook_bytes(doctor_rows)

    diff = preview_excel_roster_diff(
        session=session,
        file_bytes=file_bytes,
        department_id=dept.id,
        target_month="2026-11",
    )

    assert diff.has_existing_roster is True
    assert len(diff.missing_from_excel) == 0
    assert len(diff.new_in_excel) == 0


def test_generate_schedule_from_excel_reconciles_removed_doctor_and_preserves_user(session, user_factory):
    """Verify remove_doctor_ids soft-deletes DoctorModel, keeps linked UserModel active, and keeps unselected DB doctors."""
    dept = department_repository.create_department(
        session, DepartmentCreate(name="Emergency Medicine", code="EMERG")
    )

    # 1. DB doctor to be removed, linked to a User account
    doc_to_remove = doctor_repository.create_doctor(
        session, name="Dr. Bob Kelso", department_id=dept.id, team_id=None
    )
    user_factory(
        role=UserRole.DOCTOR,
        department_id=dept.id,
        doctor_id=doc_to_remove.id,
        email="kelso@hospital.org",
    )

    # 2. DB doctor to be KEPT (missing from Excel but admin did not select to remove)
    doc_to_keep = doctor_repository.create_doctor(
        session, name="Dr. John Dorian", department_id=dept.id, team_id=None
    )

    # 3. Excel has 6 doctors (House, Cameron, Foreman, Chase, Wilson, Cuddy)
    doctor_rows = [
        {"Name": "Dr. Gregory House", "Email": "house@hospital.org", "Position": "ER", "Team": "Alpha"},
        {"Name": "Dr. Allison Cameron", "Email": "cameron@hospital.org", "Position": "ER", "Team": "Alpha"},
        {"Name": "Dr. Eric Foreman", "Email": "foreman@hospital.org", "Position": "ER", "Team": "Alpha"},
        {"Name": "Dr. Robert Chase", "Email": "chase@hospital.org", "Position": "ER", "Team": "Beta"},
        {"Name": "Dr. James Wilson", "Email": "wilson@hospital.org", "Position": "ER", "Team": "Beta"},
        {"Name": "Dr. Lisa Cuddy", "Email": "cuddy@hospital.org", "Position": "ER", "Team": "Beta"},
    ]
    file_bytes = _create_test_workbook_bytes(doctor_rows)

    draft = generate_schedule_from_excel(
        session=session,
        file_bytes=file_bytes,
        department_id=dept.id,
        target_month="2026-11",
        source_filename="november_roster.xlsx",
        sync_roster=True,
        remove_doctor_ids=[doc_to_remove.id],
    )

    assert draft.status == "draft"

    # Assert Doctor is soft-deleted
    session.refresh(doc_to_remove)
    assert doc_to_remove.is_deleted is True

    # Assert linked User remains active (Doctor vs User separation)
    linked_user = session.exec(select(UserModel).where(UserModel.doctor_id == doc_to_remove.id)).first()
    assert linked_user is not None
    assert linked_user.is_deleted is False

    # Assert unselected doctor remains active in DB
    session.refresh(doc_to_keep)
    assert doc_to_keep.is_deleted is False


def test_generate_schedule_from_excel_filters_ignored_new_doctors(session):
    """Verify ignored_new_doctors are excluded from the solver assignments and never saved in the database."""
    dept = department_repository.create_department(
        session, DepartmentCreate(name="Trauma Center", code="TRAUMA")
    )

    # Excel contains 6 core doctors + 1 doctor marked to be ignored
    doctor_rows = [
        {"Name": "Dr. Gregory House", "Email": "house@hospital.org", "Position": "ER", "Team": "Alpha"},
        {"Name": "Dr. Allison Cameron", "Email": "cameron@hospital.org", "Position": "ER", "Team": "Alpha"},
        {"Name": "Dr. Eric Foreman", "Email": "foreman@hospital.org", "Position": "ER", "Team": "Alpha"},
        {"Name": "Dr. Robert Chase", "Email": "chase@hospital.org", "Position": "ER", "Team": "Beta"},
        {"Name": "Dr. James Wilson", "Email": "wilson@hospital.org", "Position": "ER", "Team": "Beta"},
        {"Name": "Dr. Lisa Cuddy", "Email": "cuddy@hospital.org", "Position": "ER", "Team": "Beta"},
        {"Name": "Dr. Temporary Resident", "Email": "temp@hospital.org", "Position": "ER", "Team": "Beta"},
    ]
    file_bytes = _create_test_workbook_bytes(doctor_rows)

    draft = generate_schedule_from_excel(
        session=session,
        file_bytes=file_bytes,
        department_id=dept.id,
        target_month="2026-11",
        source_filename="november_roster.xlsx",
        sync_roster=True,
        ignored_new_doctors=["Dr. Temporary Resident"],
    )

    assert draft.status == "draft"

    # Verify Dr. Temporary Resident was NOT saved in DB
    saved_temp = session.exec(
        select(DoctorModel).where(
            DoctorModel.department_id == dept.id,
            DoctorModel.name == "Dr. Temporary Resident",
        )
    ).first()
    assert saved_temp is None

    # Verify Dr. Temporary Resident has no shift assignments
    temp_assignments = [a for a in draft.assignments if a["doctor_name"] == "Dr. Temporary Resident"]
    assert len(temp_assignments) == 0


def test_generate_schedule_from_excel_syncs_new_teams_and_doctors_to_db(session):
    """Verify sync_roster=True upserts teams, positions, doctors, and doctor-position links into DB."""
    dept = department_repository.create_department(
        session, DepartmentCreate(name="Surgery", code="SURG")
    )

    doctor_rows = [
        {"Name": "Dr. Gregory House", "Email": "house@hospital.org", "Position": "General Surgery", "Team": "Alpha"},
        {"Name": "Dr. Allison Cameron", "Email": "cameron@hospital.org", "Position": "General Surgery", "Team": "Alpha"},
        {"Name": "Dr. Eric Foreman", "Email": "foreman@hospital.org", "Position": "General Surgery", "Team": "Alpha"},
        {"Name": "Dr. Robert Chase", "Email": "chase@hospital.org", "Position": "General Surgery", "Team": "Beta"},
        {"Name": "Dr. James Wilson", "Email": "wilson@hospital.org", "Position": "General Surgery", "Team": "Beta"},
        {"Name": "Dr. Lisa Cuddy", "Email": "cuddy@hospital.org", "Position": "General Surgery", "Team": ""},
    ]
    file_bytes = _create_test_workbook_bytes(doctor_rows)

    draft = generate_schedule_from_excel(
        session=session,
        file_bytes=file_bytes,
        department_id=dept.id,
        target_month="2026-11",
        source_filename="november_roster.xlsx",
        sync_roster=True,
    )

    assert draft.status == "draft"

    # 1. Teams synced in DB: Alpha and Beta
    db_teams = team_repository.get_active_teams_by_department(session, dept.id)
    team_names = {t.name for t in db_teams}
    assert "Alpha" in team_names
    assert "Beta" in team_names

    # 2. Position synced in DB: General Surgery
    db_positions = position_repository.get_active_positions_for_department(session, dept.id)
    pos_names = {p.name for p in db_positions}
    assert "General Surgery" in pos_names
    gen_surg_pos = next(p for p in db_positions if p.name == "General Surgery")

    # 3. Doctors synced in DB
    db_doctors = doctor_repository.get_active_doctors_for_department(session, dept.id)
    assert len(db_doctors) == 6
    doc_names = {d.name for d in db_doctors}
    assert "Dr. Gregory House" in doc_names
    assert "Dr. Lisa Cuddy" in doc_names

    # Verify teamless doctor Cuddy has team_id=None
    cuddy_db = next(d for d in db_doctors if d.name == "Dr. Lisa Cuddy")
    assert cuddy_db.team_id is None

    # 4. DoctorPosition links created
    doc_pos_links = session.exec(
        select(DoctorPositionModel).where(DoctorPositionModel.position_id == gen_surg_pos.id)
    ).all()
    assert len(doc_pos_links) == 6
