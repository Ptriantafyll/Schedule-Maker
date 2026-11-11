"""
Service layer for schedule operations and CP-SAT solver orchestration.
"""

import io
from typing import Optional
import uuid
import datetime
import calendar
from collections import defaultdict
import holidays
import openpyxl
from openpyxl.styles import PatternFill, Font
from sqlmodel import Session
from ortools.sat.python import cp_model

from src.department import repository as department_repository
from src.models import Department, ScheduleConfig
from src.scheduler import ShiftScheduler
from src.schedule.models import ScheduleDraft
from src.schedule.excel_parser import parse_excel_schedule_workbook
from src.schedule.roster_assembler import assemble_department_from_db
from src.schedule import repository as schedule_repository
from src.shift import repository as shift_repository
from src.team import repository as team_repository

from src.doctor import repository as doctor_repository
from src.doctor.models import Doctor as DoctorModel
from src.doctor.models import DoctorPosition as DoctorPositionModel
from src.position import repository as position_repository
from src.position.models import Position as PositionModel
from src.team.models import Team as TeamModel
from src.shift.models import Shift as ShiftModel, ShiftAssignment as ShiftAssignmentModel
from src.shift.schemas import ShiftCreate
from src.schedule.schemas import ExcelRosterDiffResponse, MissingDoctorDiffItem, NewDoctorDiffItem


def _norm(s: str) -> str:
    """Lowercases, strips, and collapses consecutive whitespace."""
    return " ".join(s.strip().lower().split())


def _extract_solved_assignments(scheduler: ShiftScheduler) -> list[dict]:
    """Extracts and sorts shift assignments from a solved ShiftScheduler instance."""
    assignments = []
    for (day_idx, position, shift, doc), var in scheduler.shift_assignments.items():
        if scheduler.solver.value(var) == 1:
            date = scheduler.dates[day_idx]
            assignments.append({
                "date": date.isoformat(),
                "day_name": date.strftime("%A"),
                "doctor_name": doc.name,
                "doctor_email": doc.email,
                "position": position.name,
                "shift": shift.name,
            })

    assignments.sort(key=lambda a: (a["date"], a["position"], a["shift"]))
    return assignments


def _solve_and_persist_draft(
    session: Session,
    department_id: uuid.UUID,
    dept_domain: Department,
    target_month: str,
    source_filename: str,
    country_code: Optional[str] = "GR",
    draft_status: str = "draft",
) -> ScheduleDraft:
    """Executes CP-SAT solving on a domain Department and persists the ScheduleDraft."""
    year, month = map(int, target_month.split("-"))
    scheduler = ShiftScheduler(
        department=dept_domain, country_code=country_code
    )
    solver_status = scheduler.create_schedule(month=month, year=year)
    status_name = scheduler.solver.status_name(solver_status)

    if solver_status not in (cp_model.FEASIBLE, cp_model.OPTIMAL):
        raise ValueError(
            f"Unable to generate schedule. Constraints are infeasible (status: {status_name})."
        )

    assignments = _extract_solved_assignments(scheduler)
    unavailabilities = {
        doc.name: sorted([d.day for d in doc.unavailability])
        for doc in dept_domain.doctors
        if doc.unavailability
    }

    draft = ScheduleDraft(
        department_id=department_id,
        target_month=target_month,
        source_filename=source_filename,
        total_duties=len(assignments),
        solver_status=status_name,
        status=draft_status,
        assignments=assignments,
        unavailabilities=unavailabilities,
    )

    return schedule_repository.create_schedule_draft(session, draft)


def _sync_roster_from_domain_to_db(
    session: Session,
    department_id: uuid.UUID,
    dept_domain: Department,
) -> None:
    """Syncs roster from domain to db."""
    team_id_by_name: dict[str, uuid.UUID] = {}

    # For each team in dept_domain.teams, find existing TeamModel by name or insert a new one:
    for team in dept_domain.teams:
        team_model = team_repository.get_team_by_name_for_department(
            session=session,
            name=team.name,
            department_id=department_id
        )

        if team_model:
            if team_model.is_deleted:
                team_model.is_deleted = False
                team_model.updated_at = datetime.datetime.now(
                    datetime.timezone.utc
                )
                session.add(team_model)
        else:
            team_model = TeamModel(
                name=team.name,
                department_id=department_id,
            )
            session.add(team_model)

        session.flush()
        team_id_by_name[team.name] = team_model.id

    # Sync positions and shifts
    pos_id_by_name: dict[str, uuid.UUID] = {}
    for pos in dept_domain.positions:
        pos_model = position_repository.get_position_by_name_for_department(
            session=session,
            position_name=pos.name,
            department_id=department_id,
        )
        pos_duty_days = sorted(list(pos.duty_days))
        if pos_model:
            if pos_model.is_deleted:
                pos_model.is_deleted = False
            pos_model.duty_days = pos_duty_days
            pos_model.updated_at = datetime.datetime.now(datetime.timezone.utc)
            session.add(pos_model)
        else:
            pos_model = PositionModel(
                name=pos.name,
                department_id=department_id,
                duty_days=pos_duty_days,
            )
            session.add(pos_model)
        session.flush()
        pos_id_by_name[pos.name] = pos_model.id

        # Sync shifts for this position
        for shift in pos.shifts:
            shift_model = shift_repository.get_shift_by_name_for_position(
                session=session,
                shift_name=shift.name,
                position_id=pos_model.id,
            )
            if shift_model:
                if shift_model.is_deleted:
                    shift_model.is_deleted = False
                shift_model.doctors_per_shift = shift.doctors_per_shift
                shift_model.grants_day_off = shift.grants_day_off
                shift_model.updated_at = datetime.datetime.now(
                    datetime.timezone.utc)
                session.add(shift_model)
            else:
                shift_model = ShiftModel(
                    name=shift.name,
                    position_id=pos_model.id,
                    doctors_per_shift=shift.doctors_per_shift,
                    grants_day_off=shift.grants_day_off,
                )
                session.add(shift_model)
        session.flush()

    # Determine doctor team assignments
    doc_team_id_by_name: dict[str, Optional[uuid.UUID]] = {}
    for team in dept_domain.teams:
        t_id = team_id_by_name.get(team.name)
        for doc in team.doctors:
            doc_team_id_by_name[doc.name] = t_id
    for doc in dept_domain.teamless_doctors:
        doc_team_id_by_name[doc.name] = None

    # Sync Doctors & Link Position Eligibility
    for doc in dept_domain.doctors:
        assigned_team_id = doc_team_id_by_name.get(doc.name)

        doc_model = doctor_repository.get_doctor_by_name_for_department(
            session=session,
            name=doc.name,
            department_id=department_id,
        )

        if doc_model:
            if doc_model.is_deleted:
                doc_model.is_deleted = False
            doc_model.team_id = assigned_team_id
            doc_model.updated_at = datetime.datetime.now(datetime.timezone.utc)
            session.add(doc_model)
        else:
            doc_model = DoctorModel(
                name=doc.name,
                department_id=department_id,
                team_id=assigned_team_id,
            )
            session.add(doc_model)
        session.flush()

        # Link doctor to eligible positions in DoctorPositionModel
        for pos in dept_domain.positions:
            if doc in pos.eligible_doctors:
                pos_id = pos_id_by_name[pos.name]
                link = doctor_repository.get_doctor_position_by_id(
                    session=session,
                    doctor_id=doc_model.id,
                    position_id=pos_id,
                )
                if link:
                    if link.is_deleted:
                        link.is_deleted = False
                        link.updated_at = datetime.datetime.now(
                            datetime.timezone.utc)
                        session.add(link)
                else:
                    link = DoctorPositionModel(
                        doctor_id=doc_model.id,
                        position_id=pos_id,
                    )
                    session.add(link)
        session.flush()


def generate_schedule_from_excel(
    session: Session,
    file_bytes: bytes,
    department_id: uuid.UUID,
    target_month: str,
    source_filename: str = "schedule.xlsx",
    config_override: Optional[ScheduleConfig] = None,
    country_code: Optional[str] = "GR",
    sync_roster: bool = False,
    remove_doctor_ids: Optional[list[uuid.UUID]] = None,
    ignored_new_doctors: Optional[list[str]] = None,
) -> ScheduleDraft:
    """
    Generates an optimized monthly schedule from an uploaded Excel workbook.

    Parses the spreadsheet into domain models, executes the Google OR-Tools CP-SAT
    constraint solver to balance duties across doctors and teams, extracts the
    resulting shift assignments, and persists the draft in the database.

    Args:
        session: Active database session.
        file_bytes: Raw binary bytes of the uploaded .xlsx spreadsheet.
        department_id: UUID of the target hospital department.
        target_month: Target month in "YYYY-MM" format (e.g., "2026-11").
        source_filename: Name of the uploaded file for audit and display tracking.
        config_override: Optional custom solver weights and constraint limits.
        country_code: Country code used for holiday detection (defaults to "GR").
        remove_doctor_ids: List of ids of the doctors that the user marked to remove from the roster,
        ignored_new_doctors: List of doctor names that the user marked to ignore from the roster,

    Returns:
        The newly created and persisted ScheduleDraft record with its assignments.

    Raises:
        ValueError: If the department does not exist, the spreadsheet format is invalid,
                    or the solver determines the constraints are infeasible.
    """
    department = department_repository.get_department_by_id_global(
        session=session,
        department_id=department_id,
    )

    if not department or department.is_deleted:
        raise ValueError(f"Department not found: {department_id}")

    dept_domain = parse_excel_schedule_workbook(
        file_bytes, department.name, target_month
    )
    if config_override:
        dept_domain.config = config_override

    if sync_roster and ignored_new_doctors:
        ignored_set = {_norm(name) for name in ignored_new_doctors}

        # Prune from teams
        for team in dept_domain.teams:
            team.doctors = [d for d in team.doctors if _norm(
                d.name) not in ignored_set]
        # Prune from teamless doctors
        dept_domain.teamless_doctors = [
            d for d in dept_domain.teamless_doctors if _norm(d.name) not in ignored_set
        ]
        # Prune from position eligibility
        for pos in dept_domain.positions:
            pos.eligible_doctors = [
                d for d in pos.eligible_doctors if _norm(d.name) not in ignored_set
            ]
        # Prune from doctor_order
        dept_domain.doctor_order = [
            d for d in dept_domain.doctor_order if _norm(d.name) not in ignored_set
        ]

    if sync_roster and remove_doctor_ids:
        for doc_id in remove_doctor_ids:
            doc = doctor_repository.get_doctor_by_id_for_department(
                session=session,
                doctor_id=doc_id,
                department_id=department_id,
            )
            if doc and not doc.is_deleted:
                doc.is_deleted = True
                doc.updated_at = datetime.datetime.now(datetime.timezone.utc)
                session.add(doc)
        session.flush()

    if sync_roster:
        _sync_roster_from_domain_to_db(
            session=session,
            department_id=department_id,
            dept_domain=dept_domain,
        )

    return _solve_and_persist_draft(
        session=session,
        department_id=department_id,
        dept_domain=dept_domain,
        target_month=target_month,
        source_filename=source_filename,
        country_code=country_code,
        draft_status="draft" if sync_roster else "one_time"
    )


def generate_schedule_from_roster(
    session: Session,
    department_id: uuid.UUID,
    target_month: str,
    config_override: Optional[ScheduleConfig] = None,
    country_code: Optional[str] = "GR",
) -> ScheduleDraft:
    """
    Generates an optimized monthly schedule draft directly from the department database roster.

    Assembles the department roster entities (positions, shifts, doctors, teams,
    unavailabilities, and pre-assignments) from the database into the domain model,
    executes CP-SAT constraint solving, and persists the resulting ScheduleDraft.

    Args:
        session: Active database session.
        department_id: UUID of the target hospital department.
        target_month: Target month in "YYYY-MM" format (e.g., "2026-11").
        config_override: Optional custom solver weights and constraint limits.
        country_code: Country code used for holiday detection (defaults to "GR").

    Returns:
        The newly created and persisted ScheduleDraft record with its assignments.

    Raises:
        ValueError: If the department does not exist, roster validation fails,
                    or the solver determines the constraints are infeasible.
    """
    dept_domain = assemble_department_from_db(
        session=session,
        department_id=department_id,
        target_month=target_month,
        config_override=config_override,
    )

    return _solve_and_persist_draft(
        session=session,
        department_id=department_id,
        dept_domain=dept_domain,
        target_month=target_month,
        source_filename="Current Department Roster",
        country_code=country_code,
    )


def _populate_doctor_day_cell(
    cell,
    day: int,
    doc_unavail_days: set[int],
    shifts: list[str],
    is_weekend: bool,
    unavailability_fill: PatternFill,
    weekend_fill: PatternFill,
) -> int:
    """Populate a single doctor day cell with fill or text, returning duty count."""
    if day in doc_unavail_days:
        cell.fill = unavailability_fill
        return 0

    if shifts:
        cell.value = ", ".join(shifts)

    if is_weekend:
        cell.fill = weekend_fill

    return len(shifts)


def export_schedule_draft_to_excel(draft: ScheduleDraft, country_code: str | None = "GR") -> bytes:
    """Exports a ScheduleDraft model's assignments into a formatted Excel (.xlsx) file in bytes."""
    # 1. Parse target month
    try:
        parts = draft.target_month.split("-")
        year, month = int(parts[0]), int(parts[1])
    except (ValueError, IndexError) as exc:
        raise ValueError(
            f"Invalid target_month format '{draft.target_month}'. Expected YYYY-MM."
        ) from exc

    if month < 1 or month > 12:
        raise ValueError(
            f"Invalid target_month format '{draft.target_month}'. Expected YYYY-MM."
        )

    days_in_month = calendar.monthrange(year, month)[1]

    holiday_dates = set()
    if country_code:
        try:
            holiday_dates = set(holidays.country_holidays(
                country_code, years=year).keys())
        except NotImplementedError:
            holiday_dates = set()

    # 2. Workbook styling setup
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"Schedule {draft.target_month}"

    unavailability_fill = PatternFill(
        start_color="000000", end_color="000000", fill_type="solid"
    )
    weekend_fill = PatternFill(
        start_color="F2F2F2", end_color="F2F2F2", fill_type="solid"
    )
    header_font = Font(bold=True)

    header_cell_1 = ws.cell(row=2, column=1, value="Doctor \\ Date")
    header_cell_1.font = header_font
    header_cell_total = ws.cell(
        row=2, column=days_in_month + 2, value="Total Duties")
    header_cell_total.font = header_font
    header_cell_holiday_total = ws.cell(
        row=2, column=days_in_month + 3, value="Total Holiday Duties"
    )
    header_cell_holiday_total.font = header_font

    # 3. Header rows (1, 2)
    for day_idx in range(1, days_in_month + 1):
        date = datetime.date(year, month, day_idx)
        weekday = date.strftime("%a")
        is_weekend = (date.weekday() >= 5 or date in holiday_dates)

        # Row 1 (weekday name)
        cell = ws.cell(row=1, column=day_idx + 1, value=weekday)
        if is_weekend:
            cell.fill = weekend_fill

        # Row 2 (day number)
        cell = ws.cell(row=2, column=day_idx + 1, value=date.day)
        if is_weekend:
            cell.fill = weekend_fill

    # 4. If no assignments, save and return empty grid
    if not draft.assignments:
        output = io.BytesIO()
        wb.save(output)
        return output.getvalue()

    # 5. Group assignments by (doctor_name, day_num)
    doctor_day_shifts = defaultdict(list)
    doctor_names = set()

    for assignment in draft.assignments:
        doc_name = assignment["doctor_name"]
        doctor_names.add(doc_name)

        day_num = int(assignment["date"].split("-")[2])
        doctor_day_shifts[(doc_name, day_num)].append(
            assignment.get("shift", "Duty")
        )
    sorted_doctors = sorted(list(doctor_names))

    unavailability_dict = draft.unavailabilities or {}

    # 6. Populate doctor rows
    for row_idx, doc_name in enumerate(sorted_doctors, start=3):
        ws.cell(row=row_idx, column=1, value=doc_name)
        doc_unavail_days = set(unavailability_dict.get(doc_name, []))
        total_duties = 0
        total_holiday_duties = 0

        for day in range(1, days_in_month + 1):
            cell = ws.cell(row=row_idx, column=day + 1)
            cell_date = datetime.date(year, month, day)
            is_weekend = (cell_date.weekday() >=
                          5 or cell_date in holiday_dates)
            shifts = doctor_day_shifts.get((doc_name, day), [])

            duties_count = _populate_doctor_day_cell(
                cell=cell,
                day=day,
                doc_unavail_days=doc_unavail_days,
                shifts=shifts,
                is_weekend=is_weekend,
                unavailability_fill=unavailability_fill,
                weekend_fill=weekend_fill,
            )

            total_duties += duties_count
            if is_weekend:
                total_holiday_duties += duties_count

        ws.cell(row=row_idx, column=days_in_month + 2, value=total_duties)
        ws.cell(row=row_idx, column=days_in_month +
                3, value=total_holiday_duties)

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


def _compute_next_month_str(year: int, month: int) -> str:
    """Computes the subsequent calendar month string in YYYY-MM format."""
    if month == 12:
        return f"{year + 1:04d}-01"
    return f"{year:04d}-{month + 1:02d}"


def resolve_next_target_month(
    session: Session,
    department_id: uuid.UUID,
) -> dict[str, str | None]:
    """
    Resolves the target month for schedule generation based on the department's
    latest active ShiftAssignment date. Falls back to next calendar month from today
    if no assignments exist.
    """
    latest_date = shift_repository.get_latest_shift_assignment_date_for_department(
        session=session,
        department_id=department_id,
    )

    if latest_date is not None:
        return {
            "next_target_month": _compute_next_month_str(latest_date.year, latest_date.month),
            "last_published_month": f"{latest_date.year:04d}-{latest_date.month:02d}",
        }

    today = datetime.date.today()
    return {
        "next_target_month": _compute_next_month_str(today.year, today.month),
        "last_published_month": None,
    }


def publish_schedule_draft(
    session: Session,
    draft_id: uuid.UUID,
    department_id: uuid.UUID,
) -> ScheduleDraft:
    """Finalizes a schedule draft and publishes it to the db."""

    draft = schedule_repository.get_schedule_draft_by_id(session, draft_id)
    if draft is None or draft.is_deleted or draft.department_id != department_id:
        raise ValueError(f"Schedule draft not found: {draft_id}.")

    if draft.status == "published":
        raise ValueError(f"Schedule draft {draft_id} is already published.")

    if draft.status == "one_time":
        raise ValueError(
            "One-time Excel schedules are for export only and cannot be published to the department."
        )

    year, month = map(int, draft.target_month.split("-"))

    # 1. Soft-delete prior month assignments (uncommitted flush)
    shift_repository.soft_delete_shift_assignments_for_department_month(
        session=session,
        department_id=department_id,
        year=year,
        month=month,
        commit=False
    )

    # 2. Pre-fetch existing entities to eliminate N+1 database queries
    doc_cache: dict[str, DoctorModel] = {
        d.name: d for d in doctor_repository.get_active_doctors_for_department(session, department_id)
    }
    pos_cache: dict[str, PositionModel] = {
        p.name: p for p in position_repository.get_active_positions_for_department(session, department_id)
    }
    shift_cache: dict[tuple[uuid.UUID, str], ShiftModel] = {
        (s.position_id, s.name): s for s in shift_repository.get_active_shifts_for_department(session, department_id)
    }

    new_assignments: list[ShiftAssignmentModel] = []
    for item in draft.assignments:
        pos_name = item.get("position", "General")
        shift_name = item.get("shift", "Duty")
        doc_name = item.get("doctor_name") or item.get(
            "doctor") or "Unknown Doctor"
        assignment_date = datetime.date.fromisoformat(item["date"])

        # Resolve or auto-provision Position
        if pos_name not in pos_cache:
            pos = position_repository.stage_position(
                session=session,
                position_name=pos_name,
                duty_days=[0, 1, 2, 3, 4, 5, 6],
                department_id=department_id
            )
            pos_cache[pos_name] = pos
        else:
            pos = pos_cache[pos_name]

        # Resolve or auto-provision Shift
        shift_key = (pos.id, shift_name)
        if shift_key not in shift_cache:
            shift = shift_repository.stage_shift(
                session=session,
                shift_data=ShiftCreate(
                    name=shift_name,
                    doctors_per_shift=1,
                    grants_day_off=True,
                    position_id=pos.id
                )
            )
            shift_cache[shift_key] = shift
        else:
            shift = shift_cache[shift_key]

        # Resolve or auto-provision Doctor
        if doc_name not in doc_cache:
            doc = doctor_repository.stage_doctor(
                session=session,
                name=doc_name,
                department_id=department_id,
            )
            doc_cache[doc_name] = doc
        else:
            doc = doc_cache[doc_name]

        new_assignments.append(
            ShiftAssignmentModel(
                doctor_id=doc.id,
                shift_id=shift.id,
                date=assignment_date
            )
        )

    # 3. Bulk insert all assignments (uncommitted flush)
    shift_repository.bulk_create_shift_assignments(
        session=session,
        assignments=new_assignments,
        commit=False,
    )

    draft = schedule_repository.update_schedule_draft_status(
        session=session,
        draft=draft,
        status="published",
        commit=True
    )

    return draft


def preview_excel_roster_diff(
    session: Session,
    file_bytes: bytes,
    department_id: uuid.UUID,
    target_month: str = "2026-11",
) -> ExcelRosterDiffResponse:
    """
    Previews doctor roster differences between an uploaded Excel workbook and the department database.

    Compares active department doctors in the database with doctors in the uploaded
    Excel file (using case- and whitespace-insensitive name matching), identifying
    missing database doctors and new Excel doctors.

    Args:
        session: Active database session.
        file_bytes: Raw binary bytes of the uploaded .xlsx spreadsheet.
        department_id: UUID of the target hospital department.
        target_month: Target month in "YYYY-MM" format (defaults to "2026-11").

    Returns:
        ExcelRosterDiffResponse containing existing roster status, missing doctors,
        and newly introduced doctors with their parsed positions and teams.

    Raises:
        ValueError: If the department does not exist or the Excel file is invalid.
    """
    department = department_repository.get_department_by_id_global(
        session=session,
        department_id=department_id
    )

    if not department or department.is_deleted:
        raise ValueError(
            "Department does not exist."
        )

    db_doctors = doctor_repository.get_active_doctors_for_department(
        session=session,
        department_id=department.id
    )

    db_teams = team_repository.get_active_teams_by_department(
        session=session,
        department_id=department.id
    )

    dept_domain = parse_excel_schedule_workbook(
        file_bytes=file_bytes,
        department_name=department.name,
        target_month=target_month
    )

    # Map normalized Excel doctor name -> Team name (or None)
    excel_team_by_name: dict[str, Optional[str]] = {}
    for team in dept_domain.teams:
        for doc in team.doctors:
            excel_team_by_name[_norm(doc.name)] = team.name
    for doc in dept_domain.teamless_doctors:
        excel_team_by_name[_norm(doc.name)] = None

    # Map normalized Excel doctor name -> list of position names
    excel_positions_by_name: dict[str, list[str]] = defaultdict(list)
    for pos in dept_domain.positions:
        for doc in pos.eligible_doctors:
            excel_positions_by_name[_norm(doc.name)].append(pos.name)

    excel_doc_names = {_norm(doc.name) for doc in dept_domain.doctors}
    db_doc_names = {_norm(doc.name) for doc in db_doctors}

    db_team_name_by_id = {team.id: team.name for team in db_teams}

    missing_from_excel: list[MissingDoctorDiffItem] = []
    for doc_model in db_doctors:
        if _norm(doc_model.name) not in excel_doc_names:
            team_name = db_team_name_by_id.get(
                doc_model.team_id) if doc_model.team_id else None
            missing_from_excel.append(
                MissingDoctorDiffItem(
                    id=doc_model.id,
                    name=doc_model.name,
                    team_name=team_name,
                )
            )

    new_in_excel: list[NewDoctorDiffItem] = []
    for doc in dept_domain.doctors:
        norm = _norm(doc.name)
        if norm not in db_doc_names:
            new_in_excel.append(
                NewDoctorDiffItem(
                    name=doc.name,
                    team_name=excel_team_by_name.get(norm),
                    positions=excel_positions_by_name.get(norm, [])
                )
            )

    return ExcelRosterDiffResponse(
        has_existing_roster=len(db_doctors) > 0,
        missing_from_excel=missing_from_excel,
        new_in_excel=new_in_excel,
    )
