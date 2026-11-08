"""
Assembles department in memory from the database
"""

import uuid
import datetime
from typing import Optional

from sqlmodel import Session
from src.models import Department, Shift, Position, Doctor, Team
from src.models import ScheduleConfig
from src.department import repository as department_repository
from src.position import repository as position_repository
from src.shift import repository as shift_repository
from src.doctor import repository as doctor_repository
from src.team import repository as team_repository


def assemble_department_from_db(
    session: Session,
    department_id: uuid.UUID,
    target_month: str,
    config_override: Optional[ScheduleConfig] = None,
) -> Department:
    """Assembles a department in memory from the database"""
    parts = target_month.split("-")
    if len(parts) != 2 or not parts[0].isdigit() or not parts[1].isdigit() or not (1 <= int(parts[1]) <= 12):
        raise ValueError(
            f"Invalid target month format: expected 'YYYY-MM', got '{target_month}'.")
    year, month = int(parts[0]), int(parts[1])

    dept_model = department_repository.get_department_by_id_global(
        session=session,
        department_id=department_id,
    )
    if not dept_model or dept_model.is_deleted:
        raise ValueError(f"Department {department_id} not found.")

    active_positions = position_repository.get_active_positions_for_department(
        session=session,
        department_id=department_id,
    )
    if not active_positions or len(active_positions) < 1:
        raise ValueError("No active positions found in department roster.")

    active_shifts = shift_repository.get_active_shifts_for_department(
        session=session,
        department_id=department_id,
    )
    shifts_by_position: dict[uuid.UUID, list] = {}
    for shift in active_shifts:
        shifts_by_position.setdefault(shift.position_id, []).append(shift)

    pos_domain_map: dict[uuid.UUID, Position] = {}
    shift_domain_map: dict[uuid.UUID, Shift] = {}

    for pos_model in active_positions:
        pos_shifts = shifts_by_position.get(pos_model.id, [])
        if not pos_shifts:
            raise ValueError(
                f"Position '{pos_model.name}' has no active shifts configured."
            )

        # Convert ShiftModel to domain Shift dataclass
        domain_shifts: list[Shift] = []
        for s in pos_shifts:
            domain_shift = Shift(
                name=s.name,
                doctors_per_shift=s.doctors_per_shift,
                grants_day_off=s.grants_day_off,
            )
            domain_shifts.append(domain_shift)
            shift_domain_map[s.id] = domain_shift

        # Convert PositionModel to domain Position dataclass
        domain_pos = Position(
            name=pos_model.name,
            shifts=domain_shifts,
            duty_days=set(pos_model.duty_days),
            eligible_doctors=[],
        )
        pos_domain_map[pos_model.id] = domain_pos

    active_doctors = doctor_repository.get_active_doctors_for_department(
        session=session,
        department_id=department_id,
    )
    if not active_doctors or len(active_doctors) < 1:
        raise ValueError("No active doctors found in department roster.")

    emails_map = doctor_repository.get_doctor_emails_for_department(
        session=session,
        department_id=department_id,
    )
    doc_positions = doctor_repository.get_all_doctor_positions_for_department(
        session=session,
        department_id=department_id,
    )
    unavailabilities = doctor_repository.get_department_unavailabilities_for_month(
        session=session,
        department_id=department_id,
        year=year,
        month=month,
    )
    pre_assignments = doctor_repository.get_department_pre_assignments_for_month(
        session=session,
        department_id=department_id,
        year=year,
        month=month,
    )

    unavail_by_doc: dict[uuid.UUID, set[datetime.date]] = {}
    for u in unavailabilities:
        unavail_by_doc.setdefault(u.doctor_id, set()).add(u.date)

    pre_assign_by_doc: dict[uuid.UUID, list[tuple[datetime.date, Shift]]] = {}
    for pa in pre_assignments:
        domain_shift = shift_domain_map.get(pa.shift_id)
        if domain_shift:
            pre_assign_by_doc.setdefault(
                pa.doctor_id, []).append((pa.date, domain_shift))

    # Helper map to quickly find a doctor's name by id for error messages
    doc_name_by_id = {doc.id: doc.name for doc in active_doctors}

    # Check for conflicts for each doctor
    for doc_id, unavail_set in unavail_by_doc.items():
        pre_assigned_dates = {date for date,
                              _ in pre_assign_by_doc.get(doc_id, [])}
        conflicts = sorted(unavail_set & pre_assigned_dates)
        if conflicts:
            conflicts_str = ", ".join(d.isoformat() for d in conflicts)
            doc_name = doc_name_by_id.get(doc_id, "Unknown Doctor")
            raise ValueError(
                f"Doctor '{doc_name}' cannot be pre-assigned on unavailable day(s): {conflicts_str}"
            )

    # #### 6. Build Domain Doctors & Assign Position Eligibility
    domain_doctors_by_id: dict[uuid.UUID, Doctor] = {}
    all_domain_doctors: list[Doctor] = []
    active_teams = team_repository.get_active_teams_by_department(
        session=session,
        department_id=department_id,
    )
    teamless_doctors: list[Doctor] = []
    domain_teams_by_id: dict[uuid.UUID, Team] = {}
    domain_teams: list[Team] = []
    for team_model in active_teams:
        domain_team = Team(
            name=team_model.name,
            doctors=[]
        )
        domain_teams.append(domain_team)
        domain_teams_by_id[team_model.id] = domain_team

    for doc_model in active_doctors:
        domain_doc = Doctor(
            name=doc_model.name,
            email=emails_map.get(doc_model.id, ""),
            unavailability=unavail_by_doc.get(doc_model.id, set()),
            pre_assignments=pre_assign_by_doc.get(doc_model.id, []),
        )
        domain_doctors_by_id[doc_model.id] = domain_doc
        all_domain_doctors.append(domain_doc)

        if doc_model.team_id and doc_model.team_id in domain_teams_by_id:
            domain_teams_by_id[doc_model.team_id].doctors.append(domain_doc)
        else:
            teamless_doctors.append(domain_doc)

    # Assign doctors to positions based on DoctorPosition database links
    for dp in doc_positions:
        pos = pos_domain_map.get(dp.position_id)
        doc = domain_doctors_by_id.get(dp.doctor_id)
        if pos and doc and doc not in pos.eligible_doctors:
            pos.eligible_doctors.append(doc)

    # Validate that every position has at least one eligible doctor
    for pos in pos_domain_map.values():
        if len(pos.eligible_doctors) == 0:
            raise ValueError(
                f"Position '{pos.name}' has no eligible doctors assigned."
            )

    return Department(
        name=dept_model.name,
        positions=list(pos_domain_map.values()),
        teams=domain_teams,
        config=config_override or ScheduleConfig(),
        teamless_doctors=teamless_doctors,
        doctor_order=all_domain_doctors,
    )
