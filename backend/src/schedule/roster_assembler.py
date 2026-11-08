"""
Assembles the in-memory solver Department domain graph from database roster records.
"""

import datetime
from typing import Optional
import uuid

from sqlmodel import Session
from src.department import repository as department_repository
from src.department.models import Department as DepartmentModel
from src.doctor import repository as doctor_repository
from src.doctor.models import (
    Doctor as DoctorModel,
    DoctorPosition as DoctorPositionModel,
    DoctorPreAssignment as DoctorPreAssignmentModel,
    DoctorUnavailability as DoctorUnavailabilityModel,
)
from src.models import Department, Doctor, Position, ScheduleConfig, Shift, Team
from src.position import repository as position_repository
from src.position.models import Position as PositionModel
from src.shift import repository as shift_repository
from src.shift.models import Shift as ShiftModel
from src.team import repository as team_repository
from src.team.models import Team as TeamModel


def _parse_target_month(target_month: str) -> tuple[int, int]:
    """Validates a 'YYYY-MM' string and returns (year, month)."""
    parts = target_month.split("-")
    if (
        len(parts) != 2
        or not parts[0].isdigit()
        or not parts[1].isdigit()
        or not (1 <= int(parts[1]) <= 12)
    ):
        raise ValueError(
            f"Invalid target month format: expected 'YYYY-MM', got '{target_month}'."
        )
    return int(parts[0]), int(parts[1])


def _get_active_department(
    session: Session,
    department_id: uuid.UUID,
) -> DepartmentModel:
    """Fetches an active department from the database or raises ValueError."""
    dept_model = department_repository.get_department_by_id_global(
        session=session,
        department_id=department_id,
    )
    if not dept_model or dept_model.is_deleted:
        raise ValueError(f"Department {department_id} not found.")
    return dept_model


def _build_domain_positions_and_shifts(
    position_models: list[PositionModel],
    shift_models: list[ShiftModel],
) -> tuple[dict[uuid.UUID, Position], dict[uuid.UUID, Shift]]:
    """Transforms DB position and shift models into domain dictionaries."""
    if not position_models:
        raise ValueError("No active positions found in department roster.")

    shift_models_by_position_id: dict[uuid.UUID, list[ShiftModel]] = {}
    for shift_model in shift_models:
        shift_models_by_position_id.setdefault(shift_model.position_id, []).append(
            shift_model
        )

    domain_positions_by_id: dict[uuid.UUID, Position] = {}
    domain_shifts_by_id: dict[uuid.UUID, Shift] = {}

    for pos_model in position_models:
        pos_shift_models = shift_models_by_position_id.get(pos_model.id, [])
        if not pos_shift_models:
            raise ValueError(
                f"Position '{pos_model.name}' has no active shifts configured."
            )

        domain_shifts: list[Shift] = []
        for shift_model in pos_shift_models:
            domain_shift = Shift(
                name=shift_model.name,
                doctors_per_shift=shift_model.doctors_per_shift,
                grants_day_off=shift_model.grants_day_off,
            )
            domain_shifts.append(domain_shift)
            domain_shifts_by_id[shift_model.id] = domain_shift

        domain_positions_by_id[pos_model.id] = Position(
            name=pos_model.name,
            shifts=domain_shifts,
            duty_days=set(pos_model.duty_days),
            eligible_doctors=[],
        )

    return domain_positions_by_id, domain_shifts_by_id


def _build_doctor_constraints(
    doctor_models: list[DoctorModel],
    unavail_models: list[DoctorUnavailabilityModel],
    pre_assign_models: list[DoctorPreAssignmentModel],
    domain_shifts_by_id: dict[uuid.UUID, Shift],
) -> tuple[
    dict[uuid.UUID, set[datetime.date]],
    dict[uuid.UUID, list[tuple[datetime.date, Shift]]],
]:
    """Groups unavailabilities and pre-assignments per doctor and checks for date conflicts."""
    unavailabilities_by_doctor_id: dict[uuid.UUID, set[datetime.date]] = {}
    for unavail_model in unavail_models:
        unavailabilities_by_doctor_id.setdefault(
            unavail_model.doctor_id, set()
        ).add(unavail_model.date)

    pre_assignments_by_doctor_id: dict[
        uuid.UUID, list[tuple[datetime.date, Shift]]
    ] = {}
    for pre_assign_model in pre_assign_models:
        domain_shift = domain_shifts_by_id.get(pre_assign_model.shift_id)
        if domain_shift:
            pre_assignments_by_doctor_id.setdefault(
                pre_assign_model.doctor_id, []
            ).append((pre_assign_model.date, domain_shift))

    doctor_names_by_id = {doc_model.id: doc_model.name for doc_model in doctor_models}

    for doc_id, unavail_dates in unavailabilities_by_doctor_id.items():
        pre_assigned_dates = {
            date for date, _ in pre_assignments_by_doctor_id.get(doc_id, [])
        }
        conflicts = sorted(unavail_dates & pre_assigned_dates)
        if conflicts:
            conflicts_str = ", ".join(d.isoformat() for d in conflicts)
            doc_name = doctor_names_by_id.get(doc_id, "Unknown Doctor")
            raise ValueError(
                f"Doctor '{doc_name}' cannot be pre-assigned on unavailable day(s): {conflicts_str}"
            )

    return unavailabilities_by_doctor_id, pre_assignments_by_doctor_id


def _build_domain_doctors(
    doctor_models: list[DoctorModel],
    emails_by_doctor_id: dict[uuid.UUID, str],
    unavailabilities_by_doctor_id: dict[uuid.UUID, set[datetime.date]],
    pre_assignments_by_doctor_id: dict[
        uuid.UUID, list[tuple[datetime.date, Shift]]
    ],
) -> dict[uuid.UUID, Doctor]:
    """Transforms DB doctor models into domain Doctor objects keyed by doctor UUID."""
    if not doctor_models:
        raise ValueError("No active doctors found in department roster.")

    domain_doctors_by_id: dict[uuid.UUID, Doctor] = {}
    for doc_model in doctor_models:
        domain_doctors_by_id[doc_model.id] = Doctor(
            name=doc_model.name,
            email=emails_by_doctor_id.get(doc_model.id, ""),
            unavailability=unavailabilities_by_doctor_id.get(doc_model.id, set()),
            pre_assignments=pre_assignments_by_doctor_id.get(doc_model.id, []),
        )
    return domain_doctors_by_id


def _assign_doctors_to_positions(
    domain_positions_by_id: dict[uuid.UUID, Position],
    domain_doctors_by_id: dict[uuid.UUID, Doctor],
    doc_pos_models: list[DoctorPositionModel],
) -> None:
    """Populates eligible_doctors on each position and validates coverage and pre-assignment eligibility."""
    for doc_pos_model in doc_pos_models:
        domain_pos = domain_positions_by_id.get(doc_pos_model.position_id)
        domain_doc = domain_doctors_by_id.get(doc_pos_model.doctor_id)
        if (
            domain_pos
            and domain_doc
            and domain_doc not in domain_pos.eligible_doctors
        ):
            domain_pos.eligible_doctors.append(domain_doc)

    for domain_pos in domain_positions_by_id.values():
        if not domain_pos.eligible_doctors:
            raise ValueError(
                f"Position '{domain_pos.name}' has no eligible doctors assigned."
            )

    for domain_doc in domain_doctors_by_id.values():
        if not domain_doc.pre_assignments:
            continue
        eligible_shifts = {
            shift
            for domain_pos in domain_positions_by_id.values()
            if domain_doc in domain_pos.eligible_doctors
            for shift in domain_pos.shifts
        }
        for pre_date, pre_shift in domain_doc.pre_assignments:
            if pre_shift not in eligible_shifts:
                raise ValueError(
                    f"Doctor '{domain_doc.name}' is pre-assigned to shift '{pre_shift.name}' "
                    f"on {pre_date.isoformat()}, but is not eligible for its position."
                )


def _group_doctors_into_teams(
    doctor_models: list[DoctorModel],
    domain_doctors_by_id: dict[uuid.UUID, Doctor],
    team_models: list[TeamModel],
) -> tuple[list[Team], list[Doctor]]:
    """Groups domain doctors into their assigned Team or into teamless_doctors."""
    domain_teams: list[Team] = []
    domain_teams_by_id: dict[uuid.UUID, Team] = {}
    for team_model in team_models:
        domain_team = Team(name=team_model.name, doctors=[])
        domain_teams.append(domain_team)
        domain_teams_by_id[team_model.id] = domain_team

    teamless_doctors: list[Doctor] = []
    for doc_model in doctor_models:
        domain_doc = domain_doctors_by_id[doc_model.id]
        if doc_model.team_id and doc_model.team_id in domain_teams_by_id:
            domain_teams_by_id[doc_model.team_id].doctors.append(domain_doc)
        else:
            teamless_doctors.append(domain_doc)

    return domain_teams, teamless_doctors


def assemble_department_from_db(
    session: Session,
    department_id: uuid.UUID,
    target_month: str,
    config_override: Optional[ScheduleConfig] = None,
) -> Department:
    """Assembles a solver Department domain graph from active database roster records."""
    year, month = _parse_target_month(target_month)
    dept_model = _get_active_department(session, department_id)

    # 1. Load & transform Positions and Shifts
    position_models = position_repository.get_active_positions_for_department(
        session=session,
        department_id=department_id,
    )
    shift_models = shift_repository.get_active_shifts_for_department(
        session=session,
        department_id=department_id,
    )
    domain_positions_by_id, domain_shifts_by_id = _build_domain_positions_and_shifts(
        position_models=position_models,
        shift_models=shift_models,
    )

    # 2. Load & transform Doctors and monthly constraints
    doctor_models = doctor_repository.get_active_doctors_for_department(
        session=session,
        department_id=department_id,
    )
    emails_by_doctor_id = doctor_repository.get_doctor_emails_for_department(
        session=session,
        department_id=department_id,
    )
    unavail_models = doctor_repository.get_department_unavailabilities_for_month(
        session=session,
        department_id=department_id,
        year=year,
        month=month,
    )
    pre_assign_models = doctor_repository.get_department_pre_assignments_for_month(
        session=session,
        department_id=department_id,
        year=year,
        month=month,
    )
    unavailabilities_by_doctor_id, pre_assignments_by_doctor_id = (
        _build_doctor_constraints(
            doctor_models=doctor_models,
            unavail_models=unavail_models,
            pre_assign_models=pre_assign_models,
            domain_shifts_by_id=domain_shifts_by_id,
        )
    )
    domain_doctors_by_id = _build_domain_doctors(
        doctor_models=doctor_models,
        emails_by_doctor_id=emails_by_doctor_id,
        unavailabilities_by_doctor_id=unavailabilities_by_doctor_id,
        pre_assignments_by_doctor_id=pre_assignments_by_doctor_id,
    )

    # 3. Link Doctor-Position eligibility and Teams
    doc_pos_models = doctor_repository.get_all_doctor_positions_for_department(
        session=session,
        department_id=department_id,
    )
    _assign_doctors_to_positions(
        domain_positions_by_id=domain_positions_by_id,
        domain_doctors_by_id=domain_doctors_by_id,
        doc_pos_models=doc_pos_models,
    )

    team_models = team_repository.get_active_teams_by_department(
        session=session,
        department_id=department_id,
    )
    domain_teams, teamless_doctors = _group_doctors_into_teams(
        doctor_models=doctor_models,
        domain_doctors_by_id=domain_doctors_by_id,
        team_models=team_models,
    )

    return Department(
        name=dept_model.name,
        positions=list(domain_positions_by_id.values()),
        teams=domain_teams,
        config=config_override or ScheduleConfig(),
        teamless_doctors=teamless_doctors,
        doctor_order=list(domain_doctors_by_id.values()),
    )
