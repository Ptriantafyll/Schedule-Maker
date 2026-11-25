"""
Doctor repository function for handling database operations.
"""

from typing import Optional, TypeVar
import uuid
import datetime
import calendar
from sqlmodel import Session, not_, select, col
from src.user.models import User as UserModel
from src.doctor.schemas import (
    DoctorPreAssignmentCreate,
    DoctorUnavailabilityCreate,
    DoctorPositionCreate,
)
from src.doctor.models import Doctor as DoctorModel
from src.doctor.models import DoctorPreAssignment as DoctorPreAssignmentModel
from src.doctor.models import DoctorUnavailability as DoctorUnavailabilityModel
from src.doctor.models import DoctorPosition as DoctorPositionModel
from src.position.models import Position as PositionModel
from src.shift.models import ShiftAssignment as ShiftAssignmentModel

_ModelT = TypeVar(
    "_ModelT",
    DoctorModel,
    DoctorPositionModel,
    DoctorPreAssignmentModel,
    DoctorUnavailabilityModel,
)


def _sync_metadata(
    session: Session,
    model: _ModelT,
    *,
    commit: bool = False,
) -> _ModelT:
    """Syncs metadata and commits or flushes"""
    model.sync_status = False
    model.updated_at = datetime.datetime.now(datetime.timezone.utc)

    session.add(model)
    session.flush()
    if commit:
        session.commit()
        session.refresh(model)

    return model


def _set_deleted(
    session: Session,
    model: _ModelT,
    *,
    is_deleted: bool,
    commit: bool = False,
) -> _ModelT:
    """Marks a doctor-domain record as soft-deleted or restored and syncs metadata."""
    model.is_deleted = is_deleted
    return _sync_metadata(session, model, commit=commit)


def get_doctor_by_email(session: Session, email: str) -> DoctorModel | None:
    """Retrieves a doctor by their linked user email"""
    statement = (
        select(DoctorModel)
        .join(UserModel, UserModel.doctor_id == DoctorModel.id)
        .where(
            UserModel.email == email,
            not_(DoctorModel.is_deleted),
            not_(UserModel.is_deleted),
        )
    )
    return session.exec(statement).first()


def get_doctor_by_id_for_department(
    session: Session,
    doctor_id: uuid.UUID,
    department_id: uuid.UUID,
) -> DoctorModel | None:
    """Retrieves a doctor by their unique id"""
    statement = select(DoctorModel).where(
        DoctorModel.id == doctor_id,
        DoctorModel.department_id == department_id,
        not_(DoctorModel.is_deleted),
    )

    return session.exec(statement).first()


def get_doctor_by_name_for_department(
    session: Session,
    name: str,
    department_id: uuid.UUID,
) -> DoctorModel | None:
    """Retrieves a doctor by name within a department (including soft-deleted)."""
    statement = select(DoctorModel).where(
        DoctorModel.department_id == department_id,
        DoctorModel.name == name,
    )

    return session.exec(statement).first()


def get_active_doctors_for_department(session: Session, department_id: uuid.UUID) -> list[DoctorModel]:
    """Retrieves all active doctors"""
    statement = select(DoctorModel).where(
        DoctorModel.department_id == department_id,
        not_(DoctorModel.is_deleted)
    )

    return session.exec(statement).all()


def _normalize_doctor_name(name: str) -> str:
    """
    Normalizes a doctor's name for case-insensitive matching by collapsing whitespace
    and stripping leading 'Dr.' / 'Dr ' prefixes.
    """
    normalized = " ".join(name.strip().lower().split())
    if normalized.startswith("dr."):
        normalized = normalized[3:].strip()
    elif normalized.startswith("dr "):
        normalized = normalized[3:].strip()
    return normalized


def find_unique_unlinked_doctor_by_normalized_name(
    session: Session,
    department_id: uuid.UUID,
    full_name: str,
) -> DoctorModel | None:
    """
    Finds a unique active, unlinked Doctor in the department whose normalized name
    matches full_name. Returns None if zero or multiple unlinked doctors match.
    """
    target_name = _normalize_doctor_name(full_name)
    if not target_name:
        return None

    active_doctors = get_active_doctors_for_department(session, department_id)
    linked_doctor_ids = get_doctor_emails_for_department(
        session, department_id
    )

    matches = [
        doc
        for doc in active_doctors
        if doc.id not in linked_doctor_ids
        and _normalize_doctor_name(doc.name) == target_name
    ]

    if len(matches) == 1:
        return matches[0]

    return None


def stage_doctor(
    session: Session,
    name: str,
    department_id: uuid.UUID,
    team_id: Optional[uuid.UUID] = None,
) -> DoctorModel:
    """Adds a doctor in the db without commiting"""
    new_doctor = DoctorModel(
        name=name,
        department_id=department_id,
        team_id=team_id,
    )
    session.add(new_doctor)
    session.flush()
    return new_doctor


def create_doctor(
    session: Session,
    name: str,
    department_id: uuid.UUID,
    team_id: Optional[uuid.UUID] = None,
) -> DoctorModel:
    """Creates a new doctor in the database"""
    new_doctor = DoctorModel(
        name=name,
        department_id=department_id,
        team_id=team_id,
    )
    session.add(new_doctor)
    session.commit()
    session.refresh(new_doctor)
    return new_doctor


def create_doctor_pre_assignment(session: Session, doctor_id: uuid.UUID, pre_assignment_data: DoctorPreAssignmentCreate) -> DoctorPreAssignmentModel:
    """Creates a new doctor pre assignment in the database"""
    new_pre_assignment = DoctorPreAssignmentModel(
        doctor_id=doctor_id,
        shift_id=pre_assignment_data.shift_id,
        date=pre_assignment_data.date
    )
    session.add(new_pre_assignment)
    session.commit()
    session.refresh(new_pre_assignment)
    return new_pre_assignment

# todo: add a month to get pre assignments for


def get_doctor_pre_assignment_by_date(
    session: Session,
    doctor_id: uuid.UUID | str,
    target_date: datetime.date,
    include_deleted: bool = False,
) -> DoctorPreAssignmentModel | None:
    """Retrieves a doctor's pre-assignment for a given date."""

    statement = select(DoctorPreAssignmentModel).where(
        DoctorPreAssignmentModel.doctor_id == doctor_id,
        DoctorPreAssignmentModel.date == target_date,
    )

    if not include_deleted:
        statement = statement.where(not_(DoctorPreAssignmentModel.is_deleted))

    return session.exec(statement).first()


def get_doctor_pre_assignments(session: Session, doctor_id: str) -> list[DoctorPreAssignmentModel]:
    """Retrieves all the pre assignments of a doctor"""
    statement = (
        select(DoctorPreAssignmentModel).
        where(
            DoctorPreAssignmentModel.doctor_id == doctor_id,
            not_(DoctorPreAssignmentModel.is_deleted)
        )
    )

    return list(session.exec(statement).all())


def get_doctor_pre_assignment_by_id(session: Session, pre_assignment_id: str) -> DoctorPreAssignmentModel:
    """Retrieves a pre assignment by its unique id"""
    return session.get(DoctorPreAssignmentModel, pre_assignment_id)


def get_doctor_pre_assignment_by_id_for_doctor(
    session: Session,
    pre_assignment_id: uuid.UUID,
    doctor_id: uuid.UUID,
) -> DoctorPreAssignmentModel | None:
    """Retrieves an active pre-assignment by its unique ID scoped to a doctor."""
    statement = (
        select(DoctorPreAssignmentModel)
        .where(
            DoctorPreAssignmentModel.id == pre_assignment_id,
            DoctorPreAssignmentModel.doctor_id == doctor_id,
            not_(DoctorPreAssignmentModel.is_deleted)
        )
    )

    return session.exec(statement).first()


def create_doctor_unavailability(session: Session, doctor_id: str, doctor_unavailability_data: DoctorUnavailabilityCreate) -> DoctorUnavailabilityModel:
    """Creates an unavailability for a doctor and date"""
    new_doc_unavailability = DoctorUnavailabilityModel(
        date=doctor_unavailability_data.date,
        doctor_id=doctor_id
    )

    session.add(new_doc_unavailability)
    session.commit()
    session.refresh(new_doc_unavailability)
    return new_doc_unavailability


def get_doctor_unavailability_by_date(
    session: Session,
    doctor_id: uuid.UUID | str,
    target_date: datetime.date,
    include_deleted: bool = False,
) -> DoctorUnavailabilityModel | None:
    """Retrieves the unavailability of a doctor given a date"""

    statement = select(DoctorUnavailabilityModel).where(
        DoctorUnavailabilityModel.doctor_id == doctor_id,
        DoctorUnavailabilityModel.date == target_date,
    )

    if not include_deleted:
        statement = statement.where(not_(DoctorUnavailabilityModel.is_deleted))

    return session.exec(statement).first()

# def get_doctor_unavailability_by_id(session: Session, doctor_id: str) -> list[DoctorUnavailabilityModel]:
#     pass


def get_doctor_unavailability(session: Session, doctor_id: str) -> list[DoctorUnavailabilityModel]:
    """Retrieves the unavailability dates of a doctor"""
    statement = select(DoctorUnavailabilityModel).where(
        DoctorUnavailabilityModel.doctor_id == doctor_id,
        not_(DoctorUnavailabilityModel.is_deleted),
    )

    return session.exec(statement).all()


def get_doctor_unavailability_by_id_for_doctor(
    session: Session,
    unavailability_id: uuid.UUID,
    doctor_id: uuid.UUID,
) -> DoctorUnavailabilityModel | None:
    """Retrieves an active unavailability by its unique ID scoped to a doctor."""
    statement = (
        select(DoctorUnavailabilityModel)
        .where(
            DoctorUnavailabilityModel.id == unavailability_id,
            DoctorUnavailabilityModel.doctor_id == doctor_id,
            not_(DoctorUnavailabilityModel.is_deleted)
        )
    )

    return session.exec(statement).first()


def create_doctor_position(session: Session, doctor_id: str, doctor_pos_data: DoctorPositionCreate) -> DoctorPositionModel:
    """Creates a doctor-position entry"""

    new_doctor_pos = DoctorPositionModel(
        position_id=doctor_pos_data.position_id,
        doctor_id=doctor_id
    )

    session.add(new_doctor_pos)
    session.commit()
    session.refresh(new_doctor_pos)
    return new_doctor_pos


def get_doctor_position_by_id(
    session: Session,
    doctor_id: uuid.UUID | str,
    position_id: uuid.UUID | str,
) -> DoctorPositionModel | None:
    """Retrieves a doctor-position link by doctor ID and position ID (including soft-deleted)."""
    statement = select(DoctorPositionModel).where(
        DoctorPositionModel.doctor_id == doctor_id,
        DoctorPositionModel.position_id == position_id,
    )

    return session.exec(statement).first()


def get_doctor_positions_for_department(
    session: Session,
    doctor_id: uuid.UUID,
    department_id: uuid.UUID,
) -> list[DoctorPositionModel]:
    """Retrieves the active positions of a doctor within a department."""
    statement = (
        select(DoctorPositionModel)
        .join(
            DoctorModel,
            DoctorPositionModel.doctor_id == DoctorModel.id,
        )
        .join(
            PositionModel,
            DoctorPositionModel.position_id == PositionModel.id,
        )
        .where(
            DoctorPositionModel.doctor_id == doctor_id,
            DoctorModel.department_id == department_id,
            not_(DoctorPositionModel.is_deleted),
            not_(DoctorModel.is_deleted),
            not_(PositionModel.is_deleted),
        )
    )

    return session.exec(statement).all()


def get_doctor_emails_for_department(
    session: Session, department_id: uuid.UUID
) -> dict[uuid.UUID, str]:
    """Retrieves a mapping of {doctor_id: email} for all active users in a department."""
    statement = (
        select(UserModel.doctor_id, UserModel.email)
        .where(
            UserModel.department_id == department_id,
            not_(UserModel.doctor_id == None),
            not_(UserModel.is_deleted),
        )
    )
    rows = session.exec(statement).all()

    return {doc_id: email for doc_id, email in rows if doc_id is not None}


def get_all_doctor_positions_for_department(
    session: Session, department_id: uuid.UUID
) -> list[DoctorPositionModel]:
    """Retrieves all active DoctorPosition associations for a department."""
    statement = (
        select(DoctorPositionModel)
        .join(DoctorModel, DoctorModel.id == DoctorPositionModel.doctor_id)
        .join(PositionModel, PositionModel.id == DoctorPositionModel.position_id)
        .where(
            DoctorModel.department_id == department_id,
            not_(DoctorModel.is_deleted),
            not_(PositionModel.is_deleted),
            not_(DoctorPositionModel.is_deleted),
        )
    )

    return list(session.exec(statement).all())


def get_department_unavailabilities_for_month(
    session: Session, department_id: uuid.UUID, year: int, month: int
) -> list[DoctorUnavailabilityModel]:
    """Retrieves all active doctor unavailabilities for a department in a target month."""
    _, last_day = calendar.monthrange(year, month)
    start_date = datetime.date(year, month, 1)
    end_date = datetime.date(year, month, last_day)

    statement = (
        select(DoctorUnavailabilityModel)
        .join(DoctorModel, DoctorModel.id == DoctorUnavailabilityModel.doctor_id)
        .where(
            DoctorModel.department_id == department_id,
            not_(DoctorModel.is_deleted),
            not_(DoctorUnavailabilityModel.is_deleted),
            DoctorUnavailabilityModel.date >= start_date,
            DoctorUnavailabilityModel.date <= end_date,
        )
    )

    return list(session.exec(statement).all())


def get_department_pre_assignments_for_month(
    session: Session, department_id: uuid.UUID, year: int, month: int
) -> list[DoctorPreAssignmentModel]:
    """Retrieves all active doctor pre-assignments for a department in a target month."""
    _, last_day = calendar.monthrange(year, month)
    start_date = datetime.date(year, month, 1)
    end_date = datetime.date(year, month, last_day)

    statement = (
        select(DoctorPreAssignmentModel)
        .join(DoctorModel, DoctorModel.id == DoctorPreAssignmentModel.doctor_id)
        .where(
            DoctorModel.department_id == department_id,
            not_(DoctorModel.is_deleted),
            not_(DoctorPreAssignmentModel.is_deleted),
            DoctorPreAssignmentModel.date >= start_date,
            DoctorPreAssignmentModel.date <= end_date,
        )
    )

    return list(session.exec(statement).all())


def clear_team_from_doctors_for_department(
    session: Session,
    team_id: uuid.UUID,
    department_id: uuid.UUID,
    *,
    commit: bool = False
) -> list[DoctorModel]:
    """Remove a team from doctors of that team in a department."""
    statement = (
        select(DoctorModel)
        .where(
            DoctorModel.department_id == department_id,
            DoctorModel.team_id == team_id,
            not_(DoctorModel.is_deleted),
        )
    )

    doctors_found = list(session.exec(statement).all())

    for doctor in doctors_found:
        doctor.team_id = None
        doctor.sync_status = False
        doctor.updated_at = datetime.datetime.now(datetime.timezone.utc)
        session.add(doctor)

    session.flush()

    if commit:
        session.commit()

    return doctors_found


def soft_delete_doctor_positions_for_position(
    session: Session,
    position_id: uuid.UUID,
    department_id: uuid.UUID,
    *,
    commit: bool = False
) -> list[uuid.UUID]:
    """Soft deletes doctor-position links for a position."""
    statement = (
        select(DoctorPositionModel)
        .join(DoctorModel, DoctorPositionModel.doctor_id == DoctorModel.id)
        .where(
            DoctorPositionModel.position_id == position_id,
            DoctorModel.department_id == department_id,
            not_(DoctorPositionModel.is_deleted),
        )
    )

    links_to_delete = session.exec(statement).all()
    deleted_doc_pos_uuids: list[uuid.UUID] = []
    for doc_pos in links_to_delete:
        _set_deleted(session, doc_pos, is_deleted=True, commit=False)
        deleted_doc_pos_uuids.append(doc_pos.id)

    if commit and deleted_doc_pos_uuids:
        session.commit()

    return deleted_doc_pos_uuids


def soft_delete_upcoming_pre_assignments_for_shifts(
    session: Session,
    shift_ids: list[uuid.UUID],
    department_id: uuid.UUID,
    from_date: datetime.date = None,
    *,
    commit: bool = False,
) -> list[uuid.UUID]:
    """Soft deletes upcoming pre assignments for a shift."""
    if not shift_ids:
        return []

    cutoff_date = from_date or datetime.date.today()

    statement = (
        select(DoctorPreAssignmentModel)
        .join(DoctorModel, DoctorPreAssignmentModel.doctor_id == DoctorModel.id)
        .where(
            DoctorModel.department_id == department_id,
            col(DoctorPreAssignmentModel.shift_id)  # pylint: disable=no-member
            .in_(shift_ids),
            DoctorPreAssignmentModel.date >= cutoff_date,
            not_(DoctorPreAssignmentModel.is_deleted)
        )
    )

    pre_assigments_to_delete = session.exec(statement).all()
    deleted_uuids: list[uuid.UUID] = []
    for pre_assignment in pre_assigments_to_delete:
        _set_deleted(session, pre_assignment, is_deleted=True, commit=False)
        deleted_uuids.append(pre_assignment.id)

    if commit and deleted_uuids:
        session.commit()

    return deleted_uuids


UNSET = object()


def update_doctor(
    session: Session,
    doctor: DoctorModel,
    *,
    name: Optional[str] = None,
    team_id: uuid.UUID | None | object = UNSET,
    commit: bool = False,
) -> DoctorModel:
    """Updates a doctor's name and/or team assignment."""
    if name is not None:
        doctor.name = name

    if team_id is not UNSET:
        doctor.team_id = team_id

    return _sync_metadata(session, doctor, commit=commit)


def soft_delete_doctor(
    session: Session,
    doctor: DoctorModel,
    *,
    commit: bool = False,
) -> DoctorModel:
    """Soft deletes a doctor."""
    return _set_deleted(session, doctor, is_deleted=True, commit=commit)


def soft_delete_doctor_positions_for_doctor(
    session: Session,
    doctor_id: uuid.UUID,
    department_id: uuid.UUID,
    *,
    commit: bool = False,
) -> list[uuid.UUID]:
    """Soft deletes a doctor's active doctor-position links."""
    statement = (
        select(DoctorPositionModel)
        .join(DoctorModel, DoctorPositionModel.doctor_id == DoctorModel.id)
        .where(
            DoctorModel.department_id == department_id,
            DoctorPositionModel.doctor_id == doctor_id,
            not_(DoctorPositionModel.is_deleted),
        )
    )

    doc_pos_to_delete = session.exec(statement).all()
    deleted_uuids: list[uuid.UUID] = []

    for doc_pos in doc_pos_to_delete:
        _set_deleted(session, doc_pos, is_deleted=True, commit=False)
        deleted_uuids.append(doc_pos.id)

    if commit and deleted_uuids:
        session.commit()

    return deleted_uuids


def soft_delete_doctor_position(
    session: Session,
    doctor_position: DoctorPositionModel,
    *,
    commit: bool = False,
) -> DoctorPositionModel:
    """Soft deletes a doctor-position link."""
    return _set_deleted(session, doctor_position, is_deleted=True, commit=commit)


def restore_doctor_position(
    session: Session,
    doctor_position: DoctorPositionModel,
    *,
    commit: bool = False,
) -> DoctorPositionModel:
    """Restores a soft-deleted doctor-position link in place."""
    return _set_deleted(session, doctor_position, is_deleted=False, commit=commit)


def soft_delete_doctor_unavailability(
    session: Session,
    unavailability: DoctorUnavailabilityModel,
    *,
    commit: bool = False,
) -> DoctorUnavailabilityModel:
    """Soft deletes a doctor unavailability entry."""
    return _set_deleted(session, unavailability, is_deleted=True, commit=commit)


def restore_doctor_unavailability(
    session: Session,
    unavailability: DoctorUnavailabilityModel,
    *,
    commit: bool = False,
) -> DoctorUnavailabilityModel:
    """Restores a soft-deleted doctor unavailability entry in place."""
    return _set_deleted(session, unavailability, is_deleted=False, commit=commit)


def soft_delete_doctor_pre_assignment(
    session: Session,
    pre_assignment: DoctorPreAssignmentModel,
    *,
    commit: bool = False,
) -> DoctorPreAssignmentModel:
    """Soft deletes a doctor pre-assignment entry."""
    return _set_deleted(session, pre_assignment, is_deleted=True, commit=commit)


def restore_doctor_pre_assignment(
    session: Session,
    pre_assignment: DoctorPreAssignmentModel,
    shift_id: uuid.UUID,
    *,
    commit: bool = False,
) -> DoctorPreAssignmentModel:
    """Restores a soft-deleted doctor pre-assignment in place with the updated shift."""
    pre_assignment.shift_id = shift_id
    return _set_deleted(session, pre_assignment, is_deleted=False, commit=commit)


def has_active_shift_assignments(
    session: Session,
    doctor_id: uuid.UUID,
) -> bool:
    """Returns True if the doctor has any non-deleted ShiftAssignment rows."""
    statement = select(ShiftAssignmentModel.id).where(
        ShiftAssignmentModel.doctor_id == doctor_id,
        not_(ShiftAssignmentModel.is_deleted),
    )
    return session.exec(statement).first() is not None


def migrate_doctor_schedule_dates(
    session: Session,
    from_doctor_id: uuid.UUID,
    to_doctor_id: uuid.UUID,
    *,
    commit: bool = False,
) -> None:
    """
    Moves active unavailability and pre-assignment dates from from_doctor_id onto
    to_doctor_id (skipping dates that already exist or conflict on to_doctor_id,
    restoring soft-deleted target rows in place when applicable) without committing
    unless commit=True.
    """
    for unavail in get_doctor_unavailability(session, from_doctor_id):
        if get_doctor_pre_assignment_by_date(session, to_doctor_id, unavail.date) is None:
            existing = get_doctor_unavailability_by_date(
                session, to_doctor_id, unavail.date, include_deleted=True
            )
            if existing is None:
                session.add(DoctorUnavailabilityModel(doctor_id=to_doctor_id, date=unavail.date))
                session.flush()
            elif existing.is_deleted:
                restore_doctor_unavailability(session, existing, commit=False)
        soft_delete_doctor_unavailability(session, unavail, commit=False)

    for pre in get_doctor_pre_assignments(session, from_doctor_id):
        if get_doctor_unavailability_by_date(session, to_doctor_id, pre.date) is None:
            existing = get_doctor_pre_assignment_by_date(
                session, to_doctor_id, pre.date, include_deleted=True
            )
            if existing is None:
                session.add(
                    DoctorPreAssignmentModel(
                        doctor_id=to_doctor_id, shift_id=pre.shift_id, date=pre.date
                    )
                )
                session.flush()
            elif existing.is_deleted:
                restore_doctor_pre_assignment(
                    session, existing, shift_id=pre.shift_id, commit=False
                )
        soft_delete_doctor_pre_assignment(session, pre, commit=False)

    if commit:
        session.commit()
