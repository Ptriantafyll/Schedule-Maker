"""
Doctor repository function for handling database operations.
"""

from typing import Optional
import uuid
import datetime
import calendar
from sqlmodel import Session, not_, select
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


def get_doctor_pre_assignment_by_date(session: Session, doctor_id: str, target_date: datetime.date) -> DoctorPreAssignmentModel:
    """Retrieves the unavailability of a doctor given a date"""

    statement = select(DoctorPreAssignmentModel).where(
        DoctorPreAssignmentModel.doctor_id == doctor_id,
        DoctorPreAssignmentModel.date == target_date
    )

    return session.exec(statement).first()


def get_doctor_pre_assignments(session: Session, doctor_id: str) -> list[DoctorPreAssignmentModel]:
    """Retrieves all the pre assignments of a doctor"""
    return list(
        session.exec(
            select(DoctorPreAssignmentModel).where(
                DoctorPreAssignmentModel.doctor_id == doctor_id)
        ).all()
    )


def get_doctor_pre_assignment_by_id(session: Session, pre_assignment_id: str) -> DoctorPreAssignmentModel:
    """Retrieves a pre assignment by its unique id"""
    return session.get(DoctorPreAssignmentModel, pre_assignment_id)


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


def get_doctor_unavailability_by_date(session: Session, doctor_id: str, target_date: datetime.date):
    """Retrieves the unavailability of a doctor given a date"""

    statement = select(DoctorUnavailabilityModel).where(
        DoctorUnavailabilityModel.doctor_id == doctor_id,
        DoctorUnavailabilityModel.date == target_date,
    )

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
    """Retrieves the positions of a doctor within a department"""
    statement = (
        select(DoctorPositionModel)
        .join(
            DoctorModel,
            DoctorPositionModel.doctor_id == DoctorModel.id
        )
        .where(
            DoctorPositionModel.doctor_id == doctor_id,
            DoctorModel.department_id == department_id,
            not_(DoctorPositionModel.is_deleted),
            not_(DoctorModel.is_deleted),
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
