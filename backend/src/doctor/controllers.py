"""
Module doctor.controllers.py

Doctor controller functions for handling business logic related to doctor management.
"""

import uuid
from fastapi import HTTPException, status
from sqlmodel import Session
from src.doctor import repository as doctor_repository
from src.doctor.schemas import DoctorCreate, DoctorPreAssignmentCreate, DoctorUnavailabilityCreate, DoctorPositionCreate
from src.doctor.models import Doctor as DoctorModel
from src.doctor.models import DoctorPreAssignment as DoctorPreAssignmentModel
from src.doctor.models import DoctorUnavailability as DoctorUnavailabilityModel
from src.doctor.models import DoctorPosition as DoctorPositionModel
from src.department import repository as department_repository
from src.shift import repository as shift_repository
from src.team import repository as team_repository
from src.position import repository as position_repository


def create_doctor_controller(doctor_data: DoctorCreate, session: Session) -> DoctorModel:
    """Handles the business logic for creating a new doctor"""
    existing_doctor = doctor_repository.get_doctor_by_email(
        session, doctor_data.email)
    if existing_doctor:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"A doctor with the email '{doctor_data.name}' already exists."
        )

    team = team_repository.get_team_by_id(session, doctor_data.team_id)
    department = department_repository.get_department_by_id(
        session, doctor_data.department_id)

    if (not team) or (not department):
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_CONTENT,
            detail="The team or department entered does not exist"
        )

    if team.department_id != department.id:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_CONTENT,
            detail=f"The team's department needs to match the doctor's department {team.department_id}, {department.id}"
        )

    return doctor_repository.create_doctor(session, doctor_data)


def list_doctors_controller(session: Session) -> list[DoctorModel]:
    """Handles logic of listing all active doctors"""
    return doctor_repository.get_active_doctors(session)


def get_doctor_controller(session: Session, doctor_id: uuid.UUID) -> DoctorModel:
    """Handles logic for retrieving a specific doctor by their UUID"""

    doctor = doctor_repository.get_doctor_by_id(session, doctor_id)

    if not doctor or doctor.is_deleted:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Doctor not found"
        )

    return doctor


def create_doctor_pre_assignment_controller(session: Session, doctor_id: uuid.UUID, pre_assignment_data: DoctorPreAssignmentCreate) -> DoctorPreAssignmentModel:
    """Handles logic for creating a pre assignment for a doctor"""

    # todo: test existing pre assignment by doctor id and assignment
    existing_pre_assignment = doctor_repository.get_doctor_pre_assignment_by_date(
        session=session,
        target_date=pre_assignment_data.date,
        doctor_id=doctor_id
    )
    if existing_pre_assignment:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Pre assignment already exists"
        )

    doctor = doctor_repository.get_doctor_by_id(session, doctor_id)
    shift = shift_repository.get_shift_by_id(
        session, pre_assignment_data.shift_id)
    if (not doctor) or (not shift):
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_CONTENT,
            detail="Doctor or shift does not exist"
        )

    unavailability = doctor_repository.get_doctor_unavailability_by_date(
        session=session,
        doctor_id=doctor_id,
        target_date=pre_assignment_data.date
    )

    if unavailability:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_CONTENT,
            detail=f"Doctor cannot be assigned to an unavailable day {pre_assignment_data.date}"
        )

    return doctor_repository.create_doctor_pre_assignment(
        session=session,
        doctor_id=doctor_id,
        pre_assignment_data=pre_assignment_data
    )


def list_doctor_pre_assignments_controller(session: Session, doctor_id: uuid.UUID) -> DoctorPreAssignmentModel:
    """List all doctor pre assignments"""
    return doctor_repository.get_doctor_pre_assignments(session, doctor_id)


def create_doctor_unavailabilty_controller(session: Session, doctor_id: uuid.UUID, unavailability_data: DoctorUnavailabilityCreate) -> DoctorUnavailabilityModel:
    """Handles the logic to create a new unavailability for a doctor"""

    existing_doc_unavailability = doctor_repository.get_doctor_unavailability_by_date(
        session=session,
        doctor_id=doctor_id,
        target_date=unavailability_data.date
    )

    if existing_doc_unavailability:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Unavailability already exists"
        )

    doctor = doctor_repository.get_doctor_by_id(session, doctor_id)
    if not doctor:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_CONTENT,
            detail="Doctor does not exist"
        )

    return doctor_repository.create_doctor_unavailability(
        session=session,
        doctor_id=doctor_id,
        doctor_unavailability_data=unavailability_data
    )


def list_doctor_unavailability_controller(session: Session, doctor_id: uuid.UUID) -> list[DoctorUnavailabilityModel]:
    """Handles the logic for listing all the unavailabilities of a doctor"""
    return doctor_repository.get_doctor_unavailability(
        session=session,
        doctor_id=doctor_id
    )


def create_doctor_position_controller(session: Session, doctor_id: uuid.UUID, doctor_pos_data: DoctorPositionCreate) -> DoctorPositionModel:
    """Handles the logic for creating a new doctor-position assosiation"""

    existing_doctor_pos = doctor_repository.get_doctor_position_by_id(
        session=session,
        doctor_id=doctor_id,
        position_id=doctor_pos_data.position_id
    )

    if existing_doctor_pos:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="The doctor is already assigned to this position"
        )

    doctor = doctor_repository.get_doctor_by_id(session, doctor_id)
    if not doctor:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_CONTENT,
            detail="Doctor does not exist"
        )

    position = position_repository.get_position_by_id(
        session, doctor_pos_data.position_id)
    if doctor.department_id != position.department_id:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_CONTENT,
            detail="Doctor's department needs to match position's department"
        )

    return doctor_repository.create_doctor_position(
        session=session,
        doctor_id=doctor_id,
        doctor_pos_data=doctor_pos_data,
    )


def list_doctor_positions_controller(session: Session, doctor_id: uuid.UUID):
    """Handles the logic for retrieving all positions of a doctor"""
    return doctor_repository.get_doctor_positions(
        session=session,
        doctor_id=doctor_id
    )
