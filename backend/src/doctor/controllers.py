"""
Module doctor.controllers.py

Doctor controller functions for handling business logic related to doctor management.
"""

import uuid
from fastapi import HTTPException, status
from sqlmodel import Session
from src.doctor import repository as doctor_repository
from src.doctor.schemas import DoctorCreate, DoctorPreAssignmentCreate
from src.doctor.models import Doctor as DoctorModel
from src.doctor.models import DoctorPreAssignment as DoctorPreAssignmentModel
from src.department import repository as department_repository
# from src.shift import repository as shift_repository
from src.team import repository as team_repository


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
    # existing_pre_assignment = doctor_repository.get_doctor_pre_assignment_by_id(
    #     session=session, pre_assignment_id=my_id
    # )
    # if existing_pre_assignment:
    #     raise HTTPException(
    #         status_code=status.HTTP_422_UNPROCESSABLE_CONTENT,
    #         detail="Pre assignment already exists"
    #     )

    doctor = doctor_repository.get_doctor_by_id(session, doctor_id)
    # shift = shift_repository.get_shift_by_id
    # todo add shift
    if not doctor:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_CONTENT,
            detail="Doctor does not exist"
        )

    return doctor_repository.create_doctor_pre_assignment(
        session=session,
        doctor_id=doctor_id,
        pre_assignment_data=pre_assignment_data
    )


def list_doctor_pre_assignments(session: Session, doctor_id: uuid.UUID) -> DoctorPreAssignmentModel:
    """List all doctor pre assignments"""
    return doctor_repository.get_doctor_pre_assignments(session, doctor_id)
