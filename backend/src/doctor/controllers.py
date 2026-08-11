"""
Module doctor.controllers.py

Doctor controller functions for handling business logic related to doctor management.
"""

from fastapi import HTTPException, status
from sqlmodel import Session
from src.doctor import repository as doctor_repository
from src.doctor.schemas import DoctorCreate
from src.doctor.models import Doctor as DoctorModel
from src.department import repository as department_repository
from src.team import repository as team_repository


def create_doctor_controller(doctor_data: DoctorCreate, session: Session) -> DoctorModel:
    """Handles the business logic for creating a new doctor"""
    existing_doctor = doctor_repository.get_doctor_by_email(session, doctor_data.email)
    if existing_doctor:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"A doctor with the email '{doctor_data.name}' already exists."
        )

    team = team_repository.get_team_by_id(session, doctor_data.team_id)
    department = department_repository.get_department_by_id(session, doctor_data.department_id)

    if (not team) or (not department):
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_CONTENT,
            detail="The team or department entered does not exist"
        )

    return doctor_repository.create_doctor(session, doctor_data)
