"""
Module: routes.py
Description: This module defines the API routes for managing hospital doctors.
"""

import uuid
from fastapi import APIRouter, Depends, status
from sqlmodel import Session
from src.db.connection import get_session

from src.doctor.schemas import (
    DoctorCreate,
    DoctorRead,
    DoctorPreAssignmentCreate,
    DoctorPreAssignmentRead,
    DoctorPositionCreate,
    DoctorPositionRead,
    DoctorUnavailabilityCreate,
    DoctorUnavailabilityRead
)
import src.doctor.controllers as doctor_controllers
from src.user.models import User as UserModel
from src.auth.dependencies import (
    require_department_member,
    require_department_admin,
    require_doctor_or_department_admin,
)

router = APIRouter(
    prefix="/doctors",
    tags=["Doctors"]
)


@router.post("/", response_model=DoctorRead, status_code=status.HTTP_201_CREATED)
def create_doctor(
    doctor_data: DoctorCreate,
    session: Session = Depends(get_session),
    current_user: UserModel = Depends(require_department_admin),
):
    """
    Creates a new doctor
    """
    return doctor_controllers.create_doctor_controller(doctor_data, session)


@router.get("/", response_model=list[DoctorRead])
def list_doctors(
    session: Session = Depends(get_session),
    current_user: UserModel = Depends(require_department_member),
):
    """
    Retrieves all active doctors
    """
    return doctor_controllers.list_doctors_controller(session)


@router.get("/{doctor_id}", response_model=DoctorRead)
def get_doctor(
    doctor_id: uuid.UUID,
    session: Session = Depends(get_session),
    current_user: UserModel = Depends(require_department_member),
):
    """
    Retrieves a doctor by their UUID.
    """
    return doctor_controllers.get_doctor_controller(session=session, doctor_id=doctor_id)


@router.post("/{doctor_id}/pre-assignments", response_model=DoctorPreAssignmentRead, status_code=status.HTTP_201_CREATED)
def create_doctor_pre_assignments(
    doctor_id: uuid.UUID,
    pre_assignment_data: DoctorPreAssignmentCreate,
    session: Session = Depends(get_session),
    current_user: UserModel = Depends(require_department_admin),
):
    """
    Creates pre assignments for a doctor.
    """
    return doctor_controllers.create_doctor_pre_assignment_controller(session, doctor_id, pre_assignment_data)

# todo: add month to get pre assignments for


@router.get("/{doctor_id}/pre-assignments", response_model=list[DoctorPreAssignmentRead])
def list_doctor_pre_assignments(
    doctor_id: uuid.UUID,
    session: Session = Depends(get_session),
    current_user: UserModel = Depends(require_department_admin),
):
    """
    Lists the pre assignment dates of a doctor
    """
    return doctor_controllers.list_doctor_pre_assignments_controller(session, doctor_id)


@router.post("/{doctor_id}/unavailability", response_model=DoctorUnavailabilityRead, status_code=status.HTTP_201_CREATED)
def create_doctor_unavailability(
    doctor_id: uuid.UUID,
    doctor_unavailability_data: DoctorUnavailabilityCreate,
    session: Session = Depends(get_session),
    current_user: UserModel = Depends(require_doctor_or_department_admin),
):
    """
    Creates unavailability for a doctor on a specific date
    """
    return doctor_controllers.create_doctor_unavailabilty_controller(session, doctor_id, doctor_unavailability_data)


@router.get("/{doctor_id}/unavailability", response_model=list[DoctorUnavailabilityRead])
def list_doctor_unavailabilities(
    doctor_id: uuid.UUID,
    session: Session = Depends(get_session),
    current_user: UserModel = Depends(require_doctor_or_department_admin),
):
    """
    Lists the unavailability dates of a doctor
    """
    return doctor_controllers.list_doctor_unavailability_controller(session, doctor_id)
    # todo: make this give a month and return the unav for the month


@router.post("/{doctor_id}/position", response_model=DoctorPositionRead, status_code=status.HTTP_201_CREATED)
def create_doctor_position(
    doctor_id: uuid.UUID,
    doctor_position_data: DoctorPositionCreate,
    session: Session = Depends(get_session),
    current_user: UserModel = Depends(require_department_admin),
):
    """
    Assigns a position to a doctor
    """
    return doctor_controllers.create_doctor_position_controller(session, doctor_id, doctor_position_data)


@router.get("/{doctor_id}/position", response_model=list[DoctorPositionRead])
def list_doctor_positions(
    doctor_id: uuid.UUID,
    session: Session = Depends(get_session),
    current_user: UserModel = Depends(require_department_member),
):
    """
    Lists the positions of a doctor
    """
    return doctor_controllers.list_doctor_positions_controller(session, doctor_id)
