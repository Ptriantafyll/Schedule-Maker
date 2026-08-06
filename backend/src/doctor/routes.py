"""
Module: routes.py
Description: This module defines the API routes for managing hospital doctors.
"""

import uuid
from fastapi import APIRouter, Depends, status
from sqlmodel import Session
from src.db.connection import get_session

from src.doctor.schemas import DoctorCreate, DoctorRead
import src.doctor.controllers as doctor_controllers

router = APIRouter(
    prefix="/doctors",
    tags=["Doctors"]
)


@router.post("/", response_model=DoctorRead, status_code=status.HTTP_201_CREATED)
def create_doctor(doctor_data: DoctorCreate, session: Session = Depends(get_session)):
    """
    Creates a new doctor
    """
    return doctor_controllers.create_doctor_controller(doctor_data, session)

@router.get("/", response_model=list[DoctorRead])
def list_doctors(session: Session = Depends(get_session)):
    """
    Retrieves all active doctors
    """
    return doctor_controllers.list_doctors_controller(session)


@router.get("/{doctor_id}", response_model=DoctorRead)
def get_doctor(doctor_id: uuid.UUID, session: Session = Depends(get_session)):
    """
    Retrieves a doctor by their UUID.
    """
    return doctor_controllers.get_doctor_controller(doctor_id, session)