"""
Module: routes.py
Description: This module defines the API routes for managing hospital departments.
"""
import uuid
from fastapi import APIRouter, Depends, status
from sqlmodel import Session
from src.db.connection import get_session

from src.department.schemas import DepartmentCreate, DepartmentRead
from src.department.controllers import (
    create_department_controller,
    get_department_controller,
    list_departments_controller,
)

from src.user.models import User as UserModel
from src.auth.dependencies import require_department_admin

router = APIRouter(
    prefix="/departments",
    tags=["Departments"]
)


@router.post("/", response_model=DepartmentRead, status_code=status.HTTP_201_CREATED)
def create_department(department_data: DepartmentCreate, session: Session = Depends(get_session)):
    """Creates a new hospital department.

    The backend automatically generates the ID, timestamps, and sync metadata flags.
    """
    return create_department_controller(department_data, session)


@router.get("/", response_model=list[DepartmentRead])
def list_departments(session: Session = Depends(get_session), current_user: UserModel = Depends(require_department_admin)):
    """Retrieves all active, non-deleted hospital departments."""
    return list_departments_controller(session)


@router.get("/{department_id}", response_model=DepartmentRead)
def get_department(department_id: uuid.UUID, session: Session = Depends(get_session)):
    """Fetches a specific department by its UUID."""
    return get_department_controller(department_id, session)
