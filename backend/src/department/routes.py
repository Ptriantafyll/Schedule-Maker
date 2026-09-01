"""
Module: routes.py
Description: This module defines the API routes for managing hospital departments.
"""
import uuid
from fastapi import APIRouter, Depends, HTTPException, status
from sqlmodel import Session
from src.db.connection import get_session

from src.department.schemas import DepartmentRead
from src.department.controllers import (
    get_department_controller,
    list_departments_controller,
)

from src.user.models import User as UserModel
from src.auth.dependencies import require_super_admin, require_department_member

router = APIRouter(
    prefix="/departments",
    tags=["Departments"]
)


@router.get("/", response_model=list[DepartmentRead])
def list_departments(
    session: Session = Depends(get_session),
    current_user: UserModel = Depends(require_super_admin),
):
    """Retrieves all active, non-deleted hospital departments."""
    return list_departments_controller(session)


@router.get("/{department_id}", response_model=DepartmentRead)
def get_department(
    department_id: uuid.UUID,
    session: Session = Depends(get_session),
    current_user: UserModel = Depends(require_department_member),
):
    """Fetches a specific department by its UUID."""
    if current_user.department_id is None:
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="Invalid account scope."
        )

    return get_department_controller(
        department_id=department_id,
        member_department_id=current_user.department_id,
        session=session,
    )
