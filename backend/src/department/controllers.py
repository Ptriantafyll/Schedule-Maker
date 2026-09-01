"""
Department controller functions for handling business logic.
"""
import uuid
from sqlmodel import Session
from fastapi import HTTPException, status

from src.department.schemas import DepartmentCreate
from src.department.models import Department as DepartmentModel
from src.department import repository


def create_department_controller(department_data: DepartmentCreate, session: Session) -> DepartmentModel:
    """Handles the business logic for creating a new department."""
    existing_dept = repository.get_department_by_name(
        session, department_data.name)
    if existing_dept:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"A department named '{department_data.name}' already exists."
        )
    return repository.create_department(session, department_data)


def get_department_controller(
    department_id: uuid.UUID,
    member_department_id: uuid.UUID,
    session: Session,
) -> DepartmentModel:
    """Handles logic for fetching a department, verifying existence and deletion status."""
    department = repository.get_department_by_id_for_member(
        department_id=department_id,
        member_department_id=member_department_id,
        session=session,
    )
    if not department or department.is_deleted:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Department not found."
        )
    return department


def list_departments_controller(session: Session) -> list[DepartmentModel]:
    """Handles logic for listing active departments."""
    return repository.get_active_departments(session)
