"""
Department repository functions for handling database operations.
"""
from typing import Optional
import uuid
from sqlmodel import Session, select, not_

from src.department.schemas import DepartmentCreate
from src.department.models import Department as DepartmentModel


def get_department_by_name(session: Session, name: str) -> Optional[DepartmentModel]:
    """Retrieves a department by its unique name."""
    return session.exec(
        select(DepartmentModel).where(DepartmentModel.name == name)
    ).first()


def get_department_by_id(session: Session, department_id: uuid.UUID) -> Optional[DepartmentModel]:
    """Retrieves a specific department by its UUID."""
    return session.get(DepartmentModel, department_id)


def get_active_departments(session: Session) -> list[DepartmentModel]:
    """Retrieves all active (non-deleted) departments."""
    return list(session.exec(
        select(DepartmentModel).where(not_(DepartmentModel.is_deleted))
    ).all())


def create_department(session: Session, department_data: DepartmentCreate) -> DepartmentModel:
    """Creates a new department in the database."""
    new_department = DepartmentModel(
        name=department_data.name,
        code=department_data.code,
        backup_department_id=department_data.backup_department_id
    )
    session.add(new_department)
    session.commit()
    session.refresh(new_department)
    return new_department


def get_department_by_id_for_member(
    session: Session,
    department_id: uuid.UUID,
    member_department_id: uuid.UUID,
) -> Optional[DepartmentModel]:
    """Retrieves an active department only when it belongs to the member."""
    statement = select(DepartmentModel).where(
        DepartmentModel.id == department_id,
        DepartmentModel.id == member_department_id,
        not_(DepartmentModel.is_deleted),
    )

    return session.exec(statement).first()
