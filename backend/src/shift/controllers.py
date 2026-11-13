"""
Shift controller functions for handling business logic related to team management.
"""

import uuid
from fastapi import HTTPException, status
from sqlmodel import Session
from src.shift import repository
from src.shift.schemas import ShiftCreate, ShiftAssignmentCreate, ShiftUpdate
from src.shift.models import Shift as ShiftModel
from src.shift.models import ShiftAssignment as ShiftAssignmentModel
from src.doctor import repository as doctor_repository
from src.position import repository as position_repository


def create_shift_controller(
    shift_data: ShiftCreate,
    department_id: uuid.UUID,
    session: Session,
) -> ShiftModel:
    """Handles logic for creating a shift"""
    position = position_repository.get_position_by_id_for_department(
        session=session,
        position_id=shift_data.position_id,
        department_id=department_id,
    )
    if position is None:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Position not found."
        )

    existing_shift = repository.get_shift_by_name_for_position(
        session=session,
        shift_name=shift_data.name,
        position_id=shift_data.position_id
    )
    if existing_shift:
        if not existing_shift.is_deleted:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail="Shift already exists"
            )

        return repository.restore_shift(
            session=session,
            shift=existing_shift,
            doctors_per_shift=shift_data.doctors_per_shift,
            grants_day_off=shift_data.grants_day_off,
            commit=True,
        )

    return repository.create_shift(
        session=session,
        shift_data=shift_data
    )


def list_shifts_controller(session: Session, department_id: uuid.UUID) -> list[ShiftModel]:
    """Handles logic for listing all active shifts"""
    return repository.get_active_shifts_for_department(
        session=session,
        department_id=department_id
    )


def get_shift_controller(
    shift_id: uuid.UUID,
    department_id: uuid.UUID,
    session: Session,
) -> ShiftModel:
    """Handles logic for retrieving a shift"""
    shift = repository.get_shift_by_id_for_department(
        session=session,
        shift_id=shift_id,
        department_id=department_id,
    )

    if shift is None:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Shift not found."
        )

    return shift


def create_shift_assignment_controller(
    shift_id: uuid.UUID,
    shift_assignment_data: ShiftAssignmentCreate,
    department_id: uuid.UUID,
    session: Session,
) -> ShiftAssignmentModel:
    """Handles logic for creating a shift assignment"""
    shift = repository.get_shift_by_id_for_department(
        session=session,
        shift_id=shift_id,
        department_id=department_id
    )

    if shift is None:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Shift or Doctor not found."
        )

    doctor = doctor_repository.get_doctor_by_id_for_department(
        session=session,
        doctor_id=shift_assignment_data.doctor_id,
        department_id=department_id
    )

    if doctor is None:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Shift or Doctor not found."
        )

    existing_shift_assignments = repository.get_shift_assignments_by_date(
        session=session,
        shift_id=shift_id,
        target_date=shift_assignment_data.date
    )

    if len(existing_shift_assignments) >= shift.doctors_per_shift:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Shift already has max assignments"
        )

    existing_doctor_assignment = (
        repository.get_active_shift_assignment_for_doctor_on_date(
            session=session,
            doctor_id=shift_assignment_data.doctor_id,
            target_date=shift_assignment_data.date,
        )
    )

    if existing_doctor_assignment is not None:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Doctor is already assigned on this date."
        )

    # todo use the month, or get by id
    doctor_unavailabilities = doctor_repository.get_doctor_unavailability(
        session=session,
        doctor_id=shift_assignment_data.doctor_id
    )

    for unavailability in doctor_unavailabilities:
        if unavailability.date == shift_assignment_data.date:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail=f"Doctor is unavailable on {shift_assignment_data.date}"
            )

    return repository.create_shift_assignment(session, shift_id, shift_assignment_data)


def list_shift_assignments_controller(
    session: Session,
    department_id: uuid.UUID,
) -> list[ShiftAssignmentModel]:
    """Handles logic for retrieving the active shift assignments"""

    return repository.get_active_shift_assignments_for_department(
        session=session,
        department_id=department_id,
    )


def update_shift_controller(
    shift_id: uuid.UUID,
    shift_data: ShiftUpdate,
    department_id: uuid.UUID,
    session: Session,
) -> ShiftModel:
    """Handles logic for updating a specific shift."""
    shift_to_update = repository.get_shift_by_id_for_department(
        session=session,
        shift_id=shift_id,
        department_id=department_id,
    )
    if shift_to_update is None:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Shift not found.",
        )

    update_data = shift_data.model_dump(exclude_unset=True)
    new_name = update_data.get("name")
    new_doctors = update_data.get("doctors_per_shift")
    new_grants_day_off = update_data.get("grants_day_off")

    name_changed = new_name is not None and new_name != shift_to_update.name
    doctors_changed = (
        new_doctors is not None
        and new_doctors != shift_to_update.doctors_per_shift
    )
    grants_day_off_changed = (
        new_grants_day_off is not None
        and new_grants_day_off != shift_to_update.grants_day_off
    )

    if not name_changed and not doctors_changed and not grants_day_off_changed:
        return shift_to_update

    if name_changed:
        existing_shift = repository.get_shift_by_name_for_position(
            session=session,
            shift_name=new_name,
            position_id=shift_to_update.position_id,
        )
        if existing_shift and existing_shift.id != shift_to_update.id:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail=f"A shift named '{new_name}' already exists.",
            )

    return repository.update_shift(
        session=session,
        shift=shift_to_update,
        name=new_name if name_changed else None,
        doctors_per_shift=new_doctors if doctors_changed else None,
        grants_day_off=new_grants_day_off if grants_day_off_changed else None,
        commit=True,
    )


def delete_shift_controller(
    shift_id: uuid.UUID,
    department_id: uuid.UUID,
    session: Session,
) -> None:
    """Handles logic for soft-deleting a shift and upcoming pre-assignments."""
    shift_to_delete = repository.get_shift_by_id_for_department(
        session=session,
        shift_id=shift_id,
        department_id=department_id,
    )
    if shift_to_delete is None:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Shift not found.",
        )

    doctor_repository.soft_delete_upcoming_pre_assignments_for_shifts(
        session=session,
        shift_ids=[shift_id],
        department_id=department_id,
        commit=False,
    )
    repository.soft_delete_shift(
        session=session,
        shift=shift_to_delete,
        commit=False,
    )
    session.commit()
    return None
