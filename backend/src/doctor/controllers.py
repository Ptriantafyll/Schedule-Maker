"""
Module doctor.controllers.py

Doctor controller functions for handling business logic related to doctor management.
"""

from collections import defaultdict
import uuid
from fastapi import HTTPException, status
from src.user.models import User as UserModel
from src.user.models import UserRole
from sqlmodel import Session
from src.doctor import repository as doctor_repository
from src.doctor.schemas import (
    DoctorCreate,
    DoctorRead,
    DoctorUpdate,
    DoctorPreAssignmentCreate,
    DoctorUnavailabilityCreate,
    DoctorPositionCreate,
)
from src.doctor.models import Doctor as DoctorModel
from src.doctor.models import DoctorPreAssignment as DoctorPreAssignmentModel
from src.doctor.models import DoctorUnavailability as DoctorUnavailabilityModel
from src.doctor.models import DoctorPosition as DoctorPositionModel
from src.department import repository as department_repository
from src.shift import repository as shift_repository
from src.team import repository as team_repository
from src.position import repository as position_repository
from src.user import repository as user_repository

######################
# Helpers
######################


def _get_authorized_doctor_for_unavailability(
    session: Session,
    doctor_id: uuid.UUID,
    department_id: uuid.UUID,
    current_user: UserModel,
) -> DoctorModel:
    """Returns the Doctor the current user may access for unavailability."""
    if current_user.role == UserRole.DOCTOR:
        if current_user.doctor_id is None:
            raise HTTPException(
                status_code=status.HTTP_403_FORBIDDEN,
                detail="Invalid account scope.",
            )

        if current_user.doctor_id != doctor_id:
            raise HTTPException(
                status_code=status.HTTP_403_FORBIDDEN,
                detail="Cannot access another doctor's unavailability.",
            )
    elif current_user.role != UserRole.DEPARTMENT_ADMIN:
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="Insufficient permissions for this operation.",
        )

    return _get_doctor_or_404(session=session, doctor_id=doctor_id, department_id=department_id)


def _get_doctor_or_404(
    session: Session,
    doctor_id: uuid.UUID,
    department_id: uuid.UUID,
) -> DoctorModel:
    """Retrieves an active doctor in the given department or raises 404."""
    doctor = doctor_repository.get_doctor_by_id_for_department(
        session=session,
        doctor_id=doctor_id,
        department_id=department_id,
    )
    if doctor is None:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Doctor not found.",
        )

    return doctor


def _to_doctor_read(
    doctor: DoctorModel,
    position_ids: list[uuid.UUID] | None = None,
    linked_user: UserModel | None = None,
) -> DoctorRead:
    """Serializes a Doctor ORM model into DoctorRead with position_ids and linked user info."""
    return DoctorRead.model_validate(doctor).model_copy(
        update={
            "position_ids": position_ids if position_ids is not None else [],
            "linked_user_id": linked_user.id if linked_user is not None else None,
            "linked_user_email": linked_user.email if linked_user is not None else None,
        }
    )


def _get_doctor_read_with_positions(
    session: Session,
    doctor: DoctorModel,
    department_id: uuid.UUID,
) -> DoctorRead:
    """Builds DoctorRead including active position IDs and linked user metadata."""
    links = doctor_repository.get_doctor_positions_for_department(
        session=session,
        doctor_id=doctor.id,
        department_id=department_id,
    )
    linked_user = user_repository.get_user_by_doctor_id(session, doctor.id)
    return _to_doctor_read(
        doctor,
        [link.position_id for link in links],
        linked_user=linked_user,
    )

######################
# Controllers
######################


def create_doctor_controller(
    doctor_data: DoctorCreate,
    session: Session,
    department_id: uuid.UUID,
) -> DoctorRead:
    """Handles the business logic for creating a new doctor"""
    if doctor_data.team_id is not None:
        team = team_repository.get_team_by_id_for_department(
            session=session,
            team_id=doctor_data.team_id,
            department_id=department_id,
        )

        if team is None:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="Team not found."
            )

    doctor = doctor_repository.create_doctor(
        session=session,
        name=doctor_data.name,
        team_id=doctor_data.team_id,
        department_id=department_id,
    )
    return _to_doctor_read(doctor, [])


def list_doctors_controller(session: Session, department_id: uuid.UUID) -> list[DoctorRead]:
    """Handles logic of listing all active doctors with their active position IDs and linked users."""
    doctors = doctor_repository.get_active_doctors_for_department(session, department_id)
    links = doctor_repository.get_all_doctor_positions_for_department(session, department_id)
    doctor_users = user_repository.get_active_doctor_users_for_department(session, department_id)

    position_ids_by_doctor: dict[uuid.UUID, list[uuid.UUID]] = defaultdict(list)
    for link in links:
        position_ids_by_doctor[link.doctor_id].append(link.position_id)

    users_by_doctor = {
        user.doctor_id: user
        for user in doctor_users
        if user.doctor_id is not None
    }

    return [
        _to_doctor_read(
            doctor,
            position_ids_by_doctor[doctor.id],
            linked_user=users_by_doctor.get(doctor.id),
        )
        for doctor in doctors
    ]


def list_department_doctor_users_controller(
    session: Session,
    department_id: uuid.UUID,
) -> list[UserModel]:
    """Retrieves active DOCTOR user accounts in the department."""
    return user_repository.get_active_doctor_users_for_department(
        session=session,
        department_id=department_id,
    )


def list_doctor_roster_controller(session: Session, department_id: uuid.UUID) -> list[DoctorModel]:
    """Handles logic of listing all active doctors with their positions"""
    return doctor_repository.get_active_doctors_for_department(session, department_id)


def get_doctor_controller(
    session: Session,
    doctor_id: uuid.UUID,
    department_id: uuid.UUID,
) -> DoctorRead:
    """Handles logic for retrieving a specific doctor by their UUID"""
    doctor = _get_doctor_or_404(
        session=session,
        doctor_id=doctor_id,
        department_id=department_id,
    )
    return _get_doctor_read_with_positions(session, doctor, department_id)


def update_doctor_controller(
    session: Session,
    doctor_id: uuid.UUID,
    department_id: uuid.UUID,
    doctor_data: DoctorUpdate,
) -> DoctorRead:
    """Updates a doctor's name and/or team assignment within their department."""
    doctor = _get_doctor_or_404(
        session=session,
        doctor_id=doctor_id,
        department_id=department_id,
    )

    update_fields = doctor_data.model_dump(exclude_unset=True)
    if "team_id" in update_fields and update_fields["team_id"] is not None:
        team = team_repository.get_team_by_id_for_department(
            session=session,
            team_id=update_fields["team_id"],
            department_id=department_id,
        )
        if team is None:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="Team not found.",
            )

    team_id_arg = (
        update_fields["team_id"]
        if "team_id" in update_fields
        else doctor_repository.UNSET
    )
    updated_doctor = doctor_repository.update_doctor(
        session=session,
        doctor=doctor,
        name=update_fields.get("name"),
        team_id=team_id_arg,
        commit=True,
    )
    return _get_doctor_read_with_positions(session, updated_doctor, department_id)


def delete_doctor_controller(
    session: Session,
    doctor_id: uuid.UUID,
    department_id: uuid.UUID,
) -> None:
    """Atomically soft-deletes a doctor and their active position links, keeping linked Users active."""
    doctor = _get_doctor_or_404(
        session=session,
        doctor_id=doctor_id,
        department_id=department_id,
    )

    try:
        doctor_repository.soft_delete_doctor_positions_for_doctor(
            session=session,
            doctor_id=doctor.id,
            department_id=department_id,
            commit=False,
        )
        doctor_repository.soft_delete_doctor(
            session=session,
            doctor=doctor,
            commit=False,
        )
        session.commit()
    except Exception:
        session.rollback()
        raise


def create_doctor_pre_assignment_controller(
    session: Session,
    doctor_id: uuid.UUID,
    department_id: uuid.UUID,
    pre_assignment_data: DoctorPreAssignmentCreate,
) -> DoctorPreAssignmentModel:
    """Handles logic for creating a pre assignment for a doctor"""
    doctor = doctor_repository.get_doctor_by_id_for_department(
        session=session,
        doctor_id=doctor_id,
        department_id=department_id,
    )
    shift = shift_repository.get_shift_by_id_for_department(
        session=session,
        shift_id=pre_assignment_data.shift_id,
        department_id=department_id
    )
    if doctor is None or shift is None:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Doctor or shift not found."
        )

    existing_pre_assignment = doctor_repository.get_doctor_pre_assignment_by_date(
        session=session,
        target_date=pre_assignment_data.date,
        doctor_id=doctor_id,
        include_deleted=True,
    )
    if existing_pre_assignment and not existing_pre_assignment.is_deleted:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Pre assignment already exists"
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

    if existing_pre_assignment:
        return doctor_repository.restore_doctor_pre_assignment(
            session=session,
            pre_assignment=existing_pre_assignment,
            shift_id=pre_assignment_data.shift_id,
            commit=True,
        )

    return doctor_repository.create_doctor_pre_assignment(
        session=session,
        doctor_id=doctor_id,
        pre_assignment_data=pre_assignment_data
    )


def list_doctor_pre_assignments_controller(
    session: Session,
    doctor_id: uuid.UUID,
    department_id: uuid.UUID,
) -> list[DoctorPreAssignmentModel]:
    """List all doctor pre assignments"""
    doctor = _get_doctor_or_404(
        session=session,
        doctor_id=doctor_id,
        department_id=department_id,
    )

    return doctor_repository.get_doctor_pre_assignments(session, doctor.id)


def delete_doctor_pre_assignment_controller(
    session: Session,
    doctor_id: uuid.UUID,
    pre_assignment_id: uuid.UUID,
    department_id: uuid.UUID,
) -> None:
    """Soft-deletes a pre-assignment belonging to an active doctor in the department."""
    doctor = _get_doctor_or_404(
        session=session,
        doctor_id=doctor_id,
        department_id=department_id,
    )
    pre_assignment = doctor_repository.get_doctor_pre_assignment_by_id_for_doctor(
        session=session,
        pre_assignment_id=pre_assignment_id,
        doctor_id=doctor.id,
    )
    if pre_assignment is None:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Pre-assignment not found.",
        )

    doctor_repository.soft_delete_doctor_pre_assignment(
        session=session,
        pre_assignment=pre_assignment,
        commit=True,
    )


def create_doctor_unavailability_controller(
    session: Session,
    doctor_id: uuid.UUID,
    department_id: uuid.UUID,
    unavailability_data: DoctorUnavailabilityCreate,
    current_user: UserModel,
) -> DoctorUnavailabilityModel:
    """Handles the logic to create a new unavailability for a doctor"""
    authorized_doctor = _get_authorized_doctor_for_unavailability(
        session=session,
        doctor_id=doctor_id,
        department_id=department_id,
        current_user=current_user
    )

    existing_doc_unavailability = doctor_repository.get_doctor_unavailability_by_date(
        session=session,
        doctor_id=authorized_doctor.id,
        target_date=unavailability_data.date,
        include_deleted=True,
    )

    if existing_doc_unavailability and not existing_doc_unavailability.is_deleted:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Unavailability already exists"
        )

    existing_pre_assignment = doctor_repository.get_doctor_pre_assignment_by_date(
        session=session,
        doctor_id=authorized_doctor.id,
        target_date=unavailability_data.date
    )

    if existing_pre_assignment:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_CONTENT,
            detail=f"Doctor cannot be marked unavailable on a pre-assigned day {unavailability_data.date}"
        )

    if existing_doc_unavailability:
        return doctor_repository.restore_doctor_unavailability(
            session=session,
            unavailability=existing_doc_unavailability,
            commit=True,
        )

    return doctor_repository.create_doctor_unavailability(
        session=session,
        doctor_id=authorized_doctor.id,
        doctor_unavailability_data=unavailability_data
    )


def list_doctor_unavailability_controller(
    session: Session,
    doctor_id: uuid.UUID,
    current_user: UserModel,
    department_id: uuid.UUID,
) -> list[DoctorUnavailabilityModel]:
    """Handles the logic for listing all the unavailabilities of a doctor"""
    authorized_doctor = _get_authorized_doctor_for_unavailability(
        session=session,
        doctor_id=doctor_id,
        department_id=department_id,
        current_user=current_user
    )

    return doctor_repository.get_doctor_unavailability(
        session=session,
        doctor_id=authorized_doctor.id
    )


def delete_doctor_unavailability_controller(
    session: Session,
    doctor_id: uuid.UUID,
    unavailability_id: uuid.UUID,
    department_id: uuid.UUID,
    current_user: UserModel,
) -> None:
    """Soft-deletes a doctor's unavailability entry after enforcing role and ownership."""
    authorized_doctor = _get_authorized_doctor_for_unavailability(
        session=session,
        doctor_id=doctor_id,
        department_id=department_id,
        current_user=current_user,
    )
    unavailability = doctor_repository.get_doctor_unavailability_by_id_for_doctor(
        session=session,
        unavailability_id=unavailability_id,
        doctor_id=authorized_doctor.id,
    )
    if unavailability is None:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Unavailability not found.",
        )

    doctor_repository.soft_delete_doctor_unavailability(
        session=session,
        unavailability=unavailability,
        commit=True,
    )


def create_doctor_position_controller(
    session: Session,
    doctor_id: uuid.UUID,
    doctor_pos_data: DoctorPositionCreate,
    department_id: uuid.UUID,
) -> DoctorPositionModel:
    """Handles the logic for creating a new doctor-position association"""
    doctor = doctor_repository.get_doctor_by_id_for_department(
        session=session,
        doctor_id=doctor_id,
        department_id=department_id,
    )
    position = position_repository.get_position_by_id_for_department(
        session=session,
        position_id=doctor_pos_data.position_id,
        department_id=department_id,
    )

    if doctor is None or position is None:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Doctor or position not found."
        )

    existing_doctor_pos = doctor_repository.get_doctor_position_by_id(
        session=session,
        doctor_id=doctor_id,
        position_id=doctor_pos_data.position_id
    )

    if existing_doctor_pos:
        if not existing_doctor_pos.is_deleted:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail="The doctor is already assigned to this position"
            )
        return doctor_repository.restore_doctor_position(
            session=session,
            doctor_position=existing_doctor_pos,
            commit=True,
        )

    return doctor_repository.create_doctor_position(
        session=session,
        doctor_id=doctor_id,
        doctor_pos_data=doctor_pos_data,
    )


def list_doctor_positions_controller(
    session: Session,
    doctor_id: uuid.UUID,
    department_id: uuid.UUID,
) -> list[DoctorPositionModel]:
    """Handles the logic for retrieving all positions of a doctor"""
    doctor = _get_doctor_or_404(
        session=session,
        doctor_id=doctor_id,
        department_id=department_id,
    )

    return doctor_repository.get_doctor_positions_for_department(
        session=session,
        doctor_id=doctor.id,
        department_id=department_id,
    )


def delete_doctor_position_controller(
    session: Session,
    doctor_id: uuid.UUID,
    position_id: uuid.UUID,
    department_id: uuid.UUID,
) -> None:
    """Soft-deletes a doctor-position eligibility link within the department."""
    doctor = _get_doctor_or_404(
        session=session,
        doctor_id=doctor_id,
        department_id=department_id,
    )
    doctor_position = doctor_repository.get_doctor_position_by_id(
        session=session,
        doctor_id=doctor.id,
        position_id=position_id,
    )
    if doctor_position is None or doctor_position.is_deleted:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Doctor position not found.",
        )

    doctor_repository.soft_delete_doctor_position(
        session=session,
        doctor_position=doctor_position,
        commit=True,
    )


def link_doctor_to_user_controller(
    session: Session,
    doctor_id: uuid.UUID,
    user_id: uuid.UUID,
    department_id: uuid.UUID,
) -> DoctorRead:
    """
    Atomically links an active DOCTOR user in the department to target doctor_id,
    migrates dates from the user's previous doctor card, soft-deletes an empty
    duplicate auto-created doctor card, and commits once.
    """
    doctor = _get_doctor_or_404(
        session=session,
        doctor_id=doctor_id,
        department_id=department_id,
    )
    user = user_repository.get_user_by_id(session, user_id)
    if (
        user is None
        or user.is_deleted
        or user.role != UserRole.DOCTOR
        or user.department_id != department_id
    ):
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Doctor user not found.",
        )

    if user.doctor_id == doctor.id:
        return _get_doctor_read_with_positions(session, doctor, department_id)

    if user_repository.get_user_by_doctor_id(session, doctor.id) is not None:
        raise HTTPException(
            status_code=status.HTTP_409_CONFLICT,
            detail="Doctor is already linked to another user account.",
        )

    old_doctor_id = user.doctor_id
    try:
        user_repository.reassign_user_doctor(session, user, doctor.id, commit=False)
        if old_doctor_id is not None and old_doctor_id != doctor.id:
            doctor_repository.migrate_doctor_schedule_dates(
                session=session,
                from_doctor_id=old_doctor_id,
                to_doctor_id=doctor.id,
                commit=False,
            )
            _cleanup_empty_duplicate_doctor(
                session=session,
                old_doctor_id=old_doctor_id,
                department_id=department_id,
            )
        session.commit()
    except Exception:
        session.rollback()
        raise

    session.refresh(doctor)
    return _get_doctor_read_with_positions(session, doctor, department_id)


def _cleanup_empty_duplicate_doctor(
    session: Session,
    old_doctor_id: uuid.UUID,
    department_id: uuid.UUID,
) -> None:
    """Soft-deletes an unlinked duplicate doctor card if it holds no positions or shifts."""
    old_doctor = doctor_repository.get_doctor_by_id_for_department(
        session=session,
        doctor_id=old_doctor_id,
        department_id=department_id,
    )
    if old_doctor is None:
        return
    if doctor_repository.has_active_shift_assignments(session, old_doctor_id):
        return
    if doctor_repository.get_doctor_positions_for_department(
        session=session,
        doctor_id=old_doctor_id,
        department_id=department_id,
    ):
        return
    doctor_repository.soft_delete_doctor(session, old_doctor, commit=False)
