"""
Doctor repository function for handling database operations.
"""

import uuid
from sqlmodel import Session, not_, select
from src.doctor.schemas import DoctorCreate, DoctorPreAssignmentCreate, DoctorUnavailabilityCreate, DoctorPositionCreate
from src.doctor.models import Doctor as DoctorModel
# from src.doctor.models import DoctorPreAssignment as DoctorPreAssignmentModel
# from src.doctor.models import DoctorUnavailability as DoctorUnavailabilityModel
# from src.doctor.models import DoctorPosition as DoctorPositionModel


def get_doctor_by_email(session: Session, email: str) -> DoctorModel:
    """Retrieves a doctor by their unique email"""
    return session.exec(
        select(DoctorModel).where(DoctorModel.email == email)
    ).first()

# def get_doctor_by_id(session: Session, doctor_id: str) -> DoctorModel:
#     pass


# def get_active_doctors(session: Session) -> list[DoctorModel]:
#     pass


def create_doctor(session: Session, doctor_data: DoctorCreate) -> DoctorModel:
    """Creates a new doctor in the database"""
    new_doctor = DoctorModel(
        name=doctor_data.name,
        email=doctor_data.email,
        department_id=doctor_data.department_id,
        team_id=doctor_data.team_id
    )
    session.add(new_doctor)
    session.commit()
    session.refresh(new_doctor)
    return new_doctor


# def create_doctor_pre_assignments(session: Session, doctor_id: uuid.UUID, pre_assignment_data: DoctorPreAssignmentCreate) -> DoctorPreAssignmentModel:
#     pass


# def get_doctor_pre_assignments(session: Session, doctor_id: str) -> list[DoctorPreAssignmentModel]:
#     pass


# def create_doctor_unavailability(session: Session, doctor_unavailability_data: DoctorUnavailabilityCreate) -> DoctorUnavailabilityModel:
#     pass


# def get_doctor_unavailability(session: Session, doctor_id: str) -> list[DoctorUnavailabilityModel]:
#     pass


# def create_doctor_position(session: Session, doctor_pos_data: DoctorPositionCreate) -> DoctorPositionModel:
#     pass


# def get_doctor_position(session: Session, doctor_id: str) -> DoctorPositionModel:
#     pass
