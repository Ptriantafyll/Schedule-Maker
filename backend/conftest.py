"""
Configuration for tests
"""

import uuid
import datetime
import pytest
from sqlmodel import SQLModel, create_engine, Session
from sqlalchemy.pool import StaticPool
from fastapi.testclient import TestClient

from src.department.schemas import DepartmentCreate
from src.department import repository as department_repository
from src.doctor import repository as doctor_repository
from src.team import repository as team_repository
from src.db.connection import get_session
from src.main import app
from src.user.models import UserRole
from src.user.models import User as UserModel
from src.department.models import Department as DepartmentModel
from src.doctor.models import Doctor as DoctorModel
from src.team.models import Team as TeamModel
from src.auth.models import Invitation as InvitationModel, RefreshSession as RefreshSessionModel
from src.user.schemas import UserAccountCreate
from src.user import services as user_services
from src.auth.security import create_access_token
from src.utils.rate_limiter import auth_rate_limiter


@pytest.fixture(autouse=True)
def reset_rate_limiter():
    """Resets the in-memory rate limiter to ensure complete test isolation."""
    auth_rate_limiter.reset()
    yield
    auth_rate_limiter.reset()


@pytest.fixture(name="session")
def session_fixture():
    """Creates a fresh in-memory database session for each test."""
    engine = create_engine(
        "sqlite:///:memory:",
        connect_args={"check_same_thread": False},
        poolclass=StaticPool,
    )
    SQLModel.metadata.create_all(engine)
    with Session(engine) as session:
        yield session


@pytest.fixture(name="client")
def client_fixture(session, monkeypatch):
    """Creates a TestClient for the FastAPI app with dependency override."""
    try:
        def override_get_session():
            yield session

        app.dependency_overrides[get_session] = override_get_session
        monkeypatch.setattr("src.main.init_db", lambda: None)

        with TestClient(app) as test_client:
            yield test_client
    finally:
        app.dependency_overrides.clear()


@pytest.fixture(name="department_factory")
def department_factory_fixture(session):
    """Create persisted department with customizable name and code"""
    def create_department(
        *,
        name: str | None = None,
        code: str | None = None,
    ) -> DepartmentModel:
        suffix = uuid.uuid4().hex[:8]

        return department_repository.create_department(
            session=session,
            department_data=DepartmentCreate(
                name=name or f"Department {suffix}",
                code=code or suffix.upper(),
            ),
        )

    return create_department


@pytest.fixture(name="team_factory")
def team_factory_fixture(
    session,
    department_factory
):
    """Create persisted team with customizable department field"""
    department = department_factory(
        name="Test Department",
        code="TEST",
    )

    def create_team(
        *,
        department_id: uuid.UUID = department.id,
        name: str | None = None,
    ) -> TeamModel:
        suffix = uuid.uuid4().hex[:8]

        return team_repository.create_team(
            session=session,
            name=name or f"Test team {suffix}",
            department_id=department_id,
        )

    return create_team


@pytest.fixture(name="doctor_factory")
def doctor_factory_fixture(
    session,
    team_factory
):
    """Create persisted doctor with customizable fields"""
    def create_doctor(
        *,
        team_id: uuid.UUID | None = None,
        department_id: uuid.UUID | None = None,
        full_name: str | None = None,
        email: str | None = None,
    ) -> DoctorModel:
        if department_id is None and team_id is None:
            team = team_factory()
            team_id = team.id
            department_id = team.department_id
        elif department_id is None and team_id is not None:
            team = session.get(TeamModel, team_id)
            department_id = team.department_id
        elif department_id is not None and team_id is None:
            team = team_factory(department_id=department_id)
            team_id = team.id

        suffix = uuid.uuid4().hex[:8]

        return doctor_repository.create_doctor(
            session=session,
            name=full_name or f"Doctor {suffix}",
            team_id=team_id,
            department_id=department_id,
        )

    return create_doctor


@pytest.fixture(name="user_factory")
def user_factory_fixture(session):
    """Create persisted users with customizable roles and relationships"""

    def create_user(
        *,
        role: UserRole,
        department_id: uuid.UUID | None = None,
        doctor_id: uuid.UUID | None = None,
        email: str | None = None,
        full_name: str = "Test User",
        password: str = "test-password",
    ) -> UserModel:
        user_email = email or f"user-{uuid.uuid4().hex}@test.com"

        account_data = UserAccountCreate(
            email=user_email,
            full_name=full_name,
            password=password,
            role=role,
            department_id=department_id,
            doctor_id=doctor_id,
        )

        return user_services.create_user_account(
            session=session,
            account_data=account_data,
        )

    return create_user


@pytest.fixture(name="invitation_factory")
def invitation_factory_fixture(session, department_factory, user_factory):
    """Create persisted invitations with customizable fields and defaults"""

    def create_invitation(
        *,
        role: UserRole = UserRole.DOCTOR,
        department_id: uuid.UUID | None = None,
        created_by_user_id: uuid.UUID | None = None,
        doctor_id: uuid.UUID | None = None,
        token_hash: str | None = None,
        expires_at: datetime.datetime | None = None,
        used_at: datetime.datetime | None = None,
        revoked_at: datetime.datetime | None = None,
        is_deleted: bool = False,
    ) -> InvitationModel:
        if department_id is None:
            department_id = department_factory().id
        if created_by_user_id is None:
            admin = user_factory(
                role=UserRole.DEPARTMENT_ADMIN,
                department_id=department_id,
            )
            created_by_user_id = admin.id
        if expires_at is None:
            expires_at = datetime.datetime.now(
                datetime.timezone.utc) + datetime.timedelta(days=7)
        if token_hash is None:
            token_hash = uuid.uuid4().hex * 2  # 64 hex characters

        invitation = InvitationModel(
            token_hash=token_hash,
            role=role,
            department_id=department_id,
            doctor_id=doctor_id,
            created_by_user_id=created_by_user_id,
            expires_at=expires_at,
            used_at=used_at,
            revoked_at=revoked_at,
            is_deleted=is_deleted,
        )
        session.add(invitation)
        session.commit()
        session.refresh(invitation)
        return invitation

    return create_invitation


@pytest.fixture(name="refresh_session_factory")
def refresh_session_factory_fixture(session, user_factory):
    """Create persisted refresh sessions with customizable fields and defaults."""

    def create_refresh_session(  # pylint: disable=too-many-arguments
        *,
        user_id: uuid.UUID | None = None,
        refresh_token_hash: str | None = None,
        session_family: uuid.UUID | None = None,
        csrf_token_hash: str | None = None,
        expires_at: datetime.datetime | None = None,
        last_used_at: datetime.datetime | None = None,
        revoked_at: datetime.datetime | None = None,
        revoked_reason: str | None = None,
        replaced_by_session_id: uuid.UUID | None = None,
        is_deleted: bool = False,
    ) -> RefreshSessionModel:
        if user_id is None:
            user = user_factory(role=UserRole.SUPER_ADMIN)
            user_id = user.id
        if refresh_token_hash is None:
            refresh_token_hash = uuid.uuid4().hex * 2  # 64 hex characters
        if session_family is None:
            session_family = uuid.uuid4()
        if expires_at is None:
            expires_at = datetime.datetime.now(
                datetime.timezone.utc) + datetime.timedelta(days=14)

        refresh_session = RefreshSessionModel(
            user_id=user_id,
            refresh_token_hash=refresh_token_hash,
            session_family=session_family,
            csrf_token_hash=csrf_token_hash,
            expires_at=expires_at,
            last_used_at=last_used_at,
            revoked_at=revoked_at,
            revoked_reason=revoked_reason,
            replaced_by_session_id=replaced_by_session_id,
            is_deleted=is_deleted,
        )
        session.add(refresh_session)
        session.commit()
        session.refresh(refresh_session)
        return refresh_session

    return create_refresh_session


@pytest.fixture(name="auth_headers_factory")
def auth_headers_factory_fixture():
    """Create auth headers factory fixture"""

    def create_auth_headers(user: UserModel) -> dict[str, str]:
        access_token = create_access_token({"sub": str(user.id)})

        return {"Authorization": f"Bearer {access_token}"}

    return create_auth_headers
