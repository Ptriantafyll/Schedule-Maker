"""
Unit and constraint tests for the RefreshSession SQLModel.
"""

import uuid
import datetime
import pytest
from sqlalchemy.exc import IntegrityError
from sqlmodel import Session, select

from src.auth.models import RefreshSession
from src.user.models import UserRole


def test_create_refresh_session_defaults_and_properties(session: Session, user_factory):
    """A fresh, unrevoked, unexpired session is active."""
    user = user_factory(role=UserRole.SUPER_ADMIN)
    now = datetime.datetime.now(datetime.timezone.utc)
    token_hash = "hash_sample_" + uuid.uuid4().hex
    csrf_hash = "csrf_sample_" + uuid.uuid4().hex
    family_id = uuid.uuid4()

    refresh_session = RefreshSession(
        user_id=user.id,
        refresh_token_hash=token_hash,
        session_family=family_id,
        csrf_token_hash=csrf_hash,
        expires_at=now + datetime.timedelta(days=14),
    )

    session.add(refresh_session)
    session.commit()
    session.refresh(refresh_session)

    assert isinstance(refresh_session.id, uuid.UUID)
    assert refresh_session.user_id == user.id
    assert refresh_session.refresh_token_hash == token_hash
    assert refresh_session.session_family == family_id
    assert refresh_session.csrf_token_hash == csrf_hash
    assert refresh_session.last_used_at is None
    assert refresh_session.revoked_at is None
    assert refresh_session.revoked_reason is None
    assert refresh_session.replaced_by_session_id is None
    assert refresh_session.is_deleted is False
    assert refresh_session.is_active is True
    assert refresh_session.is_expired is False


def test_refresh_session_is_expired_when_past_timestamp(session: Session, user_factory):
    """A session whose expires_at timestamp is in the past is expired and inactive."""
    user = user_factory(role=UserRole.SUPER_ADMIN)
    now = datetime.datetime.now(datetime.timezone.utc)

    past_session = RefreshSession(
        user_id=user.id,
        refresh_token_hash="hash_past_" + uuid.uuid4().hex,
        session_family=uuid.uuid4(),
        expires_at=now - datetime.timedelta(hours=1),
    )

    assert past_session.is_expired is True
    assert past_session.is_active is False


def test_refresh_session_revocation_marks_inactive(session: Session, user_factory):
    """Revoking a session marks it inactive even if expires_at is in the future."""
    user = user_factory(role=UserRole.SUPER_ADMIN)
    now = datetime.datetime.now(datetime.timezone.utc)

    revoked_session = RefreshSession(
        user_id=user.id,
        refresh_token_hash="hash_revoked_" + uuid.uuid4().hex,
        session_family=uuid.uuid4(),
        expires_at=now + datetime.timedelta(days=14),
        revoked_at=now,
        revoked_reason="logout",
    )

    assert revoked_session.is_expired is False
    assert revoked_session.is_active is False
    assert revoked_session.revoked_reason == "logout"


def test_refresh_session_replaced_by_successor_marks_inactive(session: Session, user_factory):
    """When a session is rotated and replaced_by_session_id is set, it is inactive."""
    user = user_factory(role=UserRole.SUPER_ADMIN)
    now = datetime.datetime.now(datetime.timezone.utc)
    successor_id = uuid.uuid4()

    rotated_session = RefreshSession(
        user_id=user.id,
        refresh_token_hash="hash_rotated_" + uuid.uuid4().hex,
        session_family=uuid.uuid4(),
        expires_at=now + datetime.timedelta(days=14),
        replaced_by_session_id=successor_id,
        revoked_at=now,
        revoked_reason="rotated",
    )

    assert rotated_session.is_expired is False
    assert rotated_session.is_active is False
    assert rotated_session.replaced_by_session_id == successor_id
    assert rotated_session.revoked_reason == "rotated"


def test_refresh_session_soft_delete_marks_inactive(session: Session, user_factory):
    """A soft-deleted session (is_deleted=True) is inactive."""
    user = user_factory(role=UserRole.SUPER_ADMIN)
    now = datetime.datetime.now(datetime.timezone.utc)

    deleted_session = RefreshSession(
        user_id=user.id,
        refresh_token_hash="hash_deleted_" + uuid.uuid4().hex,
        session_family=uuid.uuid4(),
        expires_at=now + datetime.timedelta(days=14),
        is_deleted=True,
    )

    assert deleted_session.is_expired is False
    assert deleted_session.is_active is False


def test_refresh_session_token_hash_unique_constraint(session: Session, user_factory):
    """The database enforces a unique constraint on refresh_token_hash."""
    user = user_factory(role=UserRole.SUPER_ADMIN)
    now = datetime.datetime.now(datetime.timezone.utc)
    duplicate_hash = "identical_token_hash_12345"

    session_1 = RefreshSession(
        user_id=user.id,
        refresh_token_hash=duplicate_hash,
        session_family=uuid.uuid4(),
        expires_at=now + datetime.timedelta(days=14),
    )
    session.add(session_1)
    session.commit()

    session_2 = RefreshSession(
        user_id=user.id,
        refresh_token_hash=duplicate_hash,
        session_family=uuid.uuid4(),
        expires_at=now + datetime.timedelta(days=14),
    )
    session.add(session_2)

    with pytest.raises(IntegrityError):
        session.commit()
    session.rollback()


def test_refresh_session_family_grouping(session: Session, user_factory):
    """Multiple sessions can share a session_family for rotation tracking."""
    user = user_factory(role=UserRole.SUPER_ADMIN)
    now = datetime.datetime.now(datetime.timezone.utc)
    family_id = uuid.uuid4()

    s1 = RefreshSession(
        user_id=user.id,
        refresh_token_hash="hash_fam_1_" + uuid.uuid4().hex,
        session_family=family_id,
        expires_at=now + datetime.timedelta(days=14),
        revoked_at=now,
        revoked_reason="rotated",
    )
    session.add(s1)
    session.commit()

    s2 = RefreshSession(
        user_id=user.id,
        refresh_token_hash="hash_fam_2_" + uuid.uuid4().hex,
        session_family=family_id,
        expires_at=now + datetime.timedelta(days=14),
    )
    session.add(s2)
    session.commit()

    family_sessions = session.exec(
        select(RefreshSession).where(RefreshSession.session_family == family_id)
    ).all()

    assert len(family_sessions) == 2
    session_ids = {s.id for s in family_sessions}
    assert s1.id in session_ids
    assert s2.id in session_ids


def test_refresh_session_factory_defaults(refresh_session_factory):
    """The refresh_session_factory fixture produces a valid, persisted active session."""
    session_obj = refresh_session_factory()

    assert isinstance(session_obj.id, uuid.UUID)
    assert isinstance(session_obj.user_id, uuid.UUID)
    assert len(session_obj.refresh_token_hash) == 64
    assert isinstance(session_obj.session_family, uuid.UUID)
    assert session_obj.csrf_token_hash is None
    assert session_obj.is_active is True
    assert session_obj.is_expired is False

