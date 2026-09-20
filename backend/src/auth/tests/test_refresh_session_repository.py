"""
Unit tests for the RefreshSession repository layer.
"""

import uuid
import datetime
from sqlmodel import Session

from src.auth import repository as auth_repository
from src.user.models import UserRole


def test_create_refresh_session_persists(session: Session, user_factory):
    """create_refresh_session persists a new session and auto-generates family."""
    user = user_factory(role=UserRole.SUPER_ADMIN)
    token_hash = "hash_" + uuid.uuid4().hex
    expires_at = datetime.datetime.now(datetime.timezone.utc) + datetime.timedelta(days=14)

    new_session = auth_repository.create_refresh_session(
        session=session,
        user_id=user.id,
        refresh_token_hash=token_hash,
        expires_at=expires_at,
    )

    assert isinstance(new_session.id, uuid.UUID)
    assert new_session.user_id == user.id
    assert new_session.refresh_token_hash == token_hash
    assert isinstance(new_session.session_family, uuid.UUID)
    assert new_session.csrf_token_hash is None
    assert new_session.is_active is True
    assert new_session.replaced_by_session_id is None


def test_create_refresh_session_with_explicit_family(session: Session, user_factory):
    """create_refresh_session respects an explicitly provided session_family."""
    user = user_factory(role=UserRole.SUPER_ADMIN)
    family_id = uuid.uuid4()
    token_hash = "hash_" + uuid.uuid4().hex
    csrf_hash = "csrf_" + uuid.uuid4().hex
    expires_at = datetime.datetime.now(datetime.timezone.utc) + datetime.timedelta(days=14)

    new_session = auth_repository.create_refresh_session(
        session=session,
        user_id=user.id,
        refresh_token_hash=token_hash,
        expires_at=expires_at,
        session_family=family_id,
        csrf_token_hash=csrf_hash,
    )

    assert new_session.session_family == family_id
    assert new_session.csrf_token_hash == csrf_hash


def test_get_refresh_session_by_token_hash_found(session: Session, refresh_session_factory):
    """get_refresh_session_by_token_hash finds a persisted session."""
    target_hash = "target_hash_" + uuid.uuid4().hex
    persisted = refresh_session_factory(refresh_token_hash=target_hash)

    found = auth_repository.get_refresh_session_by_token_hash(
        session=session,
        token_hash=target_hash,
    )

    assert found is not None
    assert found.id == persisted.id
    assert found.refresh_token_hash == target_hash


def test_get_refresh_session_by_token_hash_not_found(session: Session):
    """get_refresh_session_by_token_hash returns None for unknown hashes."""
    found = auth_repository.get_refresh_session_by_token_hash(
        session=session,
        token_hash="non_existent_hash",
    )
    assert found is None


def test_get_refresh_session_by_token_hash_ignores_deleted(session: Session, refresh_session_factory):
    """Soft-deleted sessions are ignored when querying by token hash."""
    target_hash = "deleted_hash_" + uuid.uuid4().hex
    refresh_session_factory(refresh_token_hash=target_hash, is_deleted=True)

    found = auth_repository.get_refresh_session_by_token_hash(
        session=session,
        token_hash=target_hash,
    )
    assert found is None


def test_rotate_refresh_session_atomic(session: Session, refresh_session_factory):
    """Rotating a session supersedes the old session and links the new one in the same family."""
    old_session = refresh_session_factory()
    original_family = old_session.session_family
    new_token_hash = "new_hash_" + uuid.uuid4().hex
    new_csrf_hash = "new_csrf_" + uuid.uuid4().hex
    new_expires_at = datetime.datetime.now(datetime.timezone.utc) + datetime.timedelta(days=14)

    successor = auth_repository.rotate_refresh_session(
        session=session,
        current_session=old_session,
        new_refresh_token_hash=new_token_hash,
        new_expires_at=new_expires_at,
        new_csrf_token_hash=new_csrf_hash,
    )

    session.refresh(old_session)

    # Successor assertions
    assert successor.id != old_session.id
    assert successor.user_id == old_session.user_id
    assert successor.session_family == original_family
    assert successor.refresh_token_hash == new_token_hash
    assert successor.csrf_token_hash == new_csrf_hash
    assert successor.is_active is True
    assert successor.replaced_by_session_id is None

    # Predecessor assertions
    assert old_session.replaced_by_session_id == successor.id
    assert old_session.revoked_reason == "rotated"
    assert old_session.revoked_at is not None
    assert old_session.last_used_at is not None
    assert old_session.is_active is False


def test_revoke_refresh_session(session: Session, refresh_session_factory):
    """Revoking a single session sets revoked_at and revoked_reason."""
    active_session = refresh_session_factory()

    revoked = auth_repository.revoke_refresh_session(
        session=session,
        refresh_session=active_session,
        reason="logout",
    )

    assert revoked.id == active_session.id
    assert revoked.revoked_at is not None
    assert revoked.revoked_reason == "logout"
    assert revoked.is_active is False


def test_revoke_session_family(session: Session, user_factory, refresh_session_factory):
    """Revoking a family revokes only the active sessions within that family lineage."""
    user = user_factory(role=UserRole.SUPER_ADMIN)
    family_a = uuid.uuid4()
    family_b = uuid.uuid4()

    s1_a = refresh_session_factory(user_id=user.id, session_family=family_a)
    s2_a = refresh_session_factory(user_id=user.id, session_family=family_a)
    already_revoked_a = refresh_session_factory(
        user_id=user.id,
        session_family=family_a,
        revoked_at=datetime.datetime.now(datetime.timezone.utc),
        revoked_reason="rotated",
    )
    s_b = refresh_session_factory(user_id=user.id, session_family=family_b)

    revoked_count = auth_repository.revoke_session_family(
        session=session,
        session_family=family_a,
        reason="reuse_detected",
    )

    assert revoked_count == 2

    session.refresh(s1_a)
    session.refresh(s2_a)
    session.refresh(already_revoked_a)
    session.refresh(s_b)

    assert s1_a.revoked_reason == "reuse_detected"
    assert s1_a.is_active is False

    assert s2_a.revoked_reason == "reuse_detected"
    assert s2_a.is_active is False

    # Already revoked maintains its original rotation reason
    assert already_revoked_a.revoked_reason == "rotated"

    # Family B untouched
    assert s_b.is_active is True
    assert s_b.revoked_at is None


def test_revoke_all_sessions_for_user(session: Session, user_factory, refresh_session_factory):
    """Revoking all sessions for a user affects only that user's active sessions."""
    user_1 = user_factory(role=UserRole.SUPER_ADMIN)
    user_2 = user_factory(role=UserRole.SUPER_ADMIN)

    u1_s1 = refresh_session_factory(user_id=user_1.id)
    u1_s2 = refresh_session_factory(user_id=user_1.id)
    u2_s = refresh_session_factory(user_id=user_2.id)

    count = auth_repository.revoke_all_sessions_for_user(
        session=session,
        user_id=user_1.id,
        reason="user_locked",
    )

    assert count == 2

    session.refresh(u1_s1)
    session.refresh(u1_s2)
    session.refresh(u2_s)

    assert u1_s1.revoked_reason == "user_locked"
    assert u1_s1.is_active is False

    assert u1_s2.revoked_reason == "user_locked"
    assert u1_s2.is_active is False

    assert u2_s.is_active is True
    assert u2_s.revoked_at is None


def test_list_active_sessions_for_user(session: Session, user_factory, refresh_session_factory):
    """list_active_sessions_for_user returns only non-revoked, unexpired, non-replaced sessions."""
    user_1 = user_factory(role=UserRole.SUPER_ADMIN)
    user_2 = user_factory(role=UserRole.SUPER_ADMIN)
    now = datetime.datetime.now(datetime.timezone.utc)

    # Active session
    active_session = refresh_session_factory(user_id=user_1.id)

    # Revoked session
    refresh_session_factory(
        user_id=user_1.id,
        revoked_at=now,
        revoked_reason="logout",
    )

    # Replaced session
    refresh_session_factory(
        user_id=user_1.id,
        replaced_by_session_id=active_session.id,
        revoked_at=now,
        revoked_reason="rotated",
    )

    # Expired session
    refresh_session_factory(
        user_id=user_1.id,
        expires_at=now - datetime.timedelta(days=1),
    )

    # Soft-deleted session
    refresh_session_factory(
        user_id=user_1.id,
        is_deleted=True,
    )

    # Other user's active session
    refresh_session_factory(user_id=user_2.id)

    user_1_sessions = auth_repository.list_active_sessions_for_user(
        session=session,
        user_id=user_1.id,
    )

    assert len(user_1_sessions) == 1
    assert user_1_sessions[0].id == active_session.id
