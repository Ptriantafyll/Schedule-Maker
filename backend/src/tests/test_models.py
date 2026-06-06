import uuid
import datetime
from sqlmodel import SQLModel, create_engine, Session
from db_models import Department


def test_sync_base_fields():
    """Verify that models inheriting from SyncBase automatically get UUID and sync metadata."""
    engine = create_engine("sqlite:///:memory:")
    SQLModel.metadata.create_all(engine)

    with Session(engine) as session:
        # Create a test department
        dept = Department(name="Cardiology", code="CARD")
        session.add(dept)
        session.commit()
        session.refresh(dept)

        # Verify SyncBase properties
        assert isinstance(dept.id, uuid.UUID)
        assert dept.is_deleted is False
        assert dept.sync_status is False
        assert isinstance(dept.created_at, datetime.datetime)
        assert isinstance(dept.updated_at, datetime.datetime)
