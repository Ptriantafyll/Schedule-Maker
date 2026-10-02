"""Unit tests for the ScheduleDraft SQLModel and schemas."""

import uuid
import pytest
from src.schedule.models import ScheduleDraft
from src.schedule.schemas import ScheduleDraftRead, ScheduleAssignmentItem


def test_schedule_draft_default_values():
    """Verify that ScheduleDraft instantiates with correct default values and SyncBase properties."""
    dept_id = uuid.uuid4()
    draft = ScheduleDraft(
        department_id=dept_id,
        target_month="2026-11",
        source_filename="november_roster.xlsx",
    )

    assert isinstance(draft.id, uuid.UUID)
    assert draft.department_id == dept_id
    assert draft.target_month == "2026-11"
    assert draft.source_filename == "november_roster.xlsx"
    assert draft.total_duties == 0
    assert draft.solver_status == "OPTIMAL"
    assert draft.status == "draft"
    assert draft.assignments == []
    assert draft.is_deleted is False
    assert draft.sync_status is False


def test_schedule_draft_with_assignments():
    """Verify that ScheduleDraft holds assignment dictionaries in its JSON column."""
    dept_id = uuid.uuid4()
    assignments = [
        {
            "date": "2026-11-01",
            "day_name": "Sunday",
            "doctor_name": "Dr. Gregory House",
            "doctor_email": "house@hospital.org",
            "position": "ER",
            "shift": "Night",
        },
        {
            "date": "2026-11-02",
            "day_name": "Monday",
            "doctor_name": "Dr. Allison Cameron",
            "doctor_email": "cameron@hospital.org",
            "position": "ER",
            "shift": "Night",
        },
    ]

    draft = ScheduleDraft(
        department_id=dept_id,
        target_month="2026-11",
        source_filename="november_roster.xlsx",
        total_duties=2,
        solver_status="OPTIMAL",
        status="draft",
        assignments=assignments,
    )

    assert draft.total_duties == 2
    assert len(draft.assignments) == 2
    assert draft.assignments[0]["doctor_name"] == "Dr. Gregory House"
    assert draft.assignments[1]["doctor_name"] == "Dr. Allison Cameron"


def test_schedule_assignment_item_schema_validation():
    """Verify ScheduleAssignmentItem validates assignment dictionaries."""
    data = {
        "date": "2026-11-01",
        "day_name": "Sunday",
        "doctor_name": "Dr. Gregory House",
        "doctor_email": "house@hospital.org",
        "position": "ER",
        "shift": "Night",
    }
    item = ScheduleAssignmentItem.model_validate(data)
    assert item.date == "2026-11-01"
    assert item.day_name == "Sunday"
    assert item.doctor_name == "Dr. Gregory House"
    assert item.doctor_email == "house@hospital.org"
    assert item.position == "ER"
    assert item.shift == "Night"


def test_schedule_draft_read_schema_serialization():
    """Verify ScheduleDraftRead deserializes an ORM ScheduleDraft instance."""
    dept_id = uuid.uuid4()
    draft = ScheduleDraft(
        department_id=dept_id,
        target_month="2026-11",
        source_filename="november_roster.xlsx",
        total_duties=1,
        solver_status="OPTIMAL",
        status="draft",
        assignments=[
            {
                "date": "2026-11-01",
                "day_name": "Sunday",
                "doctor_name": "Dr. Gregory House",
                "doctor_email": "house@hospital.org",
                "position": "ER",
                "shift": "Night",
            }
        ],
    )

    dto = ScheduleDraftRead.model_validate(draft)
    assert dto.id == draft.id
    assert dto.department_id == dept_id
    assert dto.target_month == "2026-11"
    assert dto.source_filename == "november_roster.xlsx"
    assert dto.total_duties == 1
    assert dto.solver_status == "OPTIMAL"
    assert dto.status == "draft"
    assert len(dto.assignments) == 1
    assert dto.assignments[0].doctor_name == "Dr. Gregory House"
