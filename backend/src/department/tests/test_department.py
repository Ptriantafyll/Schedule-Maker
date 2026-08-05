"""
Tests for the department module
"""

from fastapi.testclient import TestClient
import uuid
import datetime
from src.main import app
from src.department.repository import create_department, get_active_departments, get_department_by_name
from src.department.schemas import DepartmentCreate

# Session fixture for database tests
import pytest
from sqlalchemy.pool import StaticPool
from sqlmodel import SQLModel, create_engine, Session


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


#####################
# Repository tests
#####################
def test_create_department(session):
    """Test creating a department and verifying its fields."""

    dept_data = DepartmentCreate(name="Cardiology", code="CARD")
    new_dept = create_department(session, dept_data)

    assert isinstance(new_dept.id, uuid.UUID)
    assert new_dept.name == "Cardiology"
    assert new_dept.code == "CARD"
    assert new_dept.is_deleted is False
    assert new_dept.sync_status is False
    assert isinstance(new_dept.created_at, datetime.datetime)
    assert isinstance(new_dept.updated_at, datetime.datetime)


def test_get_department_by_name(session):
    """Test retrieving a department by its name."""
    dept_data = DepartmentCreate(name="Neurology", code="NEURO")
    new_dept = create_department(session, dept_data)

    retrieved_dept = get_department_by_name(session, "Neurology")

    assert retrieved_dept == new_dept


def test_get_active_departments(session):
    """Test retrieving only active (non-deleted) departments."""
    dept1 = create_department(
        session, DepartmentCreate(name="Oncology", code="ONC"))
    dept2 = create_department(
        session, DepartmentCreate(name="Pediatrics", code="PED"))

    # Mark one department as deleted
    dept2.is_deleted = True
    session.add(dept2)
    session.commit()

    active_departments = get_active_departments(session)
    assert dept1 in active_departments
    assert dept2 not in active_departments

#######################
# Controller tests
#######################


def test_create_department_controller_duplicate_name(session):
    """Test that creating a department with a duplicate name raises an error."""
    dept_data = DepartmentCreate(name="Radiology", code="RAD")
    create_department(session, dept_data)

    from src.department.controllers import create_department_controller
    from fastapi import HTTPException

    with pytest.raises(HTTPException) as exc_info:
        create_department_controller(dept_data, session)

    assert exc_info.value.status_code == 400
    assert "already exists" in exc_info.value.detail


def test_get_department_controller_nonexistent(session):
    """Test that fetching a non-existent department raises a 404 error."""
    from src.department.controllers import get_department_controller
    from fastapi import HTTPException

    non_existent_id = uuid.uuid4()

    with pytest.raises(HTTPException) as exc_info:
        get_department_controller(non_existent_id, session)

    assert exc_info.value.status_code == 404
    assert "not found" in exc_info.value.detail


def test_get_department_controller_deleted(session):
    """Test that fetching a deleted department raises a 404 error."""
    dept_data = DepartmentCreate(name="Gastroenterology", code="GASTRO")
    new_dept = create_department(session, dept_data)

    # Mark the department as deleted
    new_dept.is_deleted = True
    session.add(new_dept)
    session.commit()

    from src.department.controllers import get_department_controller
    from fastapi import HTTPException

    with pytest.raises(HTTPException) as exc_info:
        get_department_controller(new_dept.id, session)

    assert exc_info.value.status_code == 404
    assert "not found" in exc_info.value.detail


###############################
# Route tests
###############################
@pytest.fixture(name="client")
def client_fixture(session):
    """Creates a TestClient for the FastAPI app with dependency override."""
    from src.db.connection import get_session

    def override_get_session():
        yield session
    app.dependency_overrides[get_session] = override_get_session
    with TestClient(app) as test_client:
        yield test_client
    # 5. Clean up the overrides after the test finishes
    app.dependency_overrides.clear()


def test_create_department_route(client):
    """Test the POST /departments/ route for creating a department."""
    response = client.post(
        "api/v1/departments/", json={"name": "Urology", "code": "URO"})
    assert response.status_code == 201
    data = response.json()
    assert data["name"] == "Urology"
    assert data["code"] == "URO"
    assert "id" in data
    assert "created_at" in data
    assert "updated_at" in data


def test_create_department_route_invalid_payload(client):
    """Test that the POST /departments/ route rejects invalid payloads."""
    response = client.post(
        "api/v1/departments/", json={"code": "URO"})  # Missing 'name'
    assert response.status_code == 422


def test_get_department_by_id_route(session, client):
    """Tests that the GET /departments/{department_id} route returns a department"""

    dept_data = DepartmentCreate(name="Radiology", code="RAD")
    new_dept = create_department(session, dept_data)

    response = client.get(
        f"/api/v1/departments/{new_dept.id}"
    )

    assert response.status_code == 200
    data = response.json()
    assert data["name"] == "Radiology"
    assert data["code"] == "RAD"
    assert "created_at" in data
    assert "updated_at" in data


def test_list_departments_route():
    """Tests the GET /departments/ route"""
    # TODO
