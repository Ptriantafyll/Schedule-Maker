"""
Tests for the department module
"""

import datetime
import uuid
import pytest
from fastapi import HTTPException

from src.department.schemas import DepartmentCreate
from src.department import repository as department_repository
from src.department import controllers as department_controllers
from src.user.models import UserRole


#####################
# Fixtures
#####################
@pytest.fixture(name="department")
def department_fixture(session):
    """Creates a reusable department for tests"""
    dept_data = DepartmentCreate(name="Cardiology", code="CARD")
    return department_repository.create_department(session, dept_data)


@pytest.fixture(name="department_admin_user")
def department_admin_user_fixture(user_factory, department):
    """Creates a reusable department admin user for tests"""
    return user_factory(
        role=UserRole.DEPARTMENT_ADMIN,
        department_id=department.id,
        doctor_id=None
    )


@pytest.fixture(name="department_admin_headers")
def department_admin_headers_fixture(department_admin_user, auth_headers_factory):
    """Creates reusable department admin auth headers for tests"""

    return auth_headers_factory(department_admin_user)


@pytest.fixture(name="super_admin_user")
def super_admin_user_fixture(user_factory):
    """Creates a reusable super admin user for tests"""
    return user_factory(
        role=UserRole.SUPER_ADMIN,
        department_id=None,
        doctor_id=None
    )


@pytest.fixture(name="super_admin_headers")
def super_admin_headers_fixture(super_admin_user, auth_headers_factory):
    """Creates reusable super admin auth headers for tests"""

    return auth_headers_factory(super_admin_user)


#####################
# Repository tests
#####################


def test_create_department(session):
    """Test creating a department and verifying its fields."""

    dept_data = DepartmentCreate(name="Cardiology", code="CARD")
    new_dept = department_repository.create_department(session, dept_data)

    assert isinstance(new_dept.id, uuid.UUID)
    assert new_dept.name == "Cardiology"
    assert new_dept.code == "CARD"
    assert new_dept.is_deleted is False
    assert new_dept.sync_status is False
    assert isinstance(new_dept.created_at, datetime.datetime)
    assert isinstance(new_dept.updated_at, datetime.datetime)


def test_get_department_by_name(session, department):
    """Test retrieving a department by its name."""
    retrieved_dept = department_repository.get_department_by_name(
        session, department.name)

    assert retrieved_dept == department


def test_get_active_departments(session):
    """Test retrieving only active (non-deleted) departments."""
    dept1 = department_repository.create_department(
        session, DepartmentCreate(name="Oncology", code="ONC"))
    dept2 = department_repository.create_department(
        session, DepartmentCreate(name="Pediatrics", code="PED"))

    # Mark one department as deleted
    dept2.is_deleted = True
    session.add(dept2)
    session.commit()

    active_departments = department_repository.get_active_departments(session)
    assert dept1 in active_departments
    assert dept2 not in active_departments

#######################
# Controller tests
#######################


def test_create_department_controller_duplicate_name(session):
    """Test that creating a department with a duplicate name raises an error."""
    dept_data = DepartmentCreate(name="Radiology", code="RAD")
    department_repository.create_department(session, dept_data)

    with pytest.raises(HTTPException) as exc_info:
        department_controllers.create_department_controller(dept_data, session)

    assert exc_info.value.status_code == 400
    assert "already exists" in exc_info.value.detail


def test_get_department_controller_nonexistent(session):
    """Test that fetching a non-existent department raises a 404 error."""

    non_existent_id = uuid.uuid4()

    with pytest.raises(HTTPException) as exc_info:
        department_controllers.get_department_controller(
            non_existent_id, session)

    assert exc_info.value.status_code == 404
    assert "not found" in exc_info.value.detail


def test_get_department_controller_deleted(session, department):
    """Test that fetching a deleted department raises a 404 error."""
    # Mark the department as deleted
    department.is_deleted = True
    session.add(department)
    session.commit()

    with pytest.raises(HTTPException) as exc_info:
        department_controllers.get_department_controller(
            department.id, session)

    assert exc_info.value.status_code == 404
    assert "not found" in exc_info.value.detail


###############################
# Route tests
###############################

def test_get_department_by_id_route(client, department, department_admin_headers):
    """Tests that the GET /departments/{department_id} route returns a department"""

    response = client.get(
        f"/api/v1/departments/{department.id}",
        headers=department_admin_headers,
    )

    assert response.status_code == 200
    data = response.json()
    assert data["name"] == department.name
    assert data["code"] == department.code
    assert "created_at" in data
    assert "updated_at" in data


def test_get_department_by_id_route_requires_authentication(client, department):
    """Tests that the GET /departments/{department_id} requires authentication"""
    response = client.get(
        f"/api/v1/departments/{department.id}",
    )

    assert response.status_code == 401
    assert response.json() == {
        "detail": "Unauthorized"
    }
    assert response.headers.get("WWW-Authenticate") == "Bearer"


def test_super_admin_can_list_departments(client, super_admin_headers, department):
    """Tests the GET /departments/ route"""

    response = client.get(
        "/api/v1/departments/",
        headers=super_admin_headers,
    )

    assert response.status_code == 200
    data = response.json()
    assert isinstance(data, list)

    returned_department = next(
        (
            item
            for item in data
            if item["id"] == str(department.id)
        ),
        None
    )

    assert returned_department is not None
    assert returned_department["code"] == department.code
    assert returned_department["name"] == department.name


@pytest.mark.parametrize(
    "role",
    [
        UserRole.DEPARTMENT_ADMIN,
        UserRole.DOCTOR,
        UserRole.VIEWER,
    ]
)
def test_non_super_admin_cannot_list_departments(
    client,
    department,
    user_factory,
    auth_headers_factory,
    role,
):
    """Tests that POST /api/v1/departments/ route rejects other roles except SUPER_ADMIN"""
    user = user_factory(
        role=role,
        department_id=department.id
    )
    headers = auth_headers_factory(user)

    response = client.get(
        "/api/v1/departments",
        headers=headers,
    )

    assert response.status_code == 402
    assert response.json() == {
        "detail": "Unauthorized"
    }
    assert response.headers.get("WWW-Authenticate") == "Bearer"
