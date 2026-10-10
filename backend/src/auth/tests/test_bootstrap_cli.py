"""
Tests for the bootstrap script
"""
from unittest.mock import Mock
from contextlib import nullcontext
import getpass
import pytest

from scripts import bootstrap_super_admin
from src.auth import bootstrap
from src.auth.security import verify_password

from src.user import repository as user_repository
from src.user.models import UserRole


CLI_EMAIL = "superadmin@test.com"
CLI_FULL_NAME = "Test Super Admin"
MATCHING_PASSWORD = "SecurePassword123!"
CLI_ARGS = (
    "--email",
    CLI_EMAIL,
    "--full-name",
    CLI_FULL_NAME
)


def _mock_password_prompts(monkeypatch, *passwords):
    """Mock password helper"""
    password_reader = iter(passwords)
    monkeypatch.setattr(
        getpass,
        "getpass",
        lambda _prompt: next(password_reader)
    )


@pytest.fixture(name="cli_database_mocks")
def cli_database_mocks_fixture(monkeypatch):
    """Creates a reusable cli database for tests"""
    init_db_mock = Mock()
    fake_session = object()
    session_factory_mock = Mock(
        return_value=nullcontext(fake_session)
    )

    monkeypatch.setattr(
        bootstrap_super_admin,
        "init_db",
        init_db_mock
    )

    monkeypatch.setattr(
        bootstrap_super_admin,
        "Session",
        session_factory_mock
    )

    return init_db_mock, session_factory_mock, fake_session


@pytest.fixture(name="cli_in_memory_database")
def cli_in_memory_database_fixture(monkeypatch, session):
    """Creates an in memory cli database for tests"""
    init_db_mock = Mock()
    session_factory_mock = Mock(
        return_value=nullcontext(session)
    )

    monkeypatch.setattr(
        bootstrap_super_admin,
        "init_db",
        init_db_mock
    )

    monkeypatch.setattr(
        bootstrap_super_admin,
        "Session",
        session_factory_mock
    )

    return init_db_mock, session_factory_mock


@pytest.fixture(name="create_super_admin_mock")
def create_super_admin_mock_fixture(monkeypatch):
    """Creates a reusable create super admin mock"""
    service_mock = Mock()

    monkeypatch.setattr(
        bootstrap,
        "create_super_admin",
        service_mock
    )

    return service_mock


def test_bootstrap_cli_rejects_mismatched_passwords(
    monkeypatch,
    capsys,
    create_super_admin_mock,
    cli_database_mocks,
):
    """Tests the bootstrap cli with mismatched passwords"""
    _mock_password_prompts(
        monkeypatch,
        "first-password",
        "second-password",
    )
    init_db_mock, session_factory_mock, _ = cli_database_mocks

    exit_code = bootstrap_super_admin.main(CLI_ARGS)

    captured = capsys.readouterr()
    combined_output = captured.out + captured.err

    assert exit_code != 0
    create_super_admin_mock.assert_not_called()
    assert "passwords do not match" in captured.err.lower()
    assert "first-password" not in combined_output
    assert "second-password" not in combined_output

    init_db_mock.assert_not_called()
    session_factory_mock.assert_not_called()


def test_bootstrap_cli_creates_super_admin(
    monkeypatch,
    capsys,
    create_super_admin_mock,
    cli_database_mocks,
):
    """Tests happy path for creating a super admin with bootstrap"""
    _mock_password_prompts(
        monkeypatch,
        MATCHING_PASSWORD,
        MATCHING_PASSWORD,
    )

    init_db_mock, session_factory_mock, fake_session = (
        cli_database_mocks
    )

    exit_code = bootstrap_super_admin.main(CLI_ARGS)

    captured = capsys.readouterr()
    combined_output = captured.out + captured.err

    assert exit_code == 0
    init_db_mock.assert_called_once_with()
    create_super_admin_mock.assert_called_once_with(
        session=fake_session,
        email=CLI_EMAIL,
        full_name=CLI_FULL_NAME,
        password=MATCHING_PASSWORD
    )
    assert "created successfully" in captured.out.lower()
    assert captured.err == ""
    assert MATCHING_PASSWORD not in combined_output

    session_factory_mock.assert_called_once_with(
        bootstrap_super_admin.engine
    )


def test_bootstrap_cli_existing_super_admin(
    monkeypatch,
    capsys,
    cli_database_mocks,
    create_super_admin_mock,
):
    """Tests the bootstrap cli when a super admin already exists"""
    _mock_password_prompts(
        monkeypatch,
        MATCHING_PASSWORD,
        MATCHING_PASSWORD,
    )

    init_db_mock, session_factory_mock, fake_session = (
        cli_database_mocks
    )

    create_super_admin_mock.side_effect = (
        bootstrap.SuperAdminAlreadyExistsError(
            "A user with this email already exists"
        )
    )
    exit_code = bootstrap_super_admin.main(CLI_ARGS)

    captured = capsys.readouterr()
    combined_output = captured.out + captured.err

    assert exit_code != 0
    init_db_mock.assert_called_once_with()
    create_super_admin_mock.assert_called_once_with(
        session=fake_session,
        email=CLI_EMAIL,
        full_name=CLI_FULL_NAME,
        password=MATCHING_PASSWORD
    )
    assert "already exists" in captured.err.lower()
    assert "created successfully" not in captured.out
    assert MATCHING_PASSWORD not in combined_output
    session_factory_mock.assert_called_once_with(
        bootstrap_super_admin.engine
    )


def test_bootstrap_cli_rejects_empty_password(
    monkeypatch,
    capsys,
    cli_database_mocks,
    create_super_admin_mock
):
    """Tests that empty password is rejected by create_super_admin"""
    _mock_password_prompts(monkeypatch, "", "")
    init_db_mock, session_factory_mock, _ = cli_database_mocks

    exit_code = bootstrap_super_admin.main(CLI_ARGS)
    captured = capsys.readouterr()

    assert exit_code == 1
    assert "password cannot be empty" in captured.err.lower()
    init_db_mock.assert_not_called()
    session_factory_mock.assert_not_called()
    create_super_admin_mock.assert_not_called()


@pytest.mark.parametrize(
    "forbidden_option",
    [
        "--password",
        "--role",
        "--department-id",
        "--doctor-id"
    ]
)
def test_bootstrap_cli_rejects_forbidden_options(
    monkeypatch,
    create_super_admin_mock,
    cli_database_mocks,
    forbidden_option
):
    """Tests that forbidden options are rejected"""
    init_db_mock, session_factory_mock, _ = cli_database_mocks
    getpass_mock = Mock()

    monkeypatch.setattr(
        getpass,
        "getpass",
        getpass_mock
    )

    with pytest.raises(SystemExit) as exc_info:
        bootstrap_super_admin.main(
            [*CLI_ARGS, forbidden_option, "value"]
        )

    assert exc_info.value.code == 2
    init_db_mock.assert_not_called()
    session_factory_mock.assert_not_called()
    create_super_admin_mock.assert_not_called()
    getpass_mock.assert_not_called()


def test_bootstrap_cli_persists_super_admin(
    session,
    monkeypatch,
    capsys,
    cli_in_memory_database
):
    """Tests that bootstrap cli creates a super admin in the db"""
    _mock_password_prompts(
        monkeypatch,
        MATCHING_PASSWORD,
        MATCHING_PASSWORD,
    )

    init_db_mock, session_factory_mock = cli_in_memory_database

    exit_code = bootstrap_super_admin.main(CLI_ARGS)
    captured = capsys.readouterr()
    combined_output = captured.out + captured.err

    assert exit_code == 0
    retrieved_user = user_repository.get_user_by_email(session, CLI_EMAIL)

    assert retrieved_user is not None
    assert retrieved_user.role == UserRole.SUPER_ADMIN
    assert retrieved_user.department_id is None
    assert retrieved_user.doctor_id is None
    assert verify_password(MATCHING_PASSWORD, retrieved_user.hashed_password)
    assert retrieved_user.hashed_password != MATCHING_PASSWORD
    assert MATCHING_PASSWORD not in combined_output
    init_db_mock.assert_called_once_with()
    session_factory_mock.assert_called_once_with(
        bootstrap_super_admin.engine
    )


from src.department.models import Department as DepartmentModel

DEPT_CLI_EMAIL = "deptadmin@test.com"
DEPT_CLI_FULL_NAME = "Dr. Cardio Admin"
DEPT_CLI_NAME = "Cardiology"
DEPT_CLI_CODE = "CARD"
DEPT_CLI_ARGS = (
    "--department-name",
    DEPT_CLI_NAME,
    "--department-code",
    DEPT_CLI_CODE,
    "--email",
    DEPT_CLI_EMAIL,
    "--full-name",
    DEPT_CLI_FULL_NAME,
)


@pytest.fixture(name="dept_cli_database_mocks")
def dept_cli_database_mocks_fixture(monkeypatch):
    """Creates a reusable cli database mock for department admin cli tests."""
    from scripts import bootstrap_department_admin

    init_db_mock = Mock()
    fake_session = object()
    session_factory_mock = Mock(return_value=nullcontext(fake_session))

    monkeypatch.setattr(bootstrap_department_admin, "init_db", init_db_mock)
    monkeypatch.setattr(bootstrap_department_admin, "Session", session_factory_mock)

    return init_db_mock, session_factory_mock, fake_session


@pytest.fixture(name="create_dept_admin_mock")
def create_dept_admin_mock_fixture(monkeypatch):
    """Creates a reusable create_department_admin service mock."""
    service_mock = Mock()
    monkeypatch.setattr(bootstrap, "create_department_admin", service_mock)
    return service_mock


@pytest.fixture(name="dept_cli_in_memory_database")
def dept_cli_in_memory_database_fixture(monkeypatch, session):
    """Creates an in-memory database mock for department admin cli persistence tests."""
    from scripts import bootstrap_department_admin

    init_db_mock = Mock()
    session_factory_mock = Mock(return_value=nullcontext(session))

    monkeypatch.setattr(bootstrap_department_admin, "init_db", init_db_mock)
    monkeypatch.setattr(bootstrap_department_admin, "Session", session_factory_mock)

    return init_db_mock, session_factory_mock


def test_bootstrap_department_admin_cli_rejects_mismatched_passwords(
    monkeypatch,
    capsys,
    create_dept_admin_mock,
    dept_cli_database_mocks,
):
    """Tests the bootstrap department admin cli with mismatched passwords."""
    from scripts import bootstrap_department_admin

    _mock_password_prompts(monkeypatch, "first-pass", "second-pass")
    init_db_mock, session_factory_mock, _ = dept_cli_database_mocks

    exit_code = bootstrap_department_admin.main(DEPT_CLI_ARGS)
    captured = capsys.readouterr()

    assert exit_code != 0
    create_dept_admin_mock.assert_not_called()
    assert "passwords do not match" in captured.err.lower()
    init_db_mock.assert_not_called()
    session_factory_mock.assert_not_called()


def test_bootstrap_department_admin_cli_rejects_empty_password(
    monkeypatch,
    capsys,
    dept_cli_database_mocks,
    create_dept_admin_mock,
):
    """Tests that empty password is rejected by bootstrap department admin cli."""
    from scripts import bootstrap_department_admin

    _mock_password_prompts(monkeypatch, "", "")
    init_db_mock, session_factory_mock, _ = dept_cli_database_mocks

    exit_code = bootstrap_department_admin.main(DEPT_CLI_ARGS)
    captured = capsys.readouterr()

    assert exit_code == 1
    assert "password cannot be empty" in captured.err.lower()
    init_db_mock.assert_not_called()
    session_factory_mock.assert_not_called()
    create_dept_admin_mock.assert_not_called()


def test_bootstrap_department_admin_cli_creates_department_admin(
    monkeypatch,
    capsys,
    create_dept_admin_mock,
    dept_cli_database_mocks,
):
    """Tests happy path for creating a department admin via CLI."""
    from scripts import bootstrap_department_admin

    _mock_password_prompts(monkeypatch, MATCHING_PASSWORD, MATCHING_PASSWORD)
    init_db_mock, session_factory_mock, fake_session = dept_cli_database_mocks

    fake_dept = DepartmentModel(name=DEPT_CLI_NAME, code=DEPT_CLI_CODE)
    fake_user = Mock(id="user-123", email=DEPT_CLI_EMAIL)
    create_dept_admin_mock.return_value = (fake_dept, fake_user)

    exit_code = bootstrap_department_admin.main(DEPT_CLI_ARGS)
    captured = capsys.readouterr()

    assert exit_code == 0
    init_db_mock.assert_called_once_with()
    create_dept_admin_mock.assert_called_once_with(
        session=fake_session,
        department_name=DEPT_CLI_NAME,
        department_code=DEPT_CLI_CODE,
        email=DEPT_CLI_EMAIL,
        full_name=DEPT_CLI_FULL_NAME,
        password=MATCHING_PASSWORD,
    )
    assert "created successfully" in captured.out.lower()
    assert MATCHING_PASSWORD not in (captured.out + captured.err)


def test_bootstrap_department_admin_cli_existing_user(
    monkeypatch,
    capsys,
    dept_cli_database_mocks,
    create_dept_admin_mock,
):
    """Tests CLI behavior when user email already exists."""
    from scripts import bootstrap_department_admin

    _mock_password_prompts(monkeypatch, MATCHING_PASSWORD, MATCHING_PASSWORD)
    create_dept_admin_mock.side_effect = bootstrap.DepartmentAdminAlreadyExistsError(
        "A user with this email already exists"
    )

    exit_code = bootstrap_department_admin.main(DEPT_CLI_ARGS)
    captured = capsys.readouterr()

    assert exit_code != 0
    assert "already exists" in captured.err.lower()


def test_bootstrap_department_admin_cli_rejects_weak_password(
    monkeypatch,
    capsys,
    dept_cli_database_mocks,
    create_dept_admin_mock,
):
    """Tests CLI behavior when password fails strength validation."""
    from scripts import bootstrap_department_admin
    from src.auth.security import WeakPasswordError

    _mock_password_prompts(monkeypatch, "weak", "weak")
    create_dept_admin_mock.side_effect = WeakPasswordError(
        "Password must contain at least 10 characters long."
    )

    exit_code = bootstrap_department_admin.main(DEPT_CLI_ARGS)
    captured = capsys.readouterr()

    assert exit_code != 0
    assert "password must" in captured.err.lower()


@pytest.mark.parametrize(
    "forbidden_option",
    ["--password", "--role", "--doctor-id"],
)
def test_bootstrap_department_admin_cli_rejects_forbidden_options(
    monkeypatch,
    dept_cli_database_mocks,
    forbidden_option,
):
    """Tests that forbidden CLI flags are rejected."""
    from scripts import bootstrap_department_admin

    with pytest.raises(SystemExit) as exc_info:
        bootstrap_department_admin.main([*DEPT_CLI_ARGS, forbidden_option, "value"])

    assert exc_info.value.code == 2


def test_bootstrap_department_admin_cli_persists_department_and_user(
    session,
    monkeypatch,
    capsys,
    dept_cli_in_memory_database,
):
    """Tests that CLI persists both Department and User into the database."""
    from scripts import bootstrap_department_admin

    _mock_password_prompts(monkeypatch, MATCHING_PASSWORD, MATCHING_PASSWORD)

    exit_code = bootstrap_department_admin.main(DEPT_CLI_ARGS)
    captured = capsys.readouterr()

    assert exit_code == 0
    retrieved_user = user_repository.get_user_by_email(session, DEPT_CLI_EMAIL)
    assert retrieved_user is not None
    assert retrieved_user.role == UserRole.DEPARTMENT_ADMIN
    assert retrieved_user.doctor_id is None
    assert verify_password(MATCHING_PASSWORD, retrieved_user.hashed_password)

    retrieved_dept = session.get(DepartmentModel, retrieved_user.department_id)
    assert retrieved_dept is not None
    assert retrieved_dept.name == DEPT_CLI_NAME
    assert retrieved_dept.code == DEPT_CLI_CODE
