"""
Tests for the bootstrap script
"""
from unittest.mock import Mock
from contextlib import nullcontext
import getpass
import pytest

from scripts import bootstrap_super_admin
from src.auth import bootstrap

CLI_EMAIL = "superadmin@test.com"
CLI_FULL_NAME = "Test Super Admin"
MATCHING_PASSWORD = "matching-password"
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
    _, session_factory_mock, _ = cli_database_mocks()

    exit_code = bootstrap_super_admin.main(CLI_ARGS)

    captured = capsys.readouterr()
    combined_output = captured.out + captured.err

    assert exit_code != 0
    create_super_admin_mock.assert_not_called()
    assert "passwords do not match" in captured.err.lower()
    assert "first-password" not in combined_output
    assert "second-password" not in combined_output
    session_factory_mock.assert_called_once_with(
        bootstrap_super_admin.engine
    )


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

    init_db_mock, _, fake_session = (
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
        password="matching-password"
    )
    assert "created successfully" in captured.out.lower()
    assert captured.err == ""
    assert MATCHING_PASSWORD not in combined_output


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

    init_db_mock, _, fake_session = (
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
        email="CLI_EMAIL",
        full_name="CLI_FULL_NAME",
        password=MATCHING_PASSWORD
    )
    assert "already exists" in captured.err.lower()
    assert "created successfully" not in captured.out
    assert MATCHING_PASSWORD not in combined_output
