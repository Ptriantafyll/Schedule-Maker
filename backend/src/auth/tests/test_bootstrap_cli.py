"""
Tests for the bootstrap script
"""

import getpass
from unittest.mock import Mock

from scripts import bootstrap_super_admin
from src.auth import bootstrap


def test_bootstrap_cli_rejects_mismatched_passwords(monkeypatch, capsys):
    """Tests the bootstrap cli with mismatched passwords"""
    password_reader = iter(
        [
            "first-password",
            "second-password",
        ]
    )

    monkeypatch.setattr(
        getpass,
        "getpass",
        lambda _prompt: next(password_reader)
    )

    create_super_admin_mock = Mock()
    monkeypatch.setattr(
        bootstrap,
        "create_super_admin",
        create_super_admin_mock,
    )

    exit_code = bootstrap_super_admin.main(
        [
            "--email",
            "superadmin@test.com",
            "--ful-name",
            "Test Super Admin"
        ]
    )

    captured = capsys.readouterr()
    combined_output = captured.out + captured.err

    assert exit_code != 0
    create_super_admin_mock.assert_not_called()
    assert "passwords do not match" in captured.err.lower()
    assert "first-password" not in combined_output
    assert "second-password" not in combined_output
