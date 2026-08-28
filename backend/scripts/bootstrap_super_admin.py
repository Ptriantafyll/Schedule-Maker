"""
Create super admin script
"""

import argparse
import getpass
import sys

from src.auth import bootstrap


def build_parser() -> argparse.ArgumentParser:
    """Create the command-line argument parser."""
    parser = argparse.ArgumentParser(
        description="Create a bootstrap super-admin account"
    )
    parser.add_argument(
        "--email",
        required=True,
        help="Login email for the super-admin"
    )
    parser.add_argument(
        "--full-name",
        required=True,
        help="Display name for the super-admin"
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    """Run the super-admin bootstrap command"""
    parser = build_parser()
    parser.parse_args(argv)

    password = getpass.getpass("Password: ")
    password_confirmation = getpass.getpass("Confirm password: ")

    if password != password_confirmation:
        print("Passwords do not match", file=sys.stderr)
        return 1

    print(
        "Super admin creation is not implemented yet",
        file=sys.stderr
    )
    return 1


if __name__ == "__main__":
    raise SystemError(main())
