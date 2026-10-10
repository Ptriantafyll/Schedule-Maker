
"""
Create department admin script
"""

import argparse
import getpass
import sys

from sqlmodel import Session

from src.db.connection import engine, init_db
from src.auth import bootstrap
from src.auth.security import WeakPasswordError


def build_parser() -> argparse.ArgumentParser:
    """Create the command-line argument parser."""
    parser = argparse.ArgumentParser(
        description="Create a bootstrap department-admin account"
    )
    parser.add_argument(
        "--department-name",
        required=True,
        help="Department name for the admin"
    )
    parser.add_argument(
        "--department-code",
        required=False,
        help="Department code (derived from name if omitted)"
    )
    parser.add_argument(
        "--email",
        required=True,
        help="Login email for the department-admin"
    )
    parser.add_argument(
        "--full-name",
        required=True,
        help="Display name for the department-admin"
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    """Run the department-admin bootstrap command"""
    parser = build_parser()
    args = parser.parse_args(argv)

    password = getpass.getpass("Password: ")
    password_confirmation = getpass.getpass("Confirm password: ")

    if not password:
        print("Password cannot be empty", file=sys.stderr)
        return 1

    if password != password_confirmation:
        print("Passwords do not match", file=sys.stderr)
        return 1

    init_db()
    try:
        with Session(engine) as session:
            department, user = bootstrap.create_department_admin(
                session=session,
                department_name=args.department_name,
                department_code=args.department_code,
                email=args.email,
                full_name=args.full_name,
                password=password,
            )
    except (
        bootstrap.DepartmentAdminAlreadyExistsError,
        WeakPasswordError,
    ) as exc:
        print(str(exc), file=sys.stderr)
        return 1

    print(
        f"Department: {department.name} (CODE: {department.code}, id: {department.id})"
    )
    print(
        f"Department admin created successfully: {user.full_name}, {user.email}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
