"""
Route protection auth tests
"""

import pytest

from src.main import app

EXPECTED_PUBLIC_OPERATIONS = {
    ("GET", "/health"),
    ("POST", "/api/v1/auth/login"),
}

HTTP_METHODS = {
    "get",
    "post",
    "put",
    "patch",
    "deleted",
}


def test_only_expected_api_operations_are_public():
    """Tests that only the expected public endpoints are available without auth"""
    openapi_paths = app.openapi()["paths"]
    public_operations = set()

    for path, path_item in openapi_paths.items():
        for method, operation in path_item.items():
            if method not in HTTP_METHODS:
                continue

            if not operation.get("security"):
                public_operations.add(
                    (method.upper(), path)
                )

    assert public_operations == EXPECTED_PUBLIC_OPERATIONS
