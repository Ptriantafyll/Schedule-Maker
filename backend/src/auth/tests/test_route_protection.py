"""
Route protection auth tests
"""

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
    "delete",
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


def test_public_signup_route_is_not_exposed():
    """Tests legacy signup route is not exposed"""
    assert "/api/v1/users/signup" not in app.openapi()["paths"]


def test_user_email_lookup_route_is_not_exposed():
    """Tests legacy user email lookup route is not exposed"""
    assert "/api/v1/users/{user_email}" not in app.openapi()["paths"]


def test_department_creation_method_is_not_exposed():
    """Tests department creation method is not exposed"""
    department_operations = app.openapi()["paths"][
        "/api/v1/departments/"
    ]

    assert "post" not in department_operations
