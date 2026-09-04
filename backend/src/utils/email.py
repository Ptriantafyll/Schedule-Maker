"""
Helper functions for sending emails, including email validation and sending email messages.
"""


def normalize_email(email: str) -> str:
    """Normalizes an email address by converting it to lowercase and stripping whitespace."""
    return email.strip().lower()
