"""
Safety gates for send operations.

This module provides confirmation checking and validation for send operations
to prevent accidental email sending.
"""

from pathlib import Path
from typing import Optional

# The exact confirmation string required
CONFIRM_STRING = "YES"


class SendConfirmationError(Exception):
    """Raised when send confirmation is missing or invalid."""
    pass


def validate_send_confirmation(
    confirm_send: Optional[str] = None,
    confirm_send_file: Optional[str] = None,
) -> bool:
    """
    Validate that proper send confirmation has been provided.

    Args:
        confirm_send: Direct confirmation string (must be exactly "YES")
        confirm_send_file: Path to file containing confirmation string

    Returns:
        True if confirmation is valid

    Raises:
        SendConfirmationError: If confirmation is missing or invalid
    """
    if confirm_send is not None:
        if confirm_send == CONFIRM_STRING:
            return True
        raise SendConfirmationError(
            f"Invalid confirmation string. Expected exactly '{CONFIRM_STRING}', "
            f"got '{confirm_send}'."
        )

    if confirm_send_file is not None:
        path = Path(confirm_send_file)
        if not path.exists():
            raise SendConfirmationError(
                f"Confirmation file not found: {confirm_send_file}"
            )
        try:
            content = path.read_text().strip()
            if content == CONFIRM_STRING:
                return True
            raise SendConfirmationError(
                f"Invalid confirmation in file. Expected exactly '{CONFIRM_STRING}', "
                f"got '{content}'."
            )
        except OSError as e:
            raise SendConfirmationError(
                f"Could not read confirmation file: {e}"
            )

    raise SendConfirmationError(
        "Send confirmation required. Use --confirm-send YES or "
        "--confirm-send-file <path> with a file containing 'YES'."
    )


def validate_unsafe_send_new(
    unsafe_send_new: bool,
    confirm_send: Optional[str] = None,
    confirm_send_file: Optional[str] = None,
) -> bool:
    """
    Validate confirmation for sending a new message directly (without draft).

    This requires both --unsafe-send-new flag AND proper confirmation.

    Args:
        unsafe_send_new: Whether --unsafe-send-new flag was provided
        confirm_send: Direct confirmation string
        confirm_send_file: Path to file containing confirmation string

    Returns:
        True if all confirmations are valid

    Raises:
        SendConfirmationError: If any confirmation is missing or invalid
    """
    if not unsafe_send_new:
        raise SendConfirmationError(
            "Sending a new message directly is not recommended. "
            "Please use 'outlookctl draft' first, then 'outlookctl send'. "
            "If you must send directly, use --unsafe-send-new --confirm-send YES"
        )

    return validate_send_confirmation(confirm_send, confirm_send_file)


def check_recipients(to: list[str], cc: list[str], bcc: list[str]) -> None:
    """
    Basic validation of recipient lists.

    Args:
        to: List of To recipients
        cc: List of CC recipients
        bcc: List of BCC recipients

    Raises:
        ValueError: If no recipients specified
    """
    all_recipients = to + cc + bcc
    if not all_recipients:
        raise ValueError("At least one recipient (To, CC, or BCC) is required.")
