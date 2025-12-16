"""Tests for outlookctl safety gates."""

import os
import tempfile
import pytest
from outlookctl.safety import (
    validate_send_confirmation,
    validate_unsafe_send_new,
    check_recipients,
    SendConfirmationError,
    CONFIRM_STRING,
)


class TestValidateSendConfirmation:
    def test_valid_confirm_string(self):
        assert validate_send_confirmation(confirm_send="YES") is True

    def test_invalid_confirm_string(self):
        with pytest.raises(SendConfirmationError) as exc_info:
            validate_send_confirmation(confirm_send="yes")
        assert "Expected exactly 'YES'" in str(exc_info.value)

    def test_wrong_confirm_string(self):
        with pytest.raises(SendConfirmationError) as exc_info:
            validate_send_confirmation(confirm_send="NO")
        assert "Invalid confirmation string" in str(exc_info.value)

    def test_no_confirmation(self):
        with pytest.raises(SendConfirmationError) as exc_info:
            validate_send_confirmation()
        assert "Send confirmation required" in str(exc_info.value)

    def test_valid_confirm_file(self):
        with tempfile.NamedTemporaryFile(mode="w", delete=False, suffix=".txt") as f:
            f.write("YES")
            f.flush()
            temp_path = f.name

        try:
            assert validate_send_confirmation(confirm_send_file=temp_path) is True
        finally:
            os.unlink(temp_path)

    def test_valid_confirm_file_with_whitespace(self):
        with tempfile.NamedTemporaryFile(mode="w", delete=False, suffix=".txt") as f:
            f.write("  YES  \n")
            f.flush()
            temp_path = f.name

        try:
            assert validate_send_confirmation(confirm_send_file=temp_path) is True
        finally:
            os.unlink(temp_path)

    def test_invalid_confirm_file_content(self):
        with tempfile.NamedTemporaryFile(mode="w", delete=False, suffix=".txt") as f:
            f.write("NO")
            f.flush()
            temp_path = f.name

        try:
            with pytest.raises(SendConfirmationError) as exc_info:
                validate_send_confirmation(confirm_send_file=temp_path)
            assert "Invalid confirmation in file" in str(exc_info.value)
        finally:
            os.unlink(temp_path)

    def test_missing_confirm_file(self):
        with pytest.raises(SendConfirmationError) as exc_info:
            validate_send_confirmation(confirm_send_file="/nonexistent/path.txt")
        assert "Confirmation file not found" in str(exc_info.value)


class TestValidateUnsafeSendNew:
    def test_without_unsafe_flag(self):
        with pytest.raises(SendConfirmationError) as exc_info:
            validate_unsafe_send_new(
                unsafe_send_new=False,
                confirm_send="YES",
            )
        assert "not recommended" in str(exc_info.value)
        assert "outlookctl draft" in str(exc_info.value)

    def test_with_unsafe_flag_and_confirm(self):
        assert validate_unsafe_send_new(
            unsafe_send_new=True,
            confirm_send="YES",
        ) is True

    def test_with_unsafe_flag_without_confirm(self):
        with pytest.raises(SendConfirmationError) as exc_info:
            validate_unsafe_send_new(
                unsafe_send_new=True,
                confirm_send=None,
            )
        assert "Send confirmation required" in str(exc_info.value)


class TestCheckRecipients:
    def test_valid_to_only(self):
        # Should not raise
        check_recipients(to=["test@example.com"], cc=[], bcc=[])

    def test_valid_cc_only(self):
        # Should not raise
        check_recipients(to=[], cc=["test@example.com"], bcc=[])

    def test_valid_bcc_only(self):
        # Should not raise
        check_recipients(to=[], cc=[], bcc=["test@example.com"])

    def test_valid_multiple(self):
        # Should not raise
        check_recipients(
            to=["a@test.com", "b@test.com"],
            cc=["c@test.com"],
            bcc=["d@test.com"],
        )

    def test_no_recipients(self):
        with pytest.raises(ValueError) as exc_info:
            check_recipients(to=[], cc=[], bcc=[])
        assert "At least one recipient" in str(exc_info.value)
