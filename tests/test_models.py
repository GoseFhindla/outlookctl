"""Tests for outlookctl data models."""

import pytest
from outlookctl.models import (
    MessageId,
    EmailAddress,
    FolderInfo,
    MessageSummary,
    MessageDetail,
    ListResult,
    SearchResult,
    DraftResult,
    SendResult,
    AttachmentSaveResult,
    DoctorCheck,
    DoctorResult,
    ErrorResult,
)


class TestMessageId:
    def test_to_dict(self):
        msg_id = MessageId(entry_id="abc123", store_id="store456")
        result = msg_id.to_dict()
        assert result == {"entry_id": "abc123", "store_id": "store456"}


class TestEmailAddress:
    def test_to_dict(self):
        addr = EmailAddress(name="John Doe", email="john@example.com")
        result = addr.to_dict()
        assert result == {"name": "John Doe", "email": "john@example.com"}


class TestFolderInfo:
    def test_to_dict_minimal(self):
        folder = FolderInfo(name="Inbox")
        result = folder.to_dict()
        assert result == {"name": "Inbox"}

    def test_to_dict_with_path(self):
        folder = FolderInfo(name="Subfolder", path="Inbox/Subfolder")
        result = folder.to_dict()
        assert result == {"name": "Subfolder", "path": "Inbox/Subfolder"}


class TestMessageSummary:
    def test_to_dict_without_snippet(self):
        summary = MessageSummary(
            id=MessageId(entry_id="e1", store_id="s1"),
            received_at="2025-01-15T10:00:00",
            subject="Test Subject",
            sender=EmailAddress(name="Sender", email="sender@test.com"),
            to=["recipient@test.com"],
            cc=[],
            unread=True,
            has_attachments=False,
        )
        result = summary.to_dict()
        assert result["subject"] == "Test Subject"
        assert result["unread"] is True
        assert "body_snippet" not in result

    def test_to_dict_with_snippet(self):
        summary = MessageSummary(
            id=MessageId(entry_id="e1", store_id="s1"),
            received_at="2025-01-15T10:00:00",
            subject="Test",
            sender=EmailAddress(name="S", email="s@t.com"),
            to=[],
            cc=[],
            unread=False,
            has_attachments=False,
            body_snippet="Hello world...",
        )
        result = summary.to_dict()
        assert result["body_snippet"] == "Hello world..."


class TestListResult:
    def test_to_dict(self):
        result = ListResult(
            folder=FolderInfo(name="Inbox"),
            items=[],
        )
        output = result.to_dict()
        assert output["version"] == "1.0"
        assert output["folder"]["name"] == "Inbox"
        assert output["items"] == []


class TestSearchResult:
    def test_to_dict(self):
        result = SearchResult(
            query={"from": "test@example.com"},
            items=[],
        )
        output = result.to_dict()
        assert output["version"] == "1.0"
        assert output["query"]["from"] == "test@example.com"


class TestDraftResult:
    def test_to_dict_success(self):
        result = DraftResult(
            success=True,
            id=MessageId(entry_id="draft1", store_id="store1"),
            subject="Draft Subject",
            to=["recipient@test.com"],
        )
        output = result.to_dict()
        assert output["success"] is True
        assert output["id"]["entry_id"] == "draft1"
        assert output["subject"] == "Draft Subject"


class TestSendResult:
    def test_to_dict_success(self):
        result = SendResult(
            success=True,
            message="Sent successfully",
            sent_at="2025-01-15T10:00:00",
            to=["recipient@test.com"],
            subject="Test",
        )
        output = result.to_dict()
        assert output["success"] is True
        assert output["sent_at"] == "2025-01-15T10:00:00"


class TestDoctorResult:
    def test_to_dict(self):
        result = DoctorResult(
            all_passed=True,
            checks=[
                DoctorCheck(
                    name="test_check",
                    passed=True,
                    message="Check passed",
                )
            ],
            outlook_path="C:\\OUTLOOK.EXE",
        )
        output = result.to_dict()
        assert output["all_passed"] is True
        assert len(output["checks"]) == 1
        assert output["outlook_path"] == "C:\\OUTLOOK.EXE"


class TestErrorResult:
    def test_to_dict(self):
        result = ErrorResult(
            error="Something went wrong",
            error_code="TEST_ERROR",
            remediation="Try again",
        )
        output = result.to_dict()
        assert output["success"] is False
        assert output["error"] == "Something went wrong"
        assert output["error_code"] == "TEST_ERROR"
        assert output["remediation"] == "Try again"
