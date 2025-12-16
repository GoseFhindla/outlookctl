"""
Data models for outlookctl JSON output.

All models use dataclasses and provide to_dict() methods for JSON serialization.
"""

from dataclasses import dataclass, field, asdict
from datetime import datetime
from typing import Optional


@dataclass
class MessageId:
    """Stable identifier for an Outlook message."""
    entry_id: str
    store_id: str

    def to_dict(self) -> dict:
        return asdict(self)


@dataclass
class EmailAddress:
    """Email address with optional display name."""
    name: str
    email: str

    def to_dict(self) -> dict:
        return asdict(self)


@dataclass
class FolderInfo:
    """Basic folder information."""
    name: str
    path: Optional[str] = None
    store_id: Optional[str] = None

    def to_dict(self) -> dict:
        return {k: v for k, v in asdict(self).items() if v is not None}


@dataclass
class MessageSummary:
    """Summary of an email message for list/search results."""
    id: MessageId
    received_at: str
    subject: str
    sender: EmailAddress
    to: list[str]
    cc: list[str]
    unread: bool
    has_attachments: bool
    body_snippet: Optional[str] = None

    def to_dict(self) -> dict:
        result = {
            "id": self.id.to_dict(),
            "received_at": self.received_at,
            "subject": self.subject,
            "from": self.sender.to_dict(),
            "to": self.to,
            "cc": self.cc,
            "unread": self.unread,
            "has_attachments": self.has_attachments,
        }
        if self.body_snippet is not None:
            result["body_snippet"] = self.body_snippet
        return result


@dataclass
class MessageDetail:
    """Full message details for get command."""
    id: MessageId
    received_at: str
    subject: str
    sender: EmailAddress
    to: list[str]
    cc: list[str]
    bcc: list[str]
    unread: bool
    has_attachments: bool
    attachments: list[str]
    body: Optional[str] = None
    body_html: Optional[str] = None
    headers: Optional[dict[str, str]] = None

    def to_dict(self) -> dict:
        result = {
            "id": self.id.to_dict(),
            "received_at": self.received_at,
            "subject": self.subject,
            "from": self.sender.to_dict(),
            "to": self.to,
            "cc": self.cc,
            "bcc": self.bcc,
            "unread": self.unread,
            "has_attachments": self.has_attachments,
            "attachments": self.attachments,
        }
        if self.body is not None:
            result["body"] = self.body
        if self.body_html is not None:
            result["body_html"] = self.body_html
        if self.headers is not None:
            result["headers"] = self.headers
        return result


@dataclass
class ListResult:
    """Result of a list operation."""
    version: str = "1.0"
    folder: FolderInfo = field(default_factory=lambda: FolderInfo(name="Inbox"))
    items: list[MessageSummary] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "version": self.version,
            "folder": self.folder.to_dict(),
            "items": [item.to_dict() for item in self.items],
        }


@dataclass
class SearchResult:
    """Result of a search operation."""
    version: str = "1.0"
    query: dict = field(default_factory=dict)
    items: list[MessageSummary] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "version": self.version,
            "query": self.query,
            "items": [item.to_dict() for item in self.items],
        }


@dataclass
class DraftResult:
    """Result of a draft operation."""
    version: str = "1.0"
    success: bool = True
    id: Optional[MessageId] = None
    saved_to: str = "Drafts"
    subject: Optional[str] = None
    to: list[str] = field(default_factory=list)
    cc: list[str] = field(default_factory=list)
    attachments: list[str] = field(default_factory=list)

    def to_dict(self) -> dict:
        result = {
            "version": self.version,
            "success": self.success,
            "saved_to": self.saved_to,
        }
        if self.id:
            result["id"] = self.id.to_dict()
        if self.subject:
            result["subject"] = self.subject
        if self.to:
            result["to"] = self.to
        if self.cc:
            result["cc"] = self.cc
        if self.attachments:
            result["attachments"] = self.attachments
        return result


@dataclass
class SendResult:
    """Result of a send operation."""
    version: str = "1.0"
    success: bool = True
    message: str = ""
    sent_at: Optional[str] = None
    to: list[str] = field(default_factory=list)
    subject: Optional[str] = None

    def to_dict(self) -> dict:
        result = {
            "version": self.version,
            "success": self.success,
            "message": self.message,
        }
        if self.sent_at:
            result["sent_at"] = self.sent_at
        if self.to:
            result["to"] = self.to
        if self.subject:
            result["subject"] = self.subject
        return result


@dataclass
class AttachmentSaveResult:
    """Result of saving attachments."""
    version: str = "1.0"
    success: bool = True
    saved_files: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "version": self.version,
            "success": self.success,
            "saved_files": self.saved_files,
            "errors": self.errors,
        }


@dataclass
class DoctorCheck:
    """Single check result for doctor command."""
    name: str
    passed: bool
    message: str
    remediation: Optional[str] = None

    def to_dict(self) -> dict:
        result = {
            "name": self.name,
            "passed": self.passed,
            "message": self.message,
        }
        if self.remediation:
            result["remediation"] = self.remediation
        return result


@dataclass
class DoctorResult:
    """Result of doctor command."""
    version: str = "1.0"
    all_passed: bool = True
    checks: list[DoctorCheck] = field(default_factory=list)
    outlook_path: Optional[str] = None

    def to_dict(self) -> dict:
        return {
            "version": self.version,
            "all_passed": self.all_passed,
            "checks": [check.to_dict() for check in self.checks],
            "outlook_path": self.outlook_path,
        }


@dataclass
class ErrorResult:
    """Error result for any command."""
    version: str = "1.0"
    success: bool = False
    error: str = ""
    error_code: Optional[str] = None
    remediation: Optional[str] = None

    def to_dict(self) -> dict:
        result = {
            "version": self.version,
            "success": self.success,
            "error": self.error,
        }
        if self.error_code:
            result["error_code"] = self.error_code
        if self.remediation:
            result["remediation"] = self.remediation
        return result
