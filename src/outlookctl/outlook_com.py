"""
Outlook COM automation wrapper.

This module provides a clean interface to Outlook's COM object model
via pywin32. It handles connection management, error handling, and
data extraction from Outlook objects.
"""

import os
import subprocess
import time
from datetime import datetime
from pathlib import Path
from typing import Optional, Iterator, Any

from .models import (
    MessageId,
    EmailAddress,
    FolderInfo,
    MessageSummary,
    MessageDetail,
    DoctorCheck,
    DoctorResult,
)


# Outlook folder type constants (OlDefaultFolders enumeration)
OL_FOLDER_INBOX = 6
OL_FOLDER_SENT_MAIL = 5
OL_FOLDER_DRAFTS = 16
OL_FOLDER_DELETED_ITEMS = 3
OL_FOLDER_OUTBOX = 4
OL_FOLDER_JUNK = 23

# Map of common folder names to constants
FOLDER_MAP = {
    "inbox": OL_FOLDER_INBOX,
    "sent": OL_FOLDER_SENT_MAIL,
    "drafts": OL_FOLDER_DRAFTS,
    "deleted": OL_FOLDER_DELETED_ITEMS,
    "outbox": OL_FOLDER_OUTBOX,
    "junk": OL_FOLDER_JUNK,
}

# Common Outlook installation paths
OUTLOOK_PATHS = [
    r"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE",
    r"C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE",
    r"C:\Program Files\Microsoft Office\Office16\OUTLOOK.EXE",
    r"C:\Program Files (x86)\Microsoft Office\Office16\OUTLOOK.EXE",
]


class OutlookError(Exception):
    """Base exception for Outlook COM errors."""
    pass


class OutlookNotAvailableError(OutlookError):
    """Raised when Outlook COM object cannot be accessed."""
    pass


class NewOutlookDetectedError(OutlookError):
    """Raised when New Outlook is detected (COM not supported)."""
    pass


class FolderNotFoundError(OutlookError):
    """Raised when a folder cannot be found."""
    pass


class MessageNotFoundError(OutlookError):
    """Raised when a message cannot be found."""
    pass


def _import_win32com():
    """Import win32com with helpful error if not available."""
    try:
        import win32com.client
        import pythoncom
        return win32com.client, pythoncom
    except ImportError as e:
        raise OutlookError(
            "pywin32 is not installed. Run: uv add pywin32"
        ) from e


def get_outlook_app(retry_count: int = 3, retry_delay: float = 1.0):
    """
    Get a connection to the Outlook Application COM object.

    Args:
        retry_count: Number of times to retry connection
        retry_delay: Delay between retries in seconds

    Returns:
        Outlook.Application COM object

    Raises:
        OutlookNotAvailableError: If Outlook cannot be accessed
        NewOutlookDetectedError: If New Outlook is detected
    """
    win32com_client, pythoncom = _import_win32com()

    last_error = None
    for attempt in range(retry_count):
        try:
            # Initialize COM for this thread
            pythoncom.CoInitialize()

            # Try to connect to running Outlook
            outlook = win32com_client.Dispatch("Outlook.Application")

            # Verify we got a valid object by accessing a property
            _ = outlook.Name

            return outlook

        except Exception as e:
            last_error = e
            if attempt < retry_count - 1:
                time.sleep(retry_delay)

    # Check if this might be New Outlook
    error_msg = str(last_error).lower()
    if "class not registered" in error_msg or "invalid class string" in error_msg:
        raise OutlookNotAvailableError(
            "Outlook COM automation unavailable. This could mean:\n"
            "1. Classic Outlook is not installed or not running\n"
            "2. New Outlook is active (COM not supported)\n"
            "3. Outlook COM objects are not registered\n\n"
            "Solution: Start Classic Outlook and try again."
        )

    raise OutlookNotAvailableError(
        f"Could not connect to Outlook: {last_error}"
    )


def find_outlook_executable() -> Optional[str]:
    """Find the Outlook executable path."""
    for path in OUTLOOK_PATHS:
        if os.path.exists(path):
            return path
    return None


def start_outlook(wait_seconds: int = 10) -> bool:
    """
    Attempt to start Classic Outlook.

    Args:
        wait_seconds: How long to wait for Outlook to start

    Returns:
        True if Outlook was started, False otherwise
    """
    outlook_path = find_outlook_executable()
    if not outlook_path:
        return False

    try:
        subprocess.Popen([outlook_path], shell=False)
        time.sleep(wait_seconds)
        return True
    except Exception:
        return False


def get_namespace(outlook_app):
    """Get the MAPI namespace from Outlook."""
    return outlook_app.GetNamespace("MAPI")


def get_default_folder(outlook_app, folder_type: int):
    """
    Get a default folder by type.

    Args:
        outlook_app: Outlook Application COM object
        folder_type: OlDefaultFolders constant

    Returns:
        Folder COM object
    """
    namespace = get_namespace(outlook_app)
    return namespace.GetDefaultFolder(folder_type)


def get_folder_by_name(outlook_app, folder_name: str):
    """
    Get a folder by name from the default store.

    Args:
        outlook_app: Outlook Application COM object
        folder_name: Name of the folder to find

    Returns:
        Folder COM object

    Raises:
        FolderNotFoundError: If folder not found
    """
    namespace = get_namespace(outlook_app)
    root_folder = namespace.Folders.Item(1)  # Default store

    def search_folder(parent, name):
        for folder in parent.Folders:
            if folder.Name.lower() == name.lower():
                return folder
            # Search subfolders
            try:
                result = search_folder(folder, name)
                if result:
                    return result
            except Exception:
                pass
        return None

    folder = search_folder(root_folder, folder_name)
    if not folder:
        raise FolderNotFoundError(f"Folder not found: {folder_name}")
    return folder


def get_folder_by_path(outlook_app, folder_path: str):
    """
    Get a folder by path (e.g., "Inbox/Subfolder").

    Args:
        outlook_app: Outlook Application COM object
        folder_path: Path to the folder, separated by /

    Returns:
        Folder COM object

    Raises:
        FolderNotFoundError: If folder not found
    """
    namespace = get_namespace(outlook_app)
    parts = folder_path.strip("/").split("/")

    # Start with the default store's root
    root_folder = namespace.Folders.Item(1)
    current = root_folder

    for part in parts:
        found = False
        for folder in current.Folders:
            if folder.Name.lower() == part.lower():
                current = folder
                found = True
                break
        if not found:
            raise FolderNotFoundError(f"Folder path not found: {folder_path}")

    return current


def resolve_folder(outlook_app, folder_spec: str):
    """
    Resolve a folder specification to a folder object.

    Args:
        outlook_app: Outlook Application COM object
        folder_spec: One of:
            - "inbox", "sent", "drafts", etc. (default folders)
            - "by-name:<name>" (search by name)
            - "by-path:<path>" (search by path)

    Returns:
        Tuple of (Folder COM object, FolderInfo)
    """
    folder_spec_lower = folder_spec.lower()

    if folder_spec_lower.startswith("by-name:"):
        name = folder_spec[8:]
        folder = get_folder_by_name(outlook_app, name)
        return folder, FolderInfo(name=folder.Name)

    if folder_spec_lower.startswith("by-path:"):
        path = folder_spec[8:]
        folder = get_folder_by_path(outlook_app, path)
        return folder, FolderInfo(name=folder.Name, path=path)

    # Default folder
    if folder_spec_lower in FOLDER_MAP:
        folder = get_default_folder(outlook_app, FOLDER_MAP[folder_spec_lower])
        return folder, FolderInfo(name=folder.Name)

    raise FolderNotFoundError(
        f"Unknown folder specification: {folder_spec}. "
        f"Use one of: {', '.join(FOLDER_MAP.keys())}, by-name:<name>, or by-path:<path>"
    )


def extract_email_address(recipient) -> EmailAddress:
    """Extract email address from a recipient or sender object."""
    try:
        name = str(recipient.Name) if hasattr(recipient, "Name") else ""
        # Try to get the SMTP address
        email = ""
        if hasattr(recipient, "Address"):
            email = str(recipient.Address)
        if hasattr(recipient, "PropertyAccessor"):
            try:
                # PR_SMTP_ADDRESS
                smtp = recipient.PropertyAccessor.GetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x39FE001F"
                )
                if smtp:
                    email = smtp
            except Exception:
                pass
        return EmailAddress(name=name, email=email)
    except Exception:
        return EmailAddress(name="", email="")


def extract_recipients(recipients) -> list[str]:
    """Extract list of email addresses from recipients collection."""
    result = []
    try:
        for i in range(1, recipients.Count + 1):
            recip = recipients.Item(i)
            addr = extract_email_address(recip)
            if addr.email:
                result.append(addr.email)
            elif addr.name:
                result.append(addr.name)
    except Exception:
        pass
    return result


def format_datetime(dt) -> str:
    """Format a COM datetime to ISO format string."""
    if dt is None:
        return ""
    try:
        if hasattr(dt, "isoformat"):
            return dt.isoformat()
        # Convert from COM date
        return datetime.fromtimestamp(dt).isoformat()
    except Exception:
        return str(dt)


def extract_message_summary(
    mail_item,
    include_body_snippet: bool = False,
    body_snippet_chars: int = 200,
) -> MessageSummary:
    """
    Extract a MessageSummary from a MailItem COM object.

    Args:
        mail_item: Outlook MailItem COM object
        include_body_snippet: Whether to include body snippet
        body_snippet_chars: Max characters for body snippet

    Returns:
        MessageSummary object
    """
    # Extract sender
    sender = EmailAddress(name="", email="")
    try:
        sender = EmailAddress(
            name=str(mail_item.SenderName or ""),
            email=str(mail_item.SenderEmailAddress or ""),
        )
    except Exception:
        pass

    # Extract recipients
    to_list = []
    cc_list = []
    try:
        for i in range(1, mail_item.Recipients.Count + 1):
            recip = mail_item.Recipients.Item(i)
            addr = extract_email_address(recip)
            addr_str = addr.email if addr.email else addr.name
            if recip.Type == 1:  # olTo
                to_list.append(addr_str)
            elif recip.Type == 2:  # olCC
                cc_list.append(addr_str)
    except Exception:
        pass

    # Extract body snippet if requested
    body_snippet = None
    if include_body_snippet:
        try:
            body = str(mail_item.Body or "")
            body_snippet = body[:body_snippet_chars].strip()
            if len(body) > body_snippet_chars:
                body_snippet += "..."
        except Exception:
            body_snippet = ""

    return MessageSummary(
        id=MessageId(
            entry_id=str(mail_item.EntryID),
            store_id=str(mail_item.Parent.StoreID),
        ),
        received_at=format_datetime(mail_item.ReceivedTime),
        subject=str(mail_item.Subject or ""),
        sender=sender,
        to=to_list,
        cc=cc_list,
        unread=bool(mail_item.UnRead),
        has_attachments=mail_item.Attachments.Count > 0,
        body_snippet=body_snippet,
    )


def extract_message_detail(
    mail_item,
    include_body: bool = False,
    max_body_chars: Optional[int] = None,
    include_headers: bool = False,
) -> MessageDetail:
    """
    Extract full MessageDetail from a MailItem COM object.

    Args:
        mail_item: Outlook MailItem COM object
        include_body: Whether to include full body
        max_body_chars: Max characters for body (None = unlimited)
        include_headers: Whether to include headers

    Returns:
        MessageDetail object
    """
    # Extract sender
    sender = EmailAddress(name="", email="")
    try:
        sender = EmailAddress(
            name=str(mail_item.SenderName or ""),
            email=str(mail_item.SenderEmailAddress or ""),
        )
    except Exception:
        pass

    # Extract recipients by type
    to_list = []
    cc_list = []
    bcc_list = []
    try:
        for i in range(1, mail_item.Recipients.Count + 1):
            recip = mail_item.Recipients.Item(i)
            addr = extract_email_address(recip)
            addr_str = addr.email if addr.email else addr.name
            if recip.Type == 1:  # olTo
                to_list.append(addr_str)
            elif recip.Type == 2:  # olCC
                cc_list.append(addr_str)
            elif recip.Type == 3:  # olBCC
                bcc_list.append(addr_str)
    except Exception:
        pass

    # Extract attachments
    attachment_names = []
    try:
        for i in range(1, mail_item.Attachments.Count + 1):
            att = mail_item.Attachments.Item(i)
            attachment_names.append(str(att.FileName))
    except Exception:
        pass

    # Extract body if requested
    body = None
    body_html = None
    if include_body:
        try:
            body_text = str(mail_item.Body or "")
            if max_body_chars and len(body_text) > max_body_chars:
                body = body_text[:max_body_chars] + "..."
            else:
                body = body_text
        except Exception:
            pass

        try:
            body_html = str(mail_item.HTMLBody or "")
            if max_body_chars and len(body_html) > max_body_chars:
                body_html = body_html[:max_body_chars] + "..."
        except Exception:
            pass

    # Extract headers if requested
    headers = None
    if include_headers:
        try:
            headers = {}
            # Get transport message headers
            prop_accessor = mail_item.PropertyAccessor
            header_prop = "http://schemas.microsoft.com/mapi/proptag/0x007D001F"
            raw_headers = prop_accessor.GetProperty(header_prop)
            if raw_headers:
                for line in str(raw_headers).split("\n"):
                    if ": " in line:
                        key, value = line.split(": ", 1)
                        headers[key.strip()] = value.strip()
        except Exception:
            headers = None

    return MessageDetail(
        id=MessageId(
            entry_id=str(mail_item.EntryID),
            store_id=str(mail_item.Parent.StoreID),
        ),
        received_at=format_datetime(mail_item.ReceivedTime),
        subject=str(mail_item.Subject or ""),
        sender=sender,
        to=to_list,
        cc=cc_list,
        bcc=bcc_list,
        unread=bool(mail_item.UnRead),
        has_attachments=mail_item.Attachments.Count > 0,
        attachments=attachment_names,
        body=body,
        body_html=body_html,
        headers=headers,
    )


def get_message_by_id(outlook_app, entry_id: str, store_id: str):
    """
    Get a message by its entry ID and store ID.

    Args:
        outlook_app: Outlook Application COM object
        entry_id: Message entry ID
        store_id: Store ID

    Returns:
        MailItem COM object

    Raises:
        MessageNotFoundError: If message not found
    """
    try:
        namespace = get_namespace(outlook_app)
        return namespace.GetItemFromID(entry_id, store_id)
    except Exception as e:
        raise MessageNotFoundError(
            f"Message not found with entry_id={entry_id}: {e}"
        )


def list_messages(
    outlook_app,
    folder_spec: str = "inbox",
    count: int = 10,
    unread_only: bool = False,
    since: Optional[datetime] = None,
    until: Optional[datetime] = None,
    include_body_snippet: bool = False,
    body_snippet_chars: int = 200,
) -> Iterator[MessageSummary]:
    """
    List messages from a folder.

    Args:
        outlook_app: Outlook Application COM object
        folder_spec: Folder specification
        count: Maximum number of messages to return
        unread_only: Only return unread messages
        since: Only messages received after this date
        until: Only messages received before this date
        include_body_snippet: Include body snippet
        body_snippet_chars: Max chars for body snippet

    Yields:
        MessageSummary objects
    """
    folder, _ = resolve_folder(outlook_app, folder_spec)
    items = folder.Items
    items.Sort("[ReceivedTime]", True)  # Sort descending (newest first)

    yielded = 0
    for item in items:
        if yielded >= count:
            break

        try:
            # Skip non-mail items
            if item.Class != 43:  # olMail
                continue

            # Filter by unread
            if unread_only and not item.UnRead:
                continue

            # Filter by date range
            if since or until:
                received = item.ReceivedTime
                if since and received < since:
                    continue
                if until and received > until:
                    continue

            yield extract_message_summary(
                item,
                include_body_snippet=include_body_snippet,
                body_snippet_chars=body_snippet_chars,
            )
            yielded += 1

        except Exception:
            # Skip items that can't be processed
            continue


def search_messages(
    outlook_app,
    folder_spec: str = "inbox",
    query: Optional[str] = None,
    from_filter: Optional[str] = None,
    subject_contains: Optional[str] = None,
    unread_only: bool = False,
    since: Optional[datetime] = None,
    until: Optional[datetime] = None,
    count: int = 50,
    include_body_snippet: bool = False,
    body_snippet_chars: int = 200,
) -> Iterator[MessageSummary]:
    """
    Search messages with various filters.

    Args:
        outlook_app: Outlook Application COM object
        folder_spec: Folder to search in
        query: Free text search (subject/body)
        from_filter: Filter by sender
        subject_contains: Filter by subject content
        unread_only: Only unread messages
        since: Only messages after this date
        until: Only messages before this date
        count: Maximum results
        include_body_snippet: Include body snippet
        body_snippet_chars: Max chars for snippet

    Yields:
        MessageSummary objects
    """
    folder, _ = resolve_folder(outlook_app, folder_spec)

    # Build DASL filter for more efficient searching
    filters = []

    if from_filter:
        # Search in sender name or email
        filters.append(
            f"@SQL=\"urn:schemas:httpmail:fromemail\" LIKE '%{from_filter}%' "
            f"OR \"urn:schemas:httpmail:fromname\" LIKE '%{from_filter}%'"
        )

    if subject_contains:
        filters.append(
            f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{subject_contains}%'"
        )

    if unread_only:
        filters.append("@SQL=\"urn:schemas:httpmail:read\" = 0")

    if since:
        filters.append(
            f"@SQL=\"urn:schemas:httpmail:datereceived\" >= '{since.strftime('%Y-%m-%d')}'"
        )

    if until:
        filters.append(
            f"@SQL=\"urn:schemas:httpmail:datereceived\" <= '{until.strftime('%Y-%m-%d')}'"
        )

    items = folder.Items
    items.Sort("[ReceivedTime]", True)

    # Apply filters if we have them
    if filters:
        try:
            filter_str = " AND ".join(f"({f})" for f in filters)
            items = items.Restrict(filter_str)
        except Exception:
            # Fall back to manual filtering if DASL fails
            pass

    yielded = 0
    for item in items:
        if yielded >= count:
            break

        try:
            if item.Class != 43:  # olMail
                continue

            # Manual filtering for query (body/subject search)
            if query:
                query_lower = query.lower()
                subject = str(item.Subject or "").lower()
                body = str(item.Body or "").lower()
                if query_lower not in subject and query_lower not in body:
                    continue

            yield extract_message_summary(
                item,
                include_body_snippet=include_body_snippet,
                body_snippet_chars=body_snippet_chars,
            )
            yielded += 1

        except Exception:
            continue


def create_draft(
    outlook_app,
    to: list[str],
    cc: list[str] = None,
    bcc: list[str] = None,
    subject: str = "",
    body_text: str = None,
    body_html: str = None,
    attachments: list[str] = None,
    reply_to_entry_id: str = None,
    reply_to_store_id: str = None,
) -> tuple[str, str]:
    """
    Create a draft email.

    Args:
        outlook_app: Outlook Application COM object
        to: List of To recipients
        cc: List of CC recipients
        bcc: List of BCC recipients
        subject: Email subject
        body_text: Plain text body
        body_html: HTML body (takes precedence over body_text)
        attachments: List of file paths to attach
        reply_to_entry_id: Entry ID of message to reply to
        reply_to_store_id: Store ID of message to reply to

    Returns:
        Tuple of (entry_id, store_id) of the created draft

    Raises:
        OutlookError: If draft creation fails
    """
    cc = cc or []
    bcc = bcc or []
    attachments = attachments or []

    try:
        # Create the mail item
        if reply_to_entry_id and reply_to_store_id:
            # Get the original message and create a reply
            original = get_message_by_id(outlook_app, reply_to_entry_id, reply_to_store_id)
            mail = original.Reply()
            # Clear auto-generated body for reply
            if body_html:
                mail.HTMLBody = body_html
            elif body_text:
                mail.Body = body_text
        else:
            mail = outlook_app.CreateItem(0)  # olMailItem

            # Set body
            if body_html:
                mail.HTMLBody = body_html
            elif body_text:
                mail.Body = body_text

        # Set subject
        mail.Subject = subject

        # Set recipients
        for addr in to:
            mail.Recipients.Add(addr).Type = 1  # olTo
        for addr in cc:
            mail.Recipients.Add(addr).Type = 2  # olCC
        for addr in bcc:
            mail.Recipients.Add(addr).Type = 3  # olBCC

        # Resolve recipients
        mail.Recipients.ResolveAll()

        # Add attachments
        for att_path in attachments:
            path = Path(att_path)
            if not path.exists():
                raise OutlookError(f"Attachment not found: {att_path}")
            mail.Attachments.Add(str(path.absolute()))

        # Save as draft
        mail.Save()

        return mail.EntryID, mail.Parent.StoreID

    except Exception as e:
        raise OutlookError(f"Failed to create draft: {e}")


def send_draft(outlook_app, entry_id: str, store_id: str) -> None:
    """
    Send an existing draft.

    Args:
        outlook_app: Outlook Application COM object
        entry_id: Draft entry ID
        store_id: Draft store ID

    Raises:
        OutlookError: If send fails
    """
    try:
        mail = get_message_by_id(outlook_app, entry_id, store_id)
        mail.Send()
    except Exception as e:
        raise OutlookError(f"Failed to send draft: {e}")


def send_new_message(
    outlook_app,
    to: list[str],
    cc: list[str] = None,
    bcc: list[str] = None,
    subject: str = "",
    body_text: str = None,
    body_html: str = None,
    attachments: list[str] = None,
) -> None:
    """
    Create and immediately send a new message.

    Args:
        outlook_app: Outlook Application COM object
        to: List of To recipients
        cc: List of CC recipients
        bcc: List of BCC recipients
        subject: Email subject
        body_text: Plain text body
        body_html: HTML body
        attachments: List of file paths to attach

    Raises:
        OutlookError: If send fails
    """
    cc = cc or []
    bcc = bcc or []
    attachments = attachments or []

    try:
        mail = outlook_app.CreateItem(0)  # olMailItem

        # Set subject and body
        mail.Subject = subject
        if body_html:
            mail.HTMLBody = body_html
        elif body_text:
            mail.Body = body_text

        # Set recipients
        for addr in to:
            mail.Recipients.Add(addr).Type = 1
        for addr in cc:
            mail.Recipients.Add(addr).Type = 2
        for addr in bcc:
            mail.Recipients.Add(addr).Type = 3

        mail.Recipients.ResolveAll()

        # Add attachments
        for att_path in attachments:
            path = Path(att_path)
            if not path.exists():
                raise OutlookError(f"Attachment not found: {att_path}")
            mail.Attachments.Add(str(path.absolute()))

        # Send
        mail.Send()

    except Exception as e:
        raise OutlookError(f"Failed to send message: {e}")


def save_attachments(
    outlook_app,
    entry_id: str,
    store_id: str,
    dest_dir: str,
) -> list[str]:
    """
    Save attachments from a message to disk.

    Args:
        outlook_app: Outlook Application COM object
        entry_id: Message entry ID
        store_id: Message store ID
        dest_dir: Destination directory

    Returns:
        List of saved file paths

    Raises:
        OutlookError: If save fails
    """
    dest_path = Path(dest_dir)
    dest_path.mkdir(parents=True, exist_ok=True)

    mail = get_message_by_id(outlook_app, entry_id, store_id)
    saved_files = []

    for i in range(1, mail.Attachments.Count + 1):
        att = mail.Attachments.Item(i)
        filename = str(att.FileName)

        # Sanitize filename
        safe_name = "".join(c for c in filename if c.isalnum() or c in "._- ")
        if not safe_name:
            safe_name = f"attachment_{i}"

        # Handle duplicates
        save_path = dest_path / safe_name
        counter = 1
        while save_path.exists():
            stem = Path(safe_name).stem
            suffix = Path(safe_name).suffix
            save_path = dest_path / f"{stem}_{counter}{suffix}"
            counter += 1

        att.SaveAsFile(str(save_path))
        saved_files.append(str(save_path))

    return saved_files


def run_doctor() -> DoctorResult:
    """
    Run diagnostic checks on the environment.

    Returns:
        DoctorResult with all check results
    """
    import platform

    checks = []
    all_passed = True

    # Check 1: OS is Windows
    is_windows = platform.system() == "Windows"
    checks.append(DoctorCheck(
        name="windows_os",
        passed=is_windows,
        message="Windows OS detected" if is_windows else f"Not Windows: {platform.system()}",
        remediation=None if is_windows else "This tool requires Windows with Classic Outlook.",
    ))
    if not is_windows:
        all_passed = False

    # Check 2: pywin32 available
    try:
        import win32com.client
        checks.append(DoctorCheck(
            name="pywin32",
            passed=True,
            message="pywin32 is installed and importable",
        ))
    except ImportError:
        checks.append(DoctorCheck(
            name="pywin32",
            passed=False,
            message="pywin32 is not installed",
            remediation="Run: uv add pywin32",
        ))
        all_passed = False

    # Check 3: Outlook COM available
    outlook_path = None
    try:
        outlook = get_outlook_app(retry_count=1, retry_delay=0.5)
        _ = outlook.Name
        checks.append(DoctorCheck(
            name="outlook_com",
            passed=True,
            message="Outlook COM automation is available",
        ))
    except OutlookNotAvailableError as e:
        checks.append(DoctorCheck(
            name="outlook_com",
            passed=False,
            message=str(e),
            remediation="Ensure Classic Outlook is running. New Outlook does not support COM automation.",
        ))
        all_passed = False
    except Exception as e:
        checks.append(DoctorCheck(
            name="outlook_com",
            passed=False,
            message=f"Outlook COM check failed: {e}",
            remediation="Ensure Classic Outlook is installed and running.",
        ))
        all_passed = False

    # Check 4: Find Outlook executable
    outlook_path = find_outlook_executable()
    if outlook_path:
        checks.append(DoctorCheck(
            name="outlook_exe",
            passed=True,
            message=f"Outlook executable found: {outlook_path}",
        ))
    else:
        checks.append(DoctorCheck(
            name="outlook_exe",
            passed=False,
            message="Outlook executable not found in common paths",
            remediation="Outlook may be installed in a non-standard location.",
        ))
        # This is a warning, not a failure
        # all_passed = False

    return DoctorResult(
        all_passed=all_passed,
        checks=checks,
        outlook_path=outlook_path,
    )
