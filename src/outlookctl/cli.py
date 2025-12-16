"""
Command-line interface for outlookctl.

This module provides the main entry point and argument parsing for the CLI.
"""

import argparse
import json
import sys
from datetime import datetime
from typing import Optional

from . import __version__
from .models import (
    ListResult,
    SearchResult,
    DraftResult,
    SendResult,
    AttachmentSaveResult,
    ErrorResult,
    FolderInfo,
    MessageId,
)
from .outlook_com import (
    get_outlook_app,
    resolve_folder,
    list_messages,
    search_messages,
    get_message_by_id,
    extract_message_detail,
    create_draft,
    send_draft,
    send_new_message,
    save_attachments,
    run_doctor,
    OutlookError,
    OutlookNotAvailableError,
    FolderNotFoundError,
    MessageNotFoundError,
)
from .safety import (
    validate_send_confirmation,
    validate_unsafe_send_new,
    check_recipients,
    SendConfirmationError,
)
from .audit import log_send_operation, log_draft_operation


def output_json(data: dict, output_format: str = "json") -> None:
    """Output data in the specified format."""
    if output_format == "json":
        print(json.dumps(data, indent=2, ensure_ascii=False))
    else:
        # Simple text format
        print(json.dumps(data, indent=2, ensure_ascii=False))


def output_error(error: str, error_code: str = None, remediation: str = None) -> None:
    """Output an error in JSON format."""
    result = ErrorResult(
        error=error,
        error_code=error_code,
        remediation=remediation,
    )
    print(json.dumps(result.to_dict(), indent=2, ensure_ascii=False))
    sys.exit(1)


def parse_date(date_str: str) -> Optional[datetime]:
    """Parse an ISO date string."""
    if not date_str:
        return None
    try:
        return datetime.fromisoformat(date_str)
    except ValueError:
        try:
            return datetime.strptime(date_str, "%Y-%m-%d")
        except ValueError:
            raise ValueError(f"Invalid date format: {date_str}. Use ISO format (YYYY-MM-DD).")


def cmd_doctor(args: argparse.Namespace) -> None:
    """Run diagnostic checks."""
    result = run_doctor()
    output_json(result.to_dict(), args.output)
    if not result.all_passed:
        sys.exit(1)


def cmd_list(args: argparse.Namespace) -> None:
    """List messages from a folder."""
    try:
        outlook = get_outlook_app()
        folder, folder_info = resolve_folder(outlook, args.folder)

        since = parse_date(args.since) if args.since else None
        until = parse_date(args.until) if args.until else None

        messages = list(list_messages(
            outlook,
            folder_spec=args.folder,
            count=args.count,
            unread_only=args.unread_only,
            since=since,
            until=until,
            include_body_snippet=args.include_body_snippet,
            body_snippet_chars=args.body_snippet_chars,
        ))

        result = ListResult(
            folder=folder_info,
            items=messages,
        )
        output_json(result.to_dict(), args.output)

    except OutlookNotAvailableError as e:
        output_error(str(e), "OUTLOOK_UNAVAILABLE", "Start Classic Outlook and try again.")
    except FolderNotFoundError as e:
        output_error(str(e), "FOLDER_NOT_FOUND")
    except Exception as e:
        output_error(str(e), "LIST_ERROR")


def cmd_get(args: argparse.Namespace) -> None:
    """Get a single message by ID."""
    try:
        outlook = get_outlook_app()
        mail = get_message_by_id(outlook, args.id, args.store)

        detail = extract_message_detail(
            mail,
            include_body=args.include_body,
            max_body_chars=args.max_body_chars,
            include_headers=args.include_headers,
        )

        output_json({"version": "1.0", **detail.to_dict()}, args.output)

    except OutlookNotAvailableError as e:
        output_error(str(e), "OUTLOOK_UNAVAILABLE", "Start Classic Outlook and try again.")
    except MessageNotFoundError as e:
        output_error(str(e), "MESSAGE_NOT_FOUND")
    except Exception as e:
        output_error(str(e), "GET_ERROR")


def cmd_search(args: argparse.Namespace) -> None:
    """Search messages."""
    try:
        outlook = get_outlook_app()

        since = parse_date(args.since) if args.since else None
        until = parse_date(args.until) if args.until else None

        messages = list(search_messages(
            outlook,
            folder_spec=args.folder,
            query=args.query,
            from_filter=getattr(args, "from", None),
            subject_contains=args.subject_contains,
            unread_only=args.unread_only,
            since=since,
            until=until,
            count=args.count,
            include_body_snippet=args.include_body_snippet,
            body_snippet_chars=args.body_snippet_chars,
        ))

        query_info = {}
        if args.query:
            query_info["text"] = args.query
        if getattr(args, "from", None):
            query_info["from"] = getattr(args, "from")
        if args.subject_contains:
            query_info["subject_contains"] = args.subject_contains
        if args.unread_only:
            query_info["unread_only"] = True
        if since:
            query_info["since"] = since.isoformat()
        if until:
            query_info["until"] = until.isoformat()

        result = SearchResult(
            query=query_info,
            items=messages,
        )
        output_json(result.to_dict(), args.output)

    except OutlookNotAvailableError as e:
        output_error(str(e), "OUTLOOK_UNAVAILABLE", "Start Classic Outlook and try again.")
    except FolderNotFoundError as e:
        output_error(str(e), "FOLDER_NOT_FOUND")
    except Exception as e:
        output_error(str(e), "SEARCH_ERROR")


def cmd_draft(args: argparse.Namespace) -> None:
    """Create a draft message."""
    try:
        to_list = [addr.strip() for addr in args.to.split(",")] if args.to else []
        cc_list = [addr.strip() for addr in args.cc.split(",")] if args.cc else []
        bcc_list = [addr.strip() for addr in args.bcc.split(",")] if args.bcc else []

        check_recipients(to_list, cc_list, bcc_list)

        outlook = get_outlook_app()

        entry_id, store_id = create_draft(
            outlook,
            to=to_list,
            cc=cc_list,
            bcc=bcc_list,
            subject=args.subject or "",
            body_text=args.body_text,
            body_html=args.body_html,
            attachments=args.attach or [],
            reply_to_entry_id=args.reply_to_id,
            reply_to_store_id=args.reply_to_store,
        )

        log_draft_operation(
            to=to_list,
            cc=cc_list,
            bcc=bcc_list,
            subject=args.subject or "",
            success=True,
            entry_id=entry_id,
        )

        result = DraftResult(
            success=True,
            id=MessageId(entry_id=entry_id, store_id=store_id),
            saved_to="Drafts",
            subject=args.subject,
            to=to_list,
            cc=cc_list,
            attachments=args.attach or [],
        )
        output_json(result.to_dict(), args.output)

    except OutlookNotAvailableError as e:
        output_error(str(e), "OUTLOOK_UNAVAILABLE", "Start Classic Outlook and try again.")
    except ValueError as e:
        output_error(str(e), "VALIDATION_ERROR")
    except OutlookError as e:
        log_draft_operation(
            to=to_list if "to_list" in dir() else [],
            cc=cc_list if "cc_list" in dir() else [],
            bcc=bcc_list if "bcc_list" in dir() else [],
            subject=args.subject or "",
            success=False,
            error=str(e),
        )
        output_error(str(e), "DRAFT_ERROR")
    except Exception as e:
        output_error(str(e), "DRAFT_ERROR")


def cmd_send(args: argparse.Namespace) -> None:
    """Send a draft or new message."""
    try:
        outlook = get_outlook_app()

        # Case 1: Sending an existing draft
        if args.draft_id and args.draft_store:
            validate_send_confirmation(args.confirm_send, args.confirm_send_file)

            # Get draft info for logging
            mail = get_message_by_id(outlook, args.draft_id, args.draft_store)
            to_list = [str(r.Address) for r in mail.Recipients if r.Type == 1]
            subject = str(mail.Subject)

            send_draft(outlook, args.draft_id, args.draft_store)

            log_send_operation(
                to=to_list,
                cc=[],
                bcc=[],
                subject=subject,
                success=True,
                entry_id=args.draft_id,
                log_body=args.log_body,
            )

            result = SendResult(
                success=True,
                message="Draft sent successfully",
                sent_at=datetime.now().isoformat(),
                to=to_list,
                subject=subject,
            )
            output_json(result.to_dict(), args.output)

        # Case 2: Sending a new message directly (requires --unsafe-send-new)
        elif args.to:
            validate_unsafe_send_new(
                args.unsafe_send_new,
                args.confirm_send,
                args.confirm_send_file,
            )

            to_list = [addr.strip() for addr in args.to.split(",")]
            cc_list = [addr.strip() for addr in args.cc.split(",")] if args.cc else []
            bcc_list = [addr.strip() for addr in args.bcc.split(",")] if args.bcc else []

            check_recipients(to_list, cc_list, bcc_list)

            send_new_message(
                outlook,
                to=to_list,
                cc=cc_list,
                bcc=bcc_list,
                subject=args.subject or "",
                body_text=args.body_text,
                body_html=args.body_html,
                attachments=args.attach or [],
            )

            log_send_operation(
                to=to_list,
                cc=cc_list,
                bcc=bcc_list,
                subject=args.subject or "",
                success=True,
                log_body=args.log_body,
                body=args.body_text or args.body_html,
            )

            result = SendResult(
                success=True,
                message="Message sent successfully",
                sent_at=datetime.now().isoformat(),
                to=to_list,
                subject=args.subject,
            )
            output_json(result.to_dict(), args.output)

        else:
            output_error(
                "Either --draft-id/--draft-store or --to is required",
                "MISSING_ARGUMENTS",
                "Use --draft-id and --draft-store to send an existing draft, "
                "or use --to with --unsafe-send-new to send a new message directly.",
            )

    except SendConfirmationError as e:
        output_error(str(e), "CONFIRMATION_REQUIRED")
    except OutlookNotAvailableError as e:
        output_error(str(e), "OUTLOOK_UNAVAILABLE", "Start Classic Outlook and try again.")
    except MessageNotFoundError as e:
        output_error(str(e), "DRAFT_NOT_FOUND")
    except ValueError as e:
        output_error(str(e), "VALIDATION_ERROR")
    except OutlookError as e:
        log_send_operation(
            to=[],
            cc=[],
            bcc=[],
            subject="",
            success=False,
            error=str(e),
        )
        output_error(str(e), "SEND_ERROR")
    except Exception as e:
        output_error(str(e), "SEND_ERROR")


def cmd_attachments_save(args: argparse.Namespace) -> None:
    """Save attachments from a message."""
    try:
        outlook = get_outlook_app()

        saved_files = save_attachments(
            outlook,
            entry_id=args.id,
            store_id=args.store,
            dest_dir=args.dest,
        )

        result = AttachmentSaveResult(
            success=True,
            saved_files=saved_files,
        )
        output_json(result.to_dict(), args.output)

    except OutlookNotAvailableError as e:
        output_error(str(e), "OUTLOOK_UNAVAILABLE", "Start Classic Outlook and try again.")
    except MessageNotFoundError as e:
        output_error(str(e), "MESSAGE_NOT_FOUND")
    except OutlookError as e:
        output_error(str(e), "ATTACHMENT_ERROR")
    except Exception as e:
        output_error(str(e), "ATTACHMENT_ERROR")


def create_parser() -> argparse.ArgumentParser:
    """Create the argument parser."""
    parser = argparse.ArgumentParser(
        prog="outlookctl",
        description="Local CLI bridge for Outlook Classic automation via COM",
    )
    parser.add_argument(
        "--version", action="version", version=f"%(prog)s {__version__}"
    )

    subparsers = parser.add_subparsers(dest="command", help="Available commands")

    # Doctor command
    doctor_parser = subparsers.add_parser(
        "doctor", help="Validate environment and prerequisites"
    )
    doctor_parser.add_argument(
        "--output", choices=["json", "text"], default="json",
        help="Output format (default: json)"
    )
    doctor_parser.set_defaults(func=cmd_doctor)

    # List command
    list_parser = subparsers.add_parser(
        "list", help="List messages from a folder"
    )
    list_parser.add_argument(
        "--folder", default="inbox",
        help="Folder: inbox|sent|drafts|by-name:<name>|by-path:<path> (default: inbox)"
    )
    list_parser.add_argument(
        "--count", type=int, default=10,
        help="Number of messages to return (default: 10)"
    )
    list_parser.add_argument(
        "--unread-only", action="store_true",
        help="Only return unread messages"
    )
    list_parser.add_argument(
        "--since", help="Only messages received after this date (ISO format)"
    )
    list_parser.add_argument(
        "--until", help="Only messages received before this date (ISO format)"
    )
    list_parser.add_argument(
        "--include-body-snippet", action="store_true",
        help="Include a snippet of the message body"
    )
    list_parser.add_argument(
        "--body-snippet-chars", type=int, default=200,
        help="Maximum characters for body snippet (default: 200)"
    )
    list_parser.add_argument(
        "--output", choices=["json", "text"], default="json",
        help="Output format (default: json)"
    )
    list_parser.set_defaults(func=cmd_list)

    # Get command
    get_parser = subparsers.add_parser(
        "get", help="Get a single message by ID"
    )
    get_parser.add_argument(
        "--id", required=True, help="Message entry ID"
    )
    get_parser.add_argument(
        "--store", required=True, help="Message store ID"
    )
    get_parser.add_argument(
        "--include-body", action="store_true",
        help="Include message body"
    )
    get_parser.add_argument(
        "--include-headers", action="store_true",
        help="Include message headers"
    )
    get_parser.add_argument(
        "--max-body-chars", type=int,
        help="Maximum characters for body"
    )
    get_parser.add_argument(
        "--output", choices=["json", "text"], default="json",
        help="Output format (default: json)"
    )
    get_parser.set_defaults(func=cmd_get)

    # Search command
    search_parser = subparsers.add_parser(
        "search", help="Search messages"
    )
    search_parser.add_argument(
        "--folder", default="inbox",
        help="Folder to search in (default: inbox)"
    )
    search_parser.add_argument(
        "--query", help="Free text search (subject/body)"
    )
    search_parser.add_argument(
        "--from", dest="from", help="Filter by sender email or name"
    )
    search_parser.add_argument(
        "--subject-contains", help="Filter by subject content"
    )
    search_parser.add_argument(
        "--unread-only", action="store_true",
        help="Only return unread messages"
    )
    search_parser.add_argument(
        "--since", help="Only messages after this date (ISO format)"
    )
    search_parser.add_argument(
        "--until", help="Only messages before this date (ISO format)"
    )
    search_parser.add_argument(
        "--count", type=int, default=50,
        help="Maximum results (default: 50)"
    )
    search_parser.add_argument(
        "--include-body-snippet", action="store_true",
        help="Include a snippet of the message body"
    )
    search_parser.add_argument(
        "--body-snippet-chars", type=int, default=200,
        help="Maximum characters for body snippet (default: 200)"
    )
    search_parser.add_argument(
        "--output", choices=["json", "text"], default="json",
        help="Output format (default: json)"
    )
    search_parser.set_defaults(func=cmd_search)

    # Draft command
    draft_parser = subparsers.add_parser(
        "draft", help="Create a draft message"
    )
    draft_parser.add_argument(
        "--to", help="To recipients (comma-separated)"
    )
    draft_parser.add_argument(
        "--cc", help="CC recipients (comma-separated)"
    )
    draft_parser.add_argument(
        "--bcc", help="BCC recipients (comma-separated)"
    )
    draft_parser.add_argument(
        "--subject", help="Email subject"
    )
    draft_parser.add_argument(
        "--body-text", help="Plain text body"
    )
    draft_parser.add_argument(
        "--body-html", help="HTML body"
    )
    draft_parser.add_argument(
        "--attach", action="append",
        help="File path to attach (can be used multiple times)"
    )
    draft_parser.add_argument(
        "--reply-to-id", help="Entry ID of message to reply to"
    )
    draft_parser.add_argument(
        "--reply-to-store", help="Store ID of message to reply to"
    )
    draft_parser.add_argument(
        "--output", choices=["json", "text"], default="json",
        help="Output format (default: json)"
    )
    draft_parser.set_defaults(func=cmd_draft)

    # Send command
    send_parser = subparsers.add_parser(
        "send", help="Send a draft or new message"
    )
    # For sending existing draft
    send_parser.add_argument(
        "--draft-id", help="Entry ID of draft to send"
    )
    send_parser.add_argument(
        "--draft-store", help="Store ID of draft to send"
    )
    # For sending new message directly
    send_parser.add_argument(
        "--to", help="To recipients (comma-separated)"
    )
    send_parser.add_argument(
        "--cc", help="CC recipients (comma-separated)"
    )
    send_parser.add_argument(
        "--bcc", help="BCC recipients (comma-separated)"
    )
    send_parser.add_argument(
        "--subject", help="Email subject"
    )
    send_parser.add_argument(
        "--body-text", help="Plain text body"
    )
    send_parser.add_argument(
        "--body-html", help="HTML body"
    )
    send_parser.add_argument(
        "--attach", action="append",
        help="File path to attach (can be used multiple times)"
    )
    # Safety flags
    send_parser.add_argument(
        "--confirm-send",
        help="Confirmation string (must be exactly 'YES')"
    )
    send_parser.add_argument(
        "--confirm-send-file",
        help="Path to file containing confirmation string"
    )
    send_parser.add_argument(
        "--unsafe-send-new", action="store_true",
        help="Allow sending new message directly (not recommended)"
    )
    send_parser.add_argument(
        "--log-body", action="store_true",
        help="Include body in audit log (default: metadata only)"
    )
    send_parser.add_argument(
        "--output", choices=["json", "text"], default="json",
        help="Output format (default: json)"
    )
    send_parser.set_defaults(func=cmd_send)

    # Attachments subcommand
    attachments_parser = subparsers.add_parser(
        "attachments", help="Attachment operations"
    )
    attachments_subparsers = attachments_parser.add_subparsers(
        dest="attachments_command", help="Attachment commands"
    )

    # Attachments save
    save_parser = attachments_subparsers.add_parser(
        "save", help="Save attachments from a message"
    )
    save_parser.add_argument(
        "--id", required=True, help="Message entry ID"
    )
    save_parser.add_argument(
        "--store", required=True, help="Message store ID"
    )
    save_parser.add_argument(
        "--dest", required=True, help="Destination directory"
    )
    save_parser.add_argument(
        "--output", choices=["json", "text"], default="json",
        help="Output format (default: json)"
    )
    save_parser.set_defaults(func=cmd_attachments_save)

    return parser


def main() -> None:
    """Main entry point."""
    parser = create_parser()
    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        sys.exit(1)

    # Handle attachments subcommand
    if args.command == "attachments" and not args.attachments_command:
        parser.parse_args(["attachments", "-h"])
        sys.exit(1)

    if hasattr(args, "func"):
        args.func(args)
    else:
        parser.print_help()
        sys.exit(1)


if __name__ == "__main__":
    main()
