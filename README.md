# outlookctl

Local CLI bridge for Outlook Classic automation on Windows via COM. Includes a Claude Code Skill for AI-assisted email management.

## What This Is

`outlookctl` is a **local automation tool** that controls the Outlook desktop client already running on your Windows workstation. It:

- **Uses your existing Outlook session** - No separate authentication, API keys, or OAuth tokens
- **Operates entirely locally** - No network calls to external services, no cloud dependencies
- **Controls the desktop client via COM** - Same automation interface used by VBA macros and Office add-ins
- **Respects your security context** - Runs with your existing permissions, subject to Outlook's security settings

This is **not** a workaround or bypass - it's the standard Windows COM automation interface that Microsoft provides for programmatic Outlook access. The same technology powers countless enterprise tools, email archivers, and Office integrations.

## Use Cases

- **AI-assisted email triage** - Let Claude help summarize and categorize your inbox
- **Automated drafting** - Generate draft responses with AI assistance, review before sending
- **Email search and retrieval** - Find specific messages across your mailbox
- **Attachment management** - Bulk save attachments to disk
- **Workflow automation** - Script repetitive email tasks

## Requirements

- Windows workstation with Classic Outlook (not "New Outlook")
- Outlook running and logged into your account
- Python 3.12+ and [uv](https://docs.astral.sh/uv/)

## Quick Start

### 1. Clone and Setup

```bash
git clone <repo-url>
cd outlookctl
uv sync
```

### 2. Verify Environment

```bash
uv run python -m outlookctl.cli doctor
```

All checks should pass. If not, see [Troubleshooting](skills/outlook-email-automation/reference/troubleshooting.md).

### 3. Test Commands

```bash
# List recent emails
uv run python -m outlookctl.cli list --count 5

# Search emails
uv run python -m outlookctl.cli search --from "someone@example.com" --since 2025-01-01

# Create a draft
uv run python -m outlookctl.cli draft --to "recipient@example.com" --subject "Test" --body-text "Hello"
```

## CLI Commands

| Command | Description |
|---------|-------------|
| `doctor` | Validate environment and prerequisites |
| `list` | List messages from a folder |
| `get` | Get a single message by ID |
| `search` | Search messages with filters |
| `draft` | Create a draft message |
| `send` | Send a draft or new message |
| `attachments save` | Save attachments to disk |

See [CLI Reference](skills/outlook-email-automation/reference/cli.md) for full documentation.

## Installing the Claude Code Skill

The skill enables Claude to assist with email operations safely.

### Personal Installation

```bash
uv run python tools/install_skill.py --personal
```

Installs to: `~/.claude/skills/outlook-email-automation/`

### Project Installation (for team)

```bash
uv run python tools/install_skill.py --project
```

Installs to: `.claude/skills/outlook-email-automation/`

### Verify Installation

```bash
uv run python tools/install_skill.py --verify --personal
```

## Safety Features

`outlookctl` is designed with safety as a priority:

1. **Draft-First Workflow** - Create drafts, review, then send
2. **Explicit Confirmation** - Sending requires `--confirm-send YES`
3. **Metadata by Default** - Body content only retrieved on explicit request
4. **Audit Logging** - Send operations logged to `%LOCALAPPDATA%/outlookctl/audit.log`

### Example Safe Workflow

```bash
# 1. Create draft
uv run python -m outlookctl.cli draft \
  --to "recipient@example.com" \
  --subject "Project Update" \
  --body-text "Here is the update..."

# 2. Review the draft in Outlook or via CLI

# 3. Send with explicit confirmation
uv run python -m outlookctl.cli send \
  --draft-id "<entry_id from step 1>" \
  --draft-store "<store_id from step 1>" \
  --confirm-send YES
```

## Output Format

All commands output JSON:

```json
{
  "version": "1.0",
  "folder": {"name": "Inbox"},
  "items": [
    {
      "id": {"entry_id": "...", "store_id": "..."},
      "subject": "Meeting Tomorrow",
      "from": {"name": "Jane", "email": "jane@example.com"},
      "unread": true,
      "has_attachments": false
    }
  ]
}
```

See [JSON Schema](skills/outlook-email-automation/reference/json-schema.md) for details.

## Project Structure

```
outlookctl/
├── pyproject.toml              # Project configuration (uv/hatch)
├── README.md                   # This file
├── CLAUDE.md                   # Development guide for AI assistants
├── src/outlookctl/             # Python package
│   ├── __init__.py
│   ├── cli.py                  # CLI entry point (argparse)
│   ├── models.py               # Dataclasses for JSON output
│   ├── outlook_com.py          # COM automation wrapper
│   ├── safety.py               # Send confirmation gates
│   └── audit.py                # Audit logging
├── skills/
│   └── outlook-email-automation/
│       ├── SKILL.md            # Claude Code Skill definition
│       └── reference/          # Skill documentation
│           ├── cli.md
│           ├── json-schema.md
│           ├── security.md
│           └── troubleshooting.md
├── tools/
│   └── install_skill.py        # Skill installer
├── tests/                      # pytest test suite
│   ├── test_models.py
│   └── test_safety.py
└── evals/                      # Skill evaluation scenarios
    ├── eval_summarize.md
    ├── eval_draft_reply.md
    └── eval_refuse_send.md
```

## Development

### Prerequisites

- Python 3.12+
- uv package manager
- Windows with Classic Outlook

### Setup

```bash
uv sync
```

### Run Tests

```bash
uv run python -m pytest tests/ -v
```

### Run CLI During Development

```bash
uv run python -m outlookctl.cli <command> [options]
```

### Update Skill After Changes

```bash
uv run python tools/install_skill.py --personal
```

## Technical Details

### Why COM Automation?

Windows COM (Component Object Model) is Microsoft's standard interface for inter-process communication. Outlook exposes its functionality through the `Outlook.Application` COM object, which:

- Is the same interface used by VBA macros inside Outlook
- Is how enterprise tools integrate with Outlook
- Runs in the security context of the logged-in user
- Is subject to Outlook's Trust Center settings

### Classic vs New Outlook

**Classic Outlook** (the traditional desktop app) supports COM automation.

**New Outlook** (the modern, web-based app) does **not** support COM automation - it requires Microsoft Graph API with OAuth authentication.

This tool only works with Classic Outlook. Check which version you have:
- Classic: Has File menu, Trust Center settings
- New: Toggle at top-right says "New Outlook"

### Security Model

- No API keys or tokens stored
- No network calls to external services
- Uses Windows authentication (your logged-in session)
- Outlook may show security prompts for programmatic access
- All operations logged locally for audit

## Limitations

- **Classic Outlook Only** - New Outlook requires Microsoft Graph API
- **Windows Only** - COM is a Windows technology
- **Same Session** - Must run in same Windows session as Outlook
- **Security Prompts** - Outlook may show security dialogs

## Troubleshooting

| Issue | Solution |
|-------|----------|
| "Outlook COM unavailable" | Start Classic Outlook (not New Outlook) |
| "pywin32 not installed" | Run `uv sync` |
| "Message not found" | IDs expire; re-run list/search |
| Permission denied on CLI | Use `uv run python -m outlookctl.cli` instead |

See [Troubleshooting Guide](skills/outlook-email-automation/reference/troubleshooting.md) for more.

## License

MIT
