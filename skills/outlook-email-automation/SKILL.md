---
name: outlook-email-automation
description: >
  Automates reading, searching, drafting, and sending emails in Classic Outlook
  on Windows using local COM automation. Use this skill when the user asks to
  process Outlook emails, create drafts, send messages, save attachments, or
  interact with their Inbox from the authenticated Outlook session on their
  Windows Devbox. Requires Classic Outlook (not New Outlook) to be running.
---

# Outlook Email Automation

This skill enables email automation through Classic Outlook on Windows using the `outlookctl` CLI tool.

## How to Run Commands

Run all commands using this pattern from any directory:

```bash
uv run --project "C:/Users/GordonMickel/work/outlookctl" python -m outlookctl.cli <command> [options]
```

For convenience, define this alias at the start of your session:
```bash
alias outlookctl='uv run --project "C:/Users/GordonMickel/work/outlookctl" python -m outlookctl.cli'
```

## Quick Start

Before using any commands, verify the environment:

```bash
uv run --project "C:/Users/GordonMickel/work/outlookctl" python -m outlookctl.cli doctor
```

## Available Commands

| Command | Description |
|---------|-------------|
| `doctor` | Validate environment and prerequisites |
| `list` | List messages from a folder |
| `get` | Get a single message by ID |
| `search` | Search messages with filters |
| `draft` | Create a draft message |
| `send` | Send a draft or new message |
| `attachments save` | Save attachments to disk |

## Safety Rules

**CRITICAL: Follow these rules when handling email operations:**

1. **Never auto-send emails** - Always create drafts first and get explicit user confirmation before sending
2. **Draft-first workflow** - Use `draft` to create drafts, show the user a preview, then send only after approval
3. **Explicit confirmation required** - The send command requires `--confirm-send YES` flag
4. **Metadata by default** - Body content is only retrieved when explicitly requested

## Workflows

### Reading and Searching Email

To list recent emails:
```bash
uv run --project "C:/Users/GordonMickel/work/outlookctl" python -m outlookctl.cli list --count 10
```

To search for specific emails:
```bash
uv run --project "C:/Users/GordonMickel/work/outlookctl" python -m outlookctl.cli search --from "sender@example.com" --since 2025-01-01
```

To get full message content (only when user asks):
```bash
uv run --project "C:/Users/GordonMickel/work/outlookctl" python -m outlookctl.cli get --id "<entry_id>" --store "<store_id>" --include-body
```

### Creating and Sending Email (Draft-First)

**Step 1: Create a draft**
```bash
uv run --project "C:/Users/GordonMickel/work/outlookctl" python -m outlookctl.cli draft --to "recipient@example.com" --subject "Subject" --body-text "Message body"
```

**Step 2: Show user the preview** (subject, recipients, body summary)

**Step 3: Only after user confirms, send the draft**
```bash
uv run --project "C:/Users/GordonMickel/work/outlookctl" python -m outlookctl.cli send --draft-id "<entry_id>" --draft-store "<store_id>" --confirm-send YES
```

### Replying to Messages

```bash
# Create reply draft
uv run --project "C:/Users/GordonMickel/work/outlookctl" python -m outlookctl.cli draft --to "recipient@example.com" --subject "Re: Original" --body-text "Reply text" --reply-to-id "<entry_id>" --reply-to-store "<store_id>"
```

### Saving Attachments

```bash
uv run --project "C:/Users/GordonMickel/work/outlookctl" python -m outlookctl.cli attachments save --id "<entry_id>" --store "<store_id>" --dest "./attachments"
```

## Output Format

All commands output JSON with a consistent structure. Key fields:

- `version`: Schema version (currently "1.0")
- `success`: Boolean for operation result
- Message IDs include `entry_id` and `store_id` for stable references

## Reference Documentation

For detailed information, see:
- [CLI Reference](reference/cli.md) - Complete command options
- [JSON Schema](reference/json-schema.md) - Output format details
- [Security](reference/security.md) - Data handling and safety
- [Troubleshooting](reference/troubleshooting.md) - Common issues

## Requirements

- Windows with Classic Outlook running
- uv installed and in PATH
- outlookctl project at: `C:/Users/GordonMickel/work/outlookctl`

## Error Handling

If commands fail, check:
1. Classic Outlook is running (not New Outlook)
2. `doctor` command passes all checks
3. Message IDs are valid and not expired
