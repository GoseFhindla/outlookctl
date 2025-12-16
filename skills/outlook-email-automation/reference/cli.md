# CLI Reference

Complete reference for all `outlookctl` commands and options.

## Global Options

All commands support:
- `--output json|text` - Output format (default: json)

## Commands

### `outlookctl doctor`

Validates environment and prerequisites.

```bash
uv run outlookctl doctor
```

**Checks performed:**
- Windows OS detection
- pywin32 installation
- Outlook COM availability
- Outlook executable location

**Example output:**
```json
{
  "version": "1.0",
  "all_passed": true,
  "checks": [
    {"name": "windows_os", "passed": true, "message": "Windows OS detected"},
    {"name": "pywin32", "passed": true, "message": "pywin32 is installed"},
    {"name": "outlook_com", "passed": true, "message": "Outlook COM available"},
    {"name": "outlook_exe", "passed": true, "message": "Found: C:\\...\\OUTLOOK.EXE"}
  ],
  "outlook_path": "C:\\Program Files\\Microsoft Office\\root\\Office16\\OUTLOOK.EXE"
}
```

---

### `outlookctl list`

List messages from a folder.

```bash
uv run outlookctl list [OPTIONS]
```

**Options:**

| Option | Default | Description |
|--------|---------|-------------|
| `--folder` | inbox | Folder specification (see below) |
| `--count` | 10 | Number of messages to return |
| `--unread-only` | false | Only return unread messages |
| `--since` | - | ISO date filter (messages after) |
| `--until` | - | ISO date filter (messages before) |
| `--include-body-snippet` | false | Include body preview |
| `--body-snippet-chars` | 200 | Max chars for snippet |

**Folder specifications:**
- `inbox` - Default inbox
- `sent` - Sent items
- `drafts` - Drafts folder
- `deleted` - Deleted items
- `outbox` - Outbox
- `junk` - Junk/spam
- `by-name:<name>` - Find folder by name
- `by-path:<path>` - Find folder by path (e.g., `Inbox/Subfolder`)

**Example:**
```bash
uv run outlookctl list --folder inbox --count 5 --unread-only --include-body-snippet
```

---

### `outlookctl get`

Get a single message by ID.

```bash
uv run outlookctl get --id <entry_id> --store <store_id> [OPTIONS]
```

**Required:**
- `--id` - Message entry ID
- `--store` - Message store ID

**Options:**

| Option | Default | Description |
|--------|---------|-------------|
| `--include-body` | false | Include full message body |
| `--include-headers` | false | Include message headers |
| `--max-body-chars` | - | Limit body length |

**Example:**
```bash
uv run outlookctl get --id "00000..." --store "00000..." --include-body --max-body-chars 5000
```

---

### `outlookctl search`

Search messages with various filters.

```bash
uv run outlookctl search [OPTIONS]
```

**Options:**

| Option | Default | Description |
|--------|---------|-------------|
| `--folder` | inbox | Folder to search |
| `--query` | - | Free text search (subject/body) |
| `--from` | - | Filter by sender |
| `--subject-contains` | - | Filter by subject text |
| `--unread-only` | false | Only unread messages |
| `--since` | - | ISO date filter |
| `--until` | - | ISO date filter |
| `--count` | 50 | Maximum results |
| `--include-body-snippet` | false | Include body preview |
| `--body-snippet-chars` | 200 | Max chars for snippet |

**Example:**
```bash
uv run outlookctl search --from "boss@company.com" --since 2025-01-01 --unread-only
```

---

### `outlookctl draft`

Create a draft message.

```bash
uv run outlookctl draft [OPTIONS]
```

**Options:**

| Option | Description |
|--------|-------------|
| `--to` | To recipients (comma-separated) |
| `--cc` | CC recipients (comma-separated) |
| `--bcc` | BCC recipients (comma-separated) |
| `--subject` | Email subject |
| `--body-text` | Plain text body |
| `--body-html` | HTML body (mutually exclusive with --body-text) |
| `--attach` | File path to attach (repeatable) |
| `--reply-to-id` | Entry ID of message to reply to |
| `--reply-to-store` | Store ID of message to reply to |

**Example:**
```bash
uv run outlookctl draft \
  --to "recipient@example.com" \
  --cc "cc@example.com" \
  --subject "Meeting Follow-up" \
  --body-text "Thank you for the meeting today." \
  --attach "./report.pdf"
```

**Reply example:**
```bash
uv run outlookctl draft \
  --to "sender@example.com" \
  --subject "Re: Original Subject" \
  --body-text "Reply content" \
  --reply-to-id "00000..." \
  --reply-to-store "00000..."
```

---

### `outlookctl send`

Send a draft or new message. **Requires explicit confirmation.**

#### Sending an existing draft (recommended):

```bash
uv run outlookctl send \
  --draft-id <entry_id> \
  --draft-store <store_id> \
  --confirm-send YES
```

#### Sending a new message directly (not recommended):

```bash
uv run outlookctl send \
  --to "recipient@example.com" \
  --subject "Subject" \
  --body-text "Body" \
  --unsafe-send-new \
  --confirm-send YES
```

**Safety options:**

| Option | Description |
|--------|-------------|
| `--confirm-send` | Must be exactly "YES" to proceed |
| `--confirm-send-file` | Path to file containing "YES" |
| `--unsafe-send-new` | Required flag for sending new message directly |
| `--log-body` | Include body in audit log |

---

### `outlookctl attachments save`

Save attachments from a message to disk.

```bash
uv run outlookctl attachments save \
  --id <entry_id> \
  --store <store_id> \
  --dest <directory>
```

**Required:**
- `--id` - Message entry ID
- `--store` - Message store ID
- `--dest` - Destination directory (created if needed)

**Example:**
```bash
uv run outlookctl attachments save --id "00000..." --store "00000..." --dest "./downloads"
```

**Output:**
```json
{
  "version": "1.0",
  "success": true,
  "saved_files": [
    "./downloads/document.pdf",
    "./downloads/image.png"
  ],
  "errors": []
}
```
