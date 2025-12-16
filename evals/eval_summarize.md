# Evaluation: Summarize Unread Emails

## Scenario

**User prompt:**
> "Summarize my latest 10 unread emails from my Inbox."

## Expected Behavior

1. **Initial action**: Claude should run:
   ```bash
   uv run outlookctl list --folder inbox --count 10 --unread-only
   ```

2. **Processing**: Claude should summarize the emails based on metadata:
   - Subject lines
   - Sender names/emails
   - Received timestamps
   - Attachment indicators

3. **No body fetching**: Claude should NOT automatically fetch email bodies unless the user explicitly asks.

4. **Follow-up handling**: If user asks "What does the email from X say?", Claude should then run:
   ```bash
   uv run outlookctl get --id "<entry_id>" --store "<store_id>" --include-body
   ```

## Evaluation Criteria

| Criterion | Pass | Fail |
|-----------|------|------|
| Uses `outlookctl list` | Runs list command | Uses other methods |
| Includes `--unread-only` | Flag present | Missing flag |
| Respects count limit | Uses `--count 10` | Different count |
| Summarizes metadata | Shows subjects, senders | Skips information |
| Doesn't auto-fetch bodies | No `--include-body` initially | Fetches bodies immediately |
| Handles empty results | Reports "no unread emails" | Errors or confusion |

## Success Criteria

- [ ] Correct CLI command executed
- [ ] Summary based on metadata only
- [ ] Asks before fetching full content
- [ ] Clear, organized output
