# Stable Working State - 2026-04-15

## Status

This project state is considered a working baseline as of `2026-04-15`.

It is not a `git` commit because the workspace is not currently a git repository, so the baseline is fixed through:

- this stable handoff record;
- the latest successful runtime artifacts;
- a local code snapshot archive stored under `snapshots/`:
  - `snapshots/automation-working-state-2026-04-15.zip`

## Verification

The following checks passed:

- full mailbox archive run completed successfully:
  - `logs/mail_cycle_20260414_214433.log`
  - `logs/mail_cycle_20260414_214433.stdout.log`
  - `logs/mail_cycle_20260414_214433.stderr.log`
  - `logs/mail_cycle_20260414_214433.exitcode.txt` = `0`
- Python compile check passed:
  - `services/mail_ingest/outlook_desktop_client.py`
  - `services/mail_ingest/outlook_sync.py`
  - `scripts/sync_outlook_mail.py`
- `.NET` build passed:
  - `dotnet build .\Automation.sln`
- local code snapshot archive created:
  - `snapshots/automation-working-state-2026-04-15.zip`

## Working Outcomes

- saved raw messages on disk: `19449`
- saved messages with attachments: `15368`
- saved messages missing `sourceFolderPath`: `0`
- saved messages missing `sourceFolderName`: `0`
- saved messages missing `sourceStoreName`: `0`
- saved attachment-bearing messages missing `attachments.json`: `0`
- mirrored folder structure rebuilt successfully:
  - `raw/mail/folders.json`
  - `raw/mail/by_folder/...`

## Important Note

This is a working baseline, but not a fully finished archive.

At the time this baseline was fixed, the remaining known gap was:

- `200` historical messages still not in `raw/`
- `17` pending attachments on those unread historical messages not yet in `raw/`

That means:

- the code path is working;
- the full-mailbox archive completed successfully after the attachment fallback fix;
- the remaining gap is now operational backlog, not a broken baseline.

## Resume Point

If work resumes from this stable state:

1. read `wiki/handoffs/current-session.md`
2. read this file
3. continue from the remaining `200` historical messages
4. only after that move deeper into downstream processing and UI work
