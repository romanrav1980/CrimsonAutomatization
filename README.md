# Automation Workspace

This repository is organized as a standard Visual Studio solution with an ASP.NET Core MVC portal and a preserved Python Outlook-ingestion module.

## Solution

- `Automation.sln` - Visual Studio solution
- `src/Automation.Web` - ASP.NET Core MVC portal in a dark crimson style
- `services/` - Python mail-ingestion logic
- `scripts/` - Python CLI entry points
- `raw/` - downloaded mail and operational source materials

## Subprojects

- `Outlook Mail Ingest`
  - local Outlook COM + `raw/mail` archive layer
  - see `services/mail_ingest/README.md`
  - operational note: `wiki/services/outlook-mail-ingest.md`

Open in Visual Studio:

1. Open `Automation.sln`.
2. Set `Automation.Web` as the startup project if needed.
3. Run with IIS Express or the project profile.

CLI build:

```powershell
dotnet build .\Automation.sln
```

CLI run:

```powershell
dotnet run --project .\src\Automation.Web\Automation.Web.csproj
```

## MVP mail cycle

The first end-to-end MVP flow is:

1. Sync Outlook mail into `raw/mail`
2. Catch up old unread historical mail into `raw/mail`
3. Normalize raw mail into SQLite
4. Analyze Excel, PDF, image, and other attachments into `derived/mail`
5. Classify by process type
6. Apply decision matrix
7. Show `Needs Decision` in the MVC portal

Process existing raw mail:

```powershell
python .\scripts\process_mail_mvp.py
```

Run one full cycle:

```powershell
python .\scripts\sync_outlook_mail.py --process-after-sync
```

Or use the PowerShell wrapper for scheduler-friendly execution:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\run_mail_cycle.ps1
```

Or use the visual monitor window:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\run_mail_cycle_ui.ps1
```

The visual monitor shows:

- current stage
- live stage activity with current-stage runtime and quiet-time indicator
- stage-by-stage status for `Sync`, `Catch-up`, `Process`, `Normalize`, `Attachments`, `Classify`, `Decision`, and `Backlog`
- a stage-based progress bar that keeps moving while the active stage is still running
- live counters for sync, historical unread catch-up, attachment download, attachment analysis, processed, failed, and `Needs Decision`
- elapsed time
- live log output
- periodic debug screenshots in `logs/screenshots/...`
- a stop button that terminates the child mail process

The `Needs Decision` screen now supports operator actions:

- approve suggestion
- mark manual review
- archive item
- assign owner with notes

Those actions are persisted through the Python mail-processing layer and refresh the dashboard read model.

The `Needs Decision` screen now also shows attachment analysis for each mail item:

- attachment kinds such as `excel`, `pdf`, and `image`
- workbook sheet counts and sheet names for supported Excel files
- PDF page counts and text preview when available
- image dimensions
- metadata-only fallback when deep parsing is not configured yet

The sync also writes a mail-ingest status snapshot to `data/mail_ingest_status.json` when Outlook COM is available. That snapshot is used to show historical backlog from previous years:

- unread historical messages
- historical messages not yet stored in `raw/`
- unread historical messages that still need to be ingested into `raw/`
- unread historical messages with attachments
- attachment count remaining on unread historical messages that are still not in `raw/`
- per-year backlog breakdown

The scheduled MVP now performs historical unread catch-up by default:

- only historical unread mail is downloaded
- already-read historical mail is skipped
- only messages that are not yet in `raw/` are downloaded
- attachments for those unread messages are downloaded together with the message
- the batch is incremental and idempotent, so repeated scheduled runs eventually drain the unread historical backlog

Register the interactive scheduled task:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\register_mail_task.ps1
```

The normalized SQLite database is stored in `data/automation.db`, and the dashboard read model is exported to `data/mail_triage_read_model.json`.

## Outlook sync

The primary MVP ingest now uses the local Outlook desktop profile through COM/MAPI and saves messages plus attachments into `raw/mail/`.

### Supported providers

- `desktop`: local Outlook profile on this Windows machine
- `graph`: Microsoft Graph for a future server-side version

### Recommended setup for this MVP

1. Copy `.env.example` to `.env`.
2. Fill in:
   - `OUTLOOK_PROVIDER=desktop`
   - `OUTLOOK_MAILBOX_NAME=r.rabayev@korzinka.uz`
3. Install dependencies:

```powershell
python -m pip install -r requirements.txt
```

Important:

- the desktop provider must run in an interactive logged-in Windows session;
- Outlook must already be configured for that user profile;
- in Windows Task Scheduler, use `Run only when user is logged on` for the desktop provider.
- the visual monitor kills the child mail process if you stop the run from the window, which helps prevent hidden hanging runs.

### Optional setup for the graph provider

If you later want the server-side path, set:

- `OUTLOOK_PROVIDER=graph`
- `OUTLOOK_TENANT_ID`
- `OUTLOOK_CLIENT_ID`
- `OUTLOOK_CLIENT_SECRET`
- `OUTLOOK_USER_ID`

### Run a sync

```powershell
python .\scripts\sync_outlook_mail.py --max-messages 25
```

Optional:

```powershell
python .\scripts\sync_outlook_mail.py --since-days 3 --folder inbox
```

### Output layout

Each synced message is saved into its own folder under `raw/mail/messages/` with:

- `message.json`
- `body.html`
- `body.txt`
- `source.md`
- `attachments/`

Derived attachment analysis is written separately under:

- `derived/mail/<message-key>/attachment_analysis.json`

The sync also maintains:

- `raw/mail/index.json`
- `raw/mail/state.json`

### Notes

- The repository ignores downloaded mail content by default because messages and attachments may be sensitive.
- `desktop` is the preferred mode for this MVP because Outlook is already configured on this machine.
- `graph` remains available as the later server-side path.

## Task Scheduler

Recommended interactive task command:

```powershell
schtasks /Create /TN "Crimson Automation Mail Cycle" /SC MINUTE /MO 10 /TR "powershell.exe -NoProfile -ExecutionPolicy Bypass -File \"C:\projects\automation\scripts\run_mail_cycle_ui.ps1\"" /RU "$env:USERNAME" /IT /F
```

Notes:

- `/IT` keeps the task interactive, so the visual monitor window is shown.
- This is recommended for the desktop Outlook provider.
