# Current Session Handoff

## Status

Active MVP status as of 2026-04-14:

- Outlook Desktop COM is the primary mail ingest mode for the MVP.
- Mail ingest flows into `raw/mail/`.
- Raw mail is normalized into `data/automation.db`.
- The dashboard read model is exported to `data/mail_triage_read_model.json`.
- The MVC portal includes the `Needs Decision` workbench screen.
- The visual mail-cycle monitor is working with stage-by-stage status, live counters, elapsed time, activity state, and debug screenshots.

## What Was Completed

- Switched the MVP from Graph-first to Outlook COM-first for the local Windows machine.
- Built the mail pipeline:
  - sync Outlook -> `raw/mail`
  - process `raw/mail` -> SQLite
  - classify mail
  - apply decision matrix
  - refresh `Needs Decision`
- Added the WinForms monitor in `scripts/run_mail_cycle_ui.ps1`.
- Fixed monitor reliability issues:
  - output is read from redirected `stdout` and `stderr` files;
  - final state waits for post-exit output draining before completion;
  - exit status is read from `mail_cycle_<timestamp>.exitcode.txt` instead of unreliable `Start-Process` exit metadata;
  - screenshots are now taken from the form itself via `DrawToBitmap`, not from the underlying screen area.
- Added automatic cleanup of screenshot folders from previous days at the start of a new monitor run.
- Improved the visual monitor so long-running stages are easier to read:
  - explicit `Stage activity` line;
  - per-stage runtime and quiet-time display;
  - stage-based progress bar that keeps moving during active work.
- Added stronger visual status and ETA support to the monitor so it is obvious when work is active and when everything has stopped:
  - `Pipeline state` banner with `ACTIVE`, `WAITING`, `ALIVE / QUIET`, `FINALIZING`, `COMPLETED - all processes stopped`, `FAILED - all processes stopped`, and `STOPPED - all processes stopped`;
  - `Current stage ETA`, `Next stage ETA`, and `Remaining ETA`;
  - per-stage duration learning persisted in `data/mail_cycle_timings.json`;
  - stage rows now show `avg`, `eta`, or `done in` timing details instead of only `pending/running/done`.
- Hardened monitor log writing after a live full-backfill popup:
  - `run_mail_cycle_ui.ps1` no longer uses raw `Add-Content` for every log append;
  - log writes now use shared file access with retry logic;
  - failed log appends no longer crash the monitor loop and are downgraded to in-window warnings.
- Hardened Outlook attachment fallback for COM edge cases:
  - metadata fallback no longer re-reads fragile COM properties such as `FileName` and `Size` after the primary attachment read fails;
  - a safe metadata snapshot is captured first;
  - broken attachment objects now degrade to a stable metadata-only stub instead of crashing the whole sync.
- Added operator actions for `Needs Decision`:
  - approve suggestion;
  - mark manual review;
  - archive item;
  - assign owner with notes.
- Added protection so operator actions are preserved across later raw-mail reprocessing.
- Added historical Outlook backlog snapshot support for previous years:
  - unread messages;
  - messages not yet stored in `raw/`;
  - unread messages not yet stored in `raw/`;
  - per-year breakdown.
- Extended the raw mail contract to preserve Outlook folder context:
  - folder name;
  - folder path;
  - store/mailbox name;
  - folder catalog in `raw/mail/folders.json`.
- Improved attachment persistence design:
  - keep original binary files;
  - classify attachment kind such as `excel`, `pdf`, and `image`;
  - store stable hashed filenames and metadata for future parsing and deduplication.
- Added a separate derived attachment-analysis layer:
  - writes `derived/mail/<message-key>/attachment_analysis.json`;
  - analyzes Excel, PDF, image, and other attachments without changing `raw/`;
  - saves attachment summary and detailed insights into SQLite and the dashboard read model;
  - shows attachment analysis directly in `Needs Decision`.
- Fixed the full mail-cycle integration after the attachment-analysis rollout:
  - `scripts/sync_outlook_mail.py` now passes `derived_root` into `MailProcessingPipeline`;
  - the interactive monitor now completes the full `Sync -> Process -> Normalize -> Attachments -> Classify -> Decision -> Backlog` path successfully.
- Added guaranteed historical unread catch-up behavior for scheduled runs:
  - each cycle downloads a batch of historical unread mail that is still missing from `raw/`;
  - already-read historical mail is skipped;
  - attachments for those unread historical messages are downloaded together with the message;
  - catch-up progress is written into `raw/mail/state.json` and `data/mail_ingest_status.json`.
- Extended the raw-mail layer to support full mailbox archiving:
  - `--full-backfill` runs across all Outlook mail folders, not only the default inbox;
  - `OUTLOOK_ALL_FOLDERS=true` is supported in configuration;
  - the desktop client now streams folder-by-folder progress instead of staying silent during long mailbox scans;
  - a mirrored folder-view is written under `raw/mail/by_folder/.../folder.json` without duplicating attachments.
- Added checkpoint-oriented persistence improvements for long-running mailbox sync:
  - index checkpoints can be saved during long runs;
  - folder-view rebuilds are designed to happen during checkpoints and on finalization so structure is recoverable and visible.

## Current Known Good State

Latest verified monitor run:

- log: `logs/mail_cycle_20260414_024934.log`
- stdout: `logs/mail_cycle_20260414_024934.stdout.log`
- exit code: `logs/mail_cycle_20260414_024934.exitcode.txt`
- screenshots: `logs/screenshots/20260414_024934/`

That run completed successfully with:

- `0` new synced messages
- `25` skipped existing messages
- `100` historical unread messages synced during catch-up
- `385` attachments downloaded for those historical unread messages
- `239` normalized/classified/processed records
- `1106` analyzed attachments
- `0` failed records
- `212` items in `Needs Decision`
- historical backlog snapshot:
  - unread: `4517`
  - not in raw: `6910`
  - unread and not in raw: `4303`
  - unread with attachments: `3567`
  - unread attachment count: `18641`
  - unread not in raw with attachments: `3420`
  - unread not in raw attachment count: `17638`
- historical catch-up state:
  - enabled: `true`
  - last processed count: `100`
  - last skipped existing: `0`
  - last attachment count: `385`
  - batch size: `100`

Additional verified checks:

- operator action persistence was tested safely on copied SQLite databases;
- `Assign owner` survives reprocessing;
- `Approve` removes the item from the copied `Needs Decision` queue after reprocessing.
- `python .\scripts\process_mail_mvp.py` now completes successfully with:
  - `25` processed messages
  - `0` failed messages
  - `103` analyzed attachments
  - refreshed `data/mail_triage_read_model.json`
  - derived files in `derived/mail/...`
- a live full-mailbox backfill was launched successfully:
  - log: `logs/mail_cycle_20260414_083406.log`
  - stdout immediately confirmed `Scanning 153 Outlook mail folder(s).`
  - folder-level scanning was visible for at least:
    - `r.rabayev@korzinka.uz/Удаленные`
    - `r.rabayev@korzinka.uz/Входящие`
  - live downloaded messages already contained:
    - `sourceFolderName`
    - `sourceFolderPath`
    - `sourceStoreName`
  - live downloaded attachments were present under the owning message directory.
  - during this run, `raw/mail/messages` had already grown to at least `1476` message directories while the Python child process was still alive.
  - the run later stopped on an Outlook attachment COM edge case recorded in `logs/mail_cycle_20260414_083406.stderr.log`
  - after stop-state recovery:
    - `raw/mail/messages` contained `10298` message directories
    - all `10298` checked `message.json` files had:
      - `sourceFolderPath`
      - `sourceFolderName`
      - `sourceStoreName`
    - `8463` messages had attachments
    - `0` attachment-bearing messages were missing `attachments.json`
    - `index.json`, `folders.json`, and `raw/mail/by_folder/...` were rebuilt from the saved `message.json` files
    - recovered processed folder counts:
      - `r.rabayev@korzinka.uz/Входящие`: `10021`
      - `r.rabayev@korzinka.uz/Удаленные`: `277`
- the full mailbox archive was not complete:
  - the log had confirmed `153` Outlook mail folders in the mailbox tree
  - only two folders had been reached before the crash
  - the known lower bound from just those two folders was already `11974` items, so at least `1676` messages were still missing from `raw/` even before counting the remaining folders

## Latest Verified Run

Latest verified successful full-mailbox run:

- log: `logs/mail_cycle_20260414_214433.log`
- stdout: `logs/mail_cycle_20260414_214433.stdout.log`
- stderr: `logs/mail_cycle_20260414_214433.stderr.log`
- exit code: `logs/mail_cycle_20260414_214433.exitcode.txt`
- summary: `data/mail_cycle_20260414_214433_final_summary.json`

This run completed successfully with:

- `9151` new messages synced
- `10576` previously archived messages skipped
- `35266` attachments downloaded during the run
- total saved raw messages on disk: `19449`
- total saved messages with attachments: `15368`
- `0` saved messages missing `sourceFolderPath`
- `0` saved messages missing `sourceFolderName`
- `0` saved messages missing `sourceStoreName`
- `0` attachment-bearing saved messages missing `attachments.json`
- rebuilt raw folder artifacts:
  - `raw/mail/folders.json`
  - `raw/mail/by_folder/...`
- non-empty archived folders represented in raw artifacts: `48`

Current completeness state after the successful run:

- folder metadata is complete for all saved messages
- mirrored folder structure exists for all archived non-empty folders
- mailbox-tree scan completed across `153` Outlook mail folders
- archive is still not 100% complete:
  - historical messages not yet in `raw`: `200`
  - historical unread messages not yet in `raw`: `200`
  - pending attachments on those unread historical messages not yet in `raw`: `17`

## Important Operational Rules

- Run the desktop Outlook provider only in an interactive logged-in Windows session.
- For Task Scheduler, use the interactive mode so the monitor window is visible.
- Screenshot folders from previous days should be deleted automatically on the next run.
- Debug screenshots are temporary operational artifacts, not long-term evidence.
- Treat the monitor timing file `data/mail_cycle_timings.json` as operational state; it improves ETA quality across runs and should usually be preserved.
- For very large mailbox backfills, prefer the explicit one-time archive mode:
  - `python .\scripts\sync_outlook_mail.py --full-backfill`
  - or `run_mail_cycle_ui.ps1 -SyncArgs "--full-backfill"`

## Main Files To Read First Next Time

1. `wiki/handoffs/current-session.md`
2. `wiki/processes/mail-triage.md`
3. `wiki/services/outlook-mail-ingest.md`
4. `services/mail_ingest/README.md`
5. `scripts/run_mail_cycle_ui.ps1`
6. `scripts/sync_outlook_mail.py`
7. `services/mail_processing/pipeline.py`
8. `services/mail_processing/attachment_analysis.py`
9. `wiki/processes/mail-attachment-analysis.md`
10. `data/mail_ingest_status.json`
11. `data/mail_cycle_timings.json`

## Plan For Next Changes

The next practical implementation steps are:

1. Surface historical backlog and catch-up progress more clearly in the `Needs Decision` web UI:
   - unread historical mail still not in `raw/`;
   - unread historical attachments still not in `raw/`;
   - latest catch-up batch size and last run details.
2. Finish the `Needs Decision` screen as a full operator workbench instead of a read-only view:
   - keep `Approve`, `Archive`, `Assign owner`, and `Manual review`;
   - add clearer action history and operator audit visibility in the UI.
3. Add SLA enrichment and routing context:
   - due dates or SLA timers where they can be inferred;
   - owner or queue hints;
   - process/routing context next to each decision item.
4. Keep the historical unread catch-up draining on every scheduled run until `historicalUnreadNotInRaw` reaches zero.
5. Improve heavy attachment workflows:
   - OCR for image attachments;
   - richer PDF text extraction;
   - more useful Excel sheet previews and structured metadata.

## Tomorrow Start Point

Start tomorrow from this exact sequence:

1. Read `wiki/handoffs/current-session.md`.
2. Check `data/mail_ingest_status.json` for the latest backlog numbers.
3. If Outlook is open in an interactive session, run the monitor once and confirm:
   - `Pipeline state` renders correctly;
   - `Current stage ETA`, `Next stage ETA`, and `Remaining ETA` look sane;
   - the cycle still completes cleanly.
4. Continue implementation from the first item in `Plan For Next Changes`.

## Resume Prompt

When resuming work, start from this checkpoint:

"Continue the mail-agent MVP from the current handoff. Read `wiki/handoffs/current-session.md` first, preserve the working Outlook COM ingest and monitor behavior, and continue with the next step around `Needs Decision` operator actions and auditability."
"Continue the mail-agent MVP from the current handoff. Read `wiki/handoffs/current-session.md` first, preserve the working Outlook COM ingest and monitor behavior, generate the real historical backlog snapshot through an interactive sync, and then continue with SLA enrichment and operator audit visibility."
"Continue the mail-agent MVP from the current handoff. Read `wiki/handoffs/current-session.md` first, preserve the working Outlook COM ingest and monitor behavior, keep the derived attachment-analysis layer stable, and continue with richer OCR/SLA/operator-history work."
"Continue the mail-agent MVP from the current handoff. Read `wiki/handoffs/current-session.md` first, preserve the scheduled historical unread catch-up behavior, and continue draining the unread historical backlog before moving deeper into SLA and OCR work."
"Continue the mail-agent MVP from the current handoff. Read `wiki/handoffs/current-session.md` first, preserve the ETA-enabled monitor and the all-processes-stopped visualization, then continue from the saved `Plan For Next Changes`."
