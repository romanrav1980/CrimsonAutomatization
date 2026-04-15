## [2026-04-14] process-update | Mail attachment analysis MVP

- added a derived attachment-analysis layer under `derived/mail/<message-key>/attachment_analysis.json`
- integrated attachment analysis into the SQLite/read-model pipeline
- added attachment analysis visibility to `Needs Decision`
- verified processing on current `raw/mail` snapshot: `25` messages, `103` analyzed attachments, `0` failed records

## [2026-04-14] debug | Full interactive mail cycle after attachment rollout

- fixed the missing `derived_root` wiring in `scripts/sync_outlook_mail.py`
- verified the full interactive cycle through the visual monitor
- confirmed successful path: `Sync -> Process -> Normalize -> Attachments -> Classify -> Decision -> Backlog`
- confirmed historical backlog metrics: unread `4517`, not in raw `7124`, unread and not in raw `4517`

## [2026-04-14] process-update | Historical unread catch-up

- added scheduled catch-up for historical unread mail that is still missing from `raw/`
- catch-up explicitly skips already-read historical mail
- verified a live catch-up batch:
  - `100` historical unread messages downloaded
  - `508` attachments downloaded with those messages
  - backlog reduced from `4517` unread-not-in-raw to `4417`
  - pending unread attachment count now `18133`

## [2026-04-14] process-update | Monitor ETA and stopped-state visibility

- upgraded the mail-cycle monitor with explicit pipeline-state visualization:
  - `ACTIVE`
  - `WAITING`
  - `ALIVE / QUIET`
  - `FINALIZING`
  - `COMPLETED - all processes stopped`
  - `FAILED - all processes stopped`
  - `STOPPED - all processes stopped`
- added `Current stage ETA`, `Next stage ETA`, and `Remaining ETA`
- persisted learned stage timings in `data/mail_cycle_timings.json`
- verified a later successful interactive run:
  - `logs/mail_cycle_20260414_024934.log`
  - `100` historical unread messages downloaded
  - `385` historical unread attachments downloaded
  - backlog after the run:
    - historical unread not in `raw`: `4303`
    - pending unread attachments not in `raw`: `17638`

## [2026-04-14] process-update | Full mailbox archive and folder mirror

- added full-mailbox archive mode via `--full-backfill`
- desktop Outlook ingest can now walk the full mailbox tree instead of a single default folder
- added mirrored folder views under `raw/mail/by_folder/.../folder.json`
- verified live full-backfill start in:
  - `logs/mail_cycle_20260414_083406.log`
  - `Scanning 153 Outlook mail folder(s).`
- verified live message artifacts now include correct folder metadata such as:
  - `sourceFolderName`
  - `sourceFolderPath`
  - `sourceStoreName`
- hardened the monitor after a live popup caused by log-file contention:
  - switched monitor log appends to shared-access writes with retry
  - degraded log-write failures to in-window warnings instead of crashing the monitor

## [2026-04-14] debug | Full mailbox archive stop-state

- the live full-backfill run `logs/mail_cycle_20260414_083406.*` stopped on an Outlook COM attachment edge case
- recovered archive state from the raw artifacts on disk:
  - `10298` message directories under `raw/mail/messages`
  - `8463` messages with attachments
  - `0` attachment-bearing messages missing `attachments.json`
- rebuilt `index.json`, `folders.json`, and `raw/mail/by_folder/...` from saved `message.json`
- confirmed folder metadata completeness on all checked saved messages:
  - `sourceFolderPath`
  - `sourceFolderName`
  - `sourceStoreName`
- recovered processed folder counts:
  - `r.rabayev@korzinka.uz/Входящие`: `10021`
  - `r.rabayev@korzinka.uz/Удаленные`: `277`
- archive was not complete:
  - mailbox tree had `153` folders
  - only two folders were reached before failure

## [2026-04-15] process-update | Successful resumed full mailbox archive

- fixed the Outlook COM attachment fallback so broken attachment objects degrade to metadata-only stubs
- reran the full mailbox archive successfully in `logs/mail_cycle_20260414_214433.*`
- run result:
  - `9151` new messages synced
  - `10576` messages skipped as already archived
  - `35266` attachments downloaded during the run
  - total raw saved messages on disk after completion: `19449`
  - total saved messages with attachments: `15368`
- confirmed folder metadata completeness across all saved `message.json` files:
  - `sourceFolderPath`
  - `sourceFolderName`
  - `sourceStoreName`
- confirmed no saved attachment-bearing message was missing `attachments.json`
- rebuilt and verified raw folder artifacts:
  - `raw/mail/folders.json`
  - `raw/mail/by_folder/...`
- current post-run gap remains:
  - `200` historical messages still not in `raw`
  - `17` pending attachments on those unread historical messages not yet in `raw`

## [2026-04-15] handoff | Stable working baseline fixed

- verified the current codebase as a working baseline
- checks passed:
  - `logs/mail_cycle_20260414_214433.exitcode.txt` = `0`
  - `python -m py_compile` for the active Outlook ingest entry points
  - `dotnet build .\Automation.sln`
- saved a stable handoff record:
  - `wiki/handoffs/stable-working-state-2026-04-15.md`
- created a local code snapshot archive because this workspace is not a `git` repository:
  - `snapshots/automation-working-state-2026-04-15.zip`

## [2026-04-15] bugfix | Historical backlog key alignment and zero-gap verification

- fixed a key mismatch in `services/mail_ingest/outlook_desktop_client.py`
- root cause:
  - raw storage keyed some Outlook messages by `EntryID`
  - historical backlog and catch-up logic had been keying the same items by `StoreID::EntryID`
  - this made already archived messages appear as both `existing` and `not in raw`
- aligned `_message_key_for_item()` with `MailStorage.message_key()`
- verified with interactive Outlook monitor run:
  - `logs/mail_cycle_20260415_065209.log`
  - `logs/mail_cycle_20260415_065209.stdout.log`
  - `logs/mail_cycle_20260415_065209.exitcode.txt`
- final verified backlog state:
  - `historicalNotInRaw`: `0`
  - `historicalUnreadNotInRaw`: `0`
  - `historicalUnreadNotInRawAttachmentCount`: `0`
- total raw saved messages on disk remained `19449`

## [2026-04-15] process-update | Full downstream refresh on complete raw archive

- ran the full downstream pipeline after confirming `raw/` archive completeness
- refresh result:
  - `19449` raw mail artifacts loaded
  - `81742` attachments analyzed
  - `19449` mail artifacts classified
  - `19449` mail items processed
  - `0` failed items
- refreshed read-model summary:
  - `needsDecision`: `16715`
  - `autoReady`: `2734`
  - `manualReview`: `773`
  - `highUrgency`: `2208`

## [2026-04-15] ui-update | Needs Decision operator workbench history

- upgraded the `Needs Decision` screen with a stronger operator-oriented summary strip
- added raw archive completeness visibility to the UI
- added queue visibility context (`shown` vs total queue size)
- added per-item recent audit history directly in the dashboard read model and UI
- re-exported `data/mail_triage_read_model.json` after the read-model schema extension
