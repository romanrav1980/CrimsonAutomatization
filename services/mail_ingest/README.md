# Outlook Mail Ingest Subproject

## Purpose

This subproject is responsible for reliable mail acquisition and raw artifact storage.

Its job is to:

- connect to Outlook;
- read incoming messages from the configured mailbox and folder, or from the whole mailbox tree;
- catch up historical unread mail in scheduled batches;
- save original message artifacts into `raw/mail/`;
- preserve Outlook folder context for each message;
- build a mirrored folder-view under `raw/mail/by_folder/`;
- preserve attachments;
- maintain sync state for incremental runs;
- export historical backlog status for previous years.

This layer is intentionally separate from classification, decisioning, and UI.

## Scope

Included in this subproject:

- Outlook Desktop COM ingestion for the local Windows MVP;
- full-mailbox archive sync across all Outlook mail folders;
- Microsoft Graph client kept as the future server-side path;
- raw mail storage format;
- sync state and mail index;
- mailbox/folder configuration.

Not included in this subproject:

- classification logic;
- decision matrix logic;
- SQLite normalization;
- `Needs Decision` UI;
- Fabric export.

Those belong to downstream parts of the project.

## Entry Points

- `scripts/sync_outlook_mail.py`
- `services/mail_ingest/outlook_sync.py`

## Main Modules

- `config.py`
  - environment loading and settings parsing;
- `outlook_desktop_client.py`
  - Outlook COM/MAPI access for the local desktop profile;
- `graph_client.py`
  - future Microsoft Graph path;
- `outlook_sync.py`
  - provider selection and sync orchestration;
- `storage.py`
  - raw message, attachment, index, and state persistence.

## Output Contract

Each synced message is stored under `raw/mail/messages/<message-folder>/` with:

- `message.json`
- `body.txt`
- `body.html` when available
- `source.md`
- `attachments/`

Important raw metadata now includes:

- original Outlook folder name;
- original Outlook folder path;
- Outlook store/mailbox name.

Repository-level state artifacts:

- `raw/mail/index.json`
- `raw/mail/state.json`
- `data/mail_ingest_status.json`
- `raw/mail/folders.json`
- `raw/mail/by_folder/.../folder.json`

`data/mail_ingest_status.json` is the UI-facing snapshot used to show:

- unread messages from previous years;
- messages from previous years that are still not stored in `raw/`;
- unread messages from previous years that still need to be read and saved into `raw/`;
- unread messages from previous years that have attachments;
- attachment count still pending on unread historical messages that are not yet in `raw/`;
- per-year backlog counts.

`raw/mail/folders.json` is the raw-layer folder catalog used to preserve folder structure for future rules and analytics.

`raw/mail/by_folder/` is the mirrored folder tree. It does not duplicate message binaries; instead, each Outlook folder gets a `folder.json` file with references back to the canonical message directories under `raw/mail/messages/...`.

## Attachment Strategy

Attachments are stored as original binaries without conversion or loss.

Current strategy:

- keep every attachment under the owning message folder;
- preserve the original filename in metadata;
- save the physical file using a stable sanitized name with a content hash suffix;
- record attachment `kind`, such as `excel`, `pdf`, `image`, or `other`;
- record `sha256`, extension, MIME type, saved path, and size in the attachment manifest.

Examples:

- Excel: `.xlsx`, `.xls`, `.xlsm`, `.csv`
- PDF: `.pdf`
- Images: `.jpg`, `.jpeg`, `.png`, and other image formats

This allows future parsing, indexing, OCR, preview generation, and deduplication without losing the original file.

## Runtime Rules

- Use `desktop` as the primary MVP provider on this machine.
- Run only in an interactive logged-in Windows session when using Outlook COM.
- Do not overwrite the original message content in `raw/mail/`.
- Treat `raw/mail/` as the archive of truth.
- Scheduled catch-up should download only historical unread mail that is still missing from `raw/`.
- Already-read historical mail should not be downloaded by the catch-up routine.
- One-time `full-backfill` runs may archive all historical mail across all folders, including already-read messages.

## Current Status

Current implementation status:

- Outlook Desktop COM ingest is implemented.
- Full mailbox-tree sync is implemented.
- Graph client remains available as a secondary future path.
- Incremental sync is working.
- Historical unread catch-up is working in scheduled batches.
- Attachments are downloaded.
- Output lands in `raw/mail/`.

## Known Gaps

- no server-hosted ingest yet;
- no webhook/delta-query production path yet;
- Outlook web links are not consistently available in the current desktop path;
- some subject/body text still shows encoding issues in downstream views and should be normalized better.

## Run Commands

- normal incremental sync:
  - `python .\scripts\sync_outlook_mail.py --process-after-sync`
- one-time full mailbox archive:
  - `python .\scripts\sync_outlook_mail.py --full-backfill`
- interactive monitor for full mailbox archive:
  - `powershell -ExecutionPolicy Bypass -File .\scripts\run_mail_cycle_ui.ps1 -SyncArgs "--full-backfill"`

Sync only:

```powershell
python .\scripts\sync_outlook_mail.py
```

Sync and process:

```powershell
python .\scripts\sync_outlook_mail.py --process-after-sync
```

Visual monitor:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\run_mail_cycle_ui.ps1
```
