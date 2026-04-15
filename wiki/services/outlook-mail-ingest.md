# Outlook Mail Ingest

## Summary

This is a separate subproject inside the repository.

It owns Outlook connectivity and the `raw/mail/` archive, and it feeds downstream normalization, classification, and decision workflows.

## Boundaries

This subproject is responsible for:

- connecting to Outlook;
- selecting mailbox and folder;
- syncing messages incrementally;
- catching up historical unread mail incrementally;
- preserving the original Outlook folder location for each message;
- downloading attachments;
- writing stable raw artifacts;
- keeping sync state.

This subproject is not responsible for:

- deciding what messages mean for the business;
- assigning owners;
- applying SLA policy;
- driving operator actions in the web UI;
- exporting reporting facts into Fabric.

## Source Of Truth

For mail acquisition, the source-of-truth layers are:

1. Outlook mailbox
2. `raw/mail/` archived artifacts
3. sync state in `raw/mail/state.json`

All later business interpretation should happen outside this layer.

## Inputs

- configured Outlook profile;
- mailbox name;
- folder name;
- sync parameters such as max messages and date window.

## Outputs

- `raw/mail/messages/...`
- `raw/mail/index.json`
- `raw/mail/state.json`
- `raw/mail/folders.json`
- `data/mail_ingest_status.json`

## Implemented Components

- `services/mail_ingest/config.py`
- `services/mail_ingest/outlook_desktop_client.py`
- `services/mail_ingest/graph_client.py`
- `services/mail_ingest/outlook_sync.py`
- `services/mail_ingest/storage.py`
- `scripts/sync_outlook_mail.py`

## Current Status

Status: MVP working locally.

What is already working:

- Outlook Desktop COM is the primary ingest mode.
- Mail is synced into `raw/mail/`.
- Historical unread mail can be drained into `raw/mail/` in scheduled batches.
- Historical catch-up only downloads unread mail that is still missing from `raw/`.
- Outlook folder name/path/store can be preserved in raw metadata.
- Attachments are saved.
- Attachments can be categorized as `excel`, `pdf`, `image`, or `other`.
- Incremental sync avoids rewriting already known messages.
- The sync can be launched directly or through the visual monitor.
- The sync can export a historical backlog snapshot for previous years.

What is still pending:

- hardening for a server-side Graph-based mode;
- stronger encoding normalization for Cyrillic subjects and bodies;
- richer metadata extraction such as durable thread and Outlook link fields;
- more explicit ingest health reporting in the dashboard.

## Integration Points

Downstream dependencies:

- `services/mail_processing/`
- `data/automation.db`
- `data/mail_triage_read_model.json`
- `src/Automation.Web`

Upstream dependency:

- local Outlook profile on the Windows workstation.

## Runbook Notes

- use interactive logged-in Windows mode for COM;
- prefer Task Scheduler with visible UI for the MVP;
- keep screenshot cleanup automatic;
- do not treat screenshot folders as permanent evidence.

## Next Step

Keep this subproject stable while the next product work moves to:

1. operator actions in `Needs Decision`;
2. audit trail for human decisions;
3. SLA enrichment and routing context.
