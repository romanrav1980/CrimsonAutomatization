from __future__ import annotations

import argparse
import json
import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from services.mail_ingest.config import OutlookSettings, load_env_file
from services.mail_ingest.outlook_sync import OutlookMailSyncService
from services.mail_processing.pipeline import MailProcessingPipeline


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Sync Outlook mail into raw storage.")
    parser.add_argument("--max-messages", type=int, default=None, help="Maximum number of messages to fetch.")
    parser.add_argument("--since-days", type=int, default=None, help="Only fetch messages received in the last N days.")
    parser.add_argument("--folder", type=str, default=None, help="Override the configured mail folder, for example inbox.")
    parser.add_argument("--all-folders", action="store_true", help="Sync all mail folders in the mailbox instead of a single default folder.")
    parser.add_argument("--full-backfill", action="store_true", help="Run a one-time full mailbox archive sync across all folders and all history.")
    parser.add_argument("--skip-catchup", action="store_true", help="Skip the historical unread catch-up stage for this run.")
    parser.add_argument("--process-after-sync", action="store_true", help="Process raw mail into the MVP database after sync.")
    return parser


def emit_stage(stage: str, state: str, detail: str) -> None:
    print(f"[[STAGE|{stage}|{state}|{detail}]]")


def emit_metric(name: str, value: int, label: str) -> None:
    print(f"[[METRIC|{name}|{value}|{label}]]")


def write_ingest_status(path: Path, payload: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    load_env_file()
    settings = OutlookSettings.from_env()
    if args.folder:
        settings.folder = args.folder.strip().lower()
    if args.all_folders:
        settings.all_folders = True
    if args.full_backfill:
        settings.all_folders = True
        args.max_messages = 0
        args.since_days = None
        args.skip_catchup = True

    service = OutlookMailSyncService(settings)
    target_scope = "all mail folders" if settings.all_folders else f"folder '{settings.folder}'"
    emit_stage("sync", "running", f"Syncing Outlook provider '{settings.provider}' from {target_scope}")
    result = service.sync(max_messages=args.max_messages, since_days=args.since_days)
    emit_stage("sync", "done", f"Stored {result.processed} new messages, skipped {result.skipped_existing}")
    emit_metric("synced_new", result.processed, "New messages synced")
    emit_metric("skipped_existing", result.skipped_existing, "Existing messages skipped")
    emit_metric("attachments_downloaded", result.attachments_downloaded, "Attachments downloaded")

    catchup_result = None
    if args.skip_catchup:
        emit_stage("catchup", "done", "Historical unread catch-up was skipped for this run")
    elif settings.historical_unread_catchup_enabled:
        emit_stage(
            "catchup",
            "running",
            (
                "Downloading up to {0} historical unread messages that are not yet in raw".format(
                    settings.historical_unread_catchup_batch_size
                )
            ),
        )
        catchup_result = service.sync_historical_unread_backlog(
            max_messages=settings.historical_unread_catchup_batch_size,
        )
        emit_stage(
            "catchup",
            "done",
            (
                "Stored {0} historical unread messages, skipped {1}, downloaded {2} attachment(s)".format(
                    catchup_result.processed,
                    catchup_result.skipped_existing,
                    catchup_result.attachments_downloaded,
                )
            ),
        )
        emit_metric("catchup_unread_synced", catchup_result.processed, "Historical unread messages synced")
        emit_metric(
            "catchup_attachments_downloaded",
            catchup_result.attachments_downloaded,
            "Historical unread attachments downloaded",
        )
    else:
        emit_stage("catchup", "done", "Historical unread catch-up is disabled by configuration")

    if args.process_after_sync:
        emit_stage("process", "running", "Starting raw mail processing pipeline")
        pipeline = MailProcessingPipeline(
            raw_root=settings.output_dir,
            derived_root=Path(os.getenv("AUTOMATION_DERIVED_ROOT", "derived/mail")),
            db_path=Path(os.getenv("AUTOMATION_DB_PATH", "data/automation.db")),
            read_model_path=Path(os.getenv("AUTOMATION_READ_MODEL_PATH", "data/mail_triage_read_model.json")),
            internal_domains={item.strip().lower() for item in os.getenv("MAIL_INTERNAL_DOMAINS", "").split(",") if item.strip()},
        )
        processed = pipeline.process()
        emit_stage("process", "done", f"Updated database and read model with {processed.processed} items")
        emit_metric("processed", processed.processed, "Processed items")
        emit_metric("failed", processed.failed, "Failed items")
        emit_metric("needs_decision", processed.needs_decision, "Items in Needs Decision queue")
        print(f"Processed raw messages into MVP database: {processed.processed}")
        print(f"Failed raw messages during processing: {processed.failed}")
        print(f"Dashboard read model: {processed.exported_read_model_path}")

    emit_stage("backlog", "running", "Scanning historical Outlook backlog from previous years")
    backlog_snapshot = service.build_historical_backlog_snapshot()
    if backlog_snapshot:
        backlog_snapshot["historicalCatchup"] = service.storage.load_state().get("historical_unread_catchup", {})
        ingest_status_path = Path(os.getenv("AUTOMATION_MAIL_INGEST_STATUS_PATH", "data/mail_ingest_status.json"))
        write_ingest_status(ingest_status_path, backlog_snapshot)
        historical_backlog = backlog_snapshot.get("historicalBacklog", {})
        emit_metric(
            "historical_unread",
            int(historical_backlog.get("historicalUnread", 0)),
            "Historical unread messages",
        )
        emit_metric(
            "historical_not_in_raw",
            int(historical_backlog.get("historicalNotInRaw", 0)),
            "Historical messages not yet stored in raw",
        )
        emit_metric(
            "historical_unread_not_in_raw",
            int(historical_backlog.get("historicalUnreadNotInRaw", 0)),
            "Historical unread messages not yet stored in raw",
        )
        emit_metric(
            "historical_unread_with_attachments",
            int(historical_backlog.get("historicalUnreadWithAttachments", 0)),
            "Historical unread messages with attachments",
        )
        emit_metric(
            "historical_unread_attachment_count",
            int(historical_backlog.get("historicalUnreadAttachmentCount", 0)),
            "Attachments on historical unread messages",
        )
        emit_metric(
            "historical_unread_not_in_raw_attachment_count",
            int(historical_backlog.get("historicalUnreadNotInRawAttachmentCount", 0)),
            "Attachments on historical unread messages not yet in raw",
        )
        emit_stage(
            "backlog",
            "done",
            (
                "Historical backlog unread {0}, not in raw {1}, unread and not in raw {2}, pending attachments {3}".format(
                    int(historical_backlog.get("historicalUnread", 0)),
                    int(historical_backlog.get("historicalNotInRaw", 0)),
                    int(historical_backlog.get("historicalUnreadNotInRaw", 0)),
                    int(historical_backlog.get("historicalUnreadNotInRawAttachmentCount", 0)),
                )
            ),
        )
    else:
        emit_stage("backlog", "done", "Historical backlog snapshot is not available for the current provider")

    print("")
    print("Sync complete")
    print(f"Processed: {result.processed}")
    print(f"Skipped existing: {result.skipped_existing}")
    print(f"Attachments downloaded: {result.attachments_downloaded}")
    if catchup_result is not None:
        print(f"Historical unread synced: {catchup_result.processed}")
        print(f"Historical unread attachments downloaded: {catchup_result.attachments_downloaded}")
    print(f"Output directory: {settings.output_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
