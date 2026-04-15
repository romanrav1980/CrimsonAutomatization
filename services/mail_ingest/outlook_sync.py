from __future__ import annotations

from dataclasses import dataclass
from datetime import UTC, datetime, timedelta
from typing import Any

from services.mail_ingest.config import OutlookSettings
from services.mail_ingest.graph_client import OutlookGraphClient
from services.mail_ingest.outlook_desktop_client import OutlookDesktopClient
from services.mail_ingest.storage import MailStorage


@dataclass(slots=True)
class SyncResult:
    processed: int
    skipped_existing: int
    attachments_downloaded: int


@dataclass(slots=True)
class HistoricalUnreadCatchupResult:
    processed: int
    skipped_existing: int
    attachments_downloaded: int


class OutlookMailSyncService:
    INDEX_CHECKPOINT_EVERY = 25
    VIEW_CHECKPOINT_EVERY = 200

    def __init__(self, settings: OutlookSettings) -> None:
        self.settings = settings
        self.client = self._build_client()
        self.storage = MailStorage(settings.output_dir)

    def sync(self, *, max_messages: int | None = None, since_days: int | None = None) -> SyncResult:
        since_iso = None
        if since_days is not None:
            since_iso = (datetime.now(UTC) - timedelta(days=since_days)).replace(microsecond=0).isoformat().replace("+00:00", "Z")

        if hasattr(self.client, "iter_messages"):
            messages = self.client.iter_messages(top=max_messages, since_iso=since_iso)
        else:
            messages = self.client.list_messages(top=max_messages, since_iso=since_iso)
        index = self.storage.load_index()
        processed = 0
        skipped_existing = 0
        attachments_downloaded = 0
        changes_since_index_checkpoint = 0
        changes_since_view_checkpoint = 0

        for message in messages:
            key = self.storage.message_key(message)
            if key in index.get("messages", {}):
                self.storage.refresh_existing_message_metadata(message, index=index, rebuild_views=False)
                skipped_existing += 1
                changes_since_index_checkpoint += 1
                changes_since_view_checkpoint += 1
                self._checkpoint_progress(
                    index=index,
                    changes_since_index_checkpoint=changes_since_index_checkpoint,
                    changes_since_view_checkpoint=changes_since_view_checkpoint,
                )
                if changes_since_index_checkpoint >= self.INDEX_CHECKPOINT_EVERY:
                    changes_since_index_checkpoint = 0
                if changes_since_view_checkpoint >= self.VIEW_CHECKPOINT_EVERY:
                    changes_since_view_checkpoint = 0
                continue

            attachments = []
            if message.get("hasAttachments"):
                attachments = self.client.list_attachments(message["id"])

            stored = self.storage.store_message(
                message,
                attachments,
                save_attachments=self.settings.save_attachments,
                index=index,
                rebuild_views=False,
            )
            processed += 1
            attachments_downloaded += sum(1 for item in attachments if item.get("contentBytes"))
            changes_since_index_checkpoint += 1
            changes_since_view_checkpoint += 1
            print(f"Synced message {stored.key} -> {stored.directory}")
            self._checkpoint_progress(
                index=index,
                changes_since_index_checkpoint=changes_since_index_checkpoint,
                changes_since_view_checkpoint=changes_since_view_checkpoint,
            )
            if changes_since_index_checkpoint >= self.INDEX_CHECKPOINT_EVERY:
                changes_since_index_checkpoint = 0
            if changes_since_view_checkpoint >= self.VIEW_CHECKPOINT_EVERY:
                changes_since_view_checkpoint = 0

        self.storage.persist_index_and_views(index)

        self.storage.save_state(
            {
                "last_sync_utc": datetime.now(UTC).replace(microsecond=0).isoformat().replace("+00:00", "Z"),
                "last_requested_since_days": since_days,
                "last_processed_count": processed,
            }
        )
        return SyncResult(
            processed=processed,
            skipped_existing=skipped_existing,
            attachments_downloaded=attachments_downloaded,
        )

    def sync_historical_unread_backlog(self, *, max_messages: int | None = None) -> HistoricalUnreadCatchupResult:
        if not hasattr(self.client, "list_historical_unread_messages"):
            return HistoricalUnreadCatchupResult(processed=0, skipped_existing=0, attachments_downloaded=0)

        limit = max_messages if max_messages is not None else self.settings.historical_unread_catchup_batch_size
        if limit <= 0:
            return HistoricalUnreadCatchupResult(processed=0, skipped_existing=0, attachments_downloaded=0)

        index = self.storage.load_index()
        existing_keys = set(index.get("messages", {}).keys())
        messages = self.client.list_historical_unread_messages(existing_keys=existing_keys, top=limit)

        processed = 0
        skipped_existing = 0
        attachments_downloaded = 0
        changes_since_index_checkpoint = 0
        changes_since_view_checkpoint = 0

        for message in messages:
            key = self.storage.message_key(message)
            if key in index.get("messages", {}):
                self.storage.refresh_existing_message_metadata(message, index=index, rebuild_views=False)
                skipped_existing += 1
                changes_since_index_checkpoint += 1
                changes_since_view_checkpoint += 1
                self._checkpoint_progress(
                    index=index,
                    changes_since_index_checkpoint=changes_since_index_checkpoint,
                    changes_since_view_checkpoint=changes_since_view_checkpoint,
                )
                if changes_since_index_checkpoint >= self.INDEX_CHECKPOINT_EVERY:
                    changes_since_index_checkpoint = 0
                if changes_since_view_checkpoint >= self.VIEW_CHECKPOINT_EVERY:
                    changes_since_view_checkpoint = 0
                continue

            attachments = []
            if message.get("hasAttachments"):
                attachments = self.client.list_attachments(message["id"])

            stored = self.storage.store_message(
                message,
                attachments,
                save_attachments=self.settings.save_attachments,
                index=index,
                rebuild_views=False,
            )
            processed += 1
            attachments_downloaded += sum(1 for item in attachments if item.get("contentBytes"))
            changes_since_index_checkpoint += 1
            changes_since_view_checkpoint += 1
            print(f"Synced historical unread message {stored.key} -> {stored.directory}")
            self._checkpoint_progress(
                index=index,
                changes_since_index_checkpoint=changes_since_index_checkpoint,
                changes_since_view_checkpoint=changes_since_view_checkpoint,
            )
            if changes_since_index_checkpoint >= self.INDEX_CHECKPOINT_EVERY:
                changes_since_index_checkpoint = 0
            if changes_since_view_checkpoint >= self.VIEW_CHECKPOINT_EVERY:
                changes_since_view_checkpoint = 0

        self.storage.persist_index_and_views(index)

        state = self.storage.load_state()
        state["historical_unread_catchup"] = {
            "enabled": self.settings.historical_unread_catchup_enabled,
            "last_run_utc": datetime.now(UTC).replace(microsecond=0).isoformat().replace("+00:00", "Z"),
            "last_processed_count": processed,
            "last_skipped_existing": skipped_existing,
            "last_attachment_count": attachments_downloaded,
            "batch_size": limit,
        }
        self.storage.save_state(state)

        return HistoricalUnreadCatchupResult(
            processed=processed,
            skipped_existing=skipped_existing,
            attachments_downloaded=attachments_downloaded,
        )

    def build_historical_backlog_snapshot(self) -> dict[str, Any] | None:
        if not hasattr(self.client, "collect_historical_backlog_snapshot"):
            return None

        index = self.storage.load_index()
        existing_keys = set(index.get("messages", {}).keys())
        return self.client.collect_historical_backlog_snapshot(existing_keys=existing_keys)

    def _build_client(self):
        if self.settings.provider == "desktop":
            return OutlookDesktopClient(self.settings)
        return OutlookGraphClient(self.settings)

    def _checkpoint_progress(
        self,
        *,
        index: dict[str, Any],
        changes_since_index_checkpoint: int,
        changes_since_view_checkpoint: int,
    ) -> None:
        if changes_since_index_checkpoint >= self.INDEX_CHECKPOINT_EVERY:
            self.storage.save_index(index)
            print(f"Checkpoint saved index after {changes_since_index_checkpoint} changes")
        if changes_since_view_checkpoint >= self.VIEW_CHECKPOINT_EVERY:
            self.storage.rebuild_views(index)
            print(f"Checkpoint rebuilt folder views after {changes_since_view_checkpoint} changes")
