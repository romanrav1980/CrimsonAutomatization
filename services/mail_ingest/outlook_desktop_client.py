from __future__ import annotations

import base64
import mimetypes
import tempfile
from collections import defaultdict
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

from services.mail_ingest.config import OutlookSettings
from services.mail_ingest.storage import MailStorage


OL_FOLDER_IDS = {
    "inbox": 6,
    "sentitems": 5,
    "drafts": 16,
    "deleteditems": 3,
    "junk": 23,
    "outbox": 4,
    "calendar": 9,
    "contacts": 10,
    "tasks": 13,
    "notes": 12,
    "journal": 11,
    "входящие": 6,
    "отправленные": 5,
    "черновики": 16,
    "удаленные": 3,
}

MAIL_ITEM_CLASS = 43


class OutlookDesktopClient:
    def __init__(self, settings: OutlookSettings) -> None:
        self.settings = settings

        try:
            import win32com.client  # type: ignore[import-untyped]
        except ImportError as exc:
            raise RuntimeError("Desktop Outlook provider requires pywin32 / win32com.") from exc

        self._win32com = win32com.client
        try:
            application = self._win32com.gencache.EnsureDispatch("Outlook.Application")
            self._namespace = application.GetNamespace("MAPI")
        except Exception as exc:
            raise RuntimeError(
                "Desktop Outlook provider requires an interactive logged-in Windows session with an accessible Outlook profile."
            ) from exc

    def list_messages(self, *, top: int | None = None, since_iso: str | None = None) -> list[dict[str, Any]]:
        return list(self.iter_messages(top=top, since_iso=since_iso))

    def iter_messages(self, *, top: int | None = None, since_iso: str | None = None):
        folders = self._resolve_target_folders()
        limit = self.settings.max_messages if top is None else top
        if limit is not None and limit <= 0:
            limit = None
        since_dt = self._parse_since(since_iso)
        yielded = 0

        print(f"Scanning {len(folders)} Outlook mail folder(s).")
        for folder in folders:
            folder_path = self._normalize_folder_path(self._safe_str(getattr(folder, "FolderPath", ""))) or self._safe_str(getattr(folder, "Name", "")) or "_unknown"
            items = folder.Items
            items.Sort("[ReceivedTime]", True)
            print(f"Scanning Outlook folder {folder_path} ({int(getattr(items, 'Count', 0))} items)")

            index = 1
            while index <= items.Count:
                if limit is not None and yielded >= limit:
                    return

                item = items.Item(index)
                index += 1

                if getattr(item, "Class", None) != MAIL_ITEM_CLASS:
                    continue

                received_iso = self._to_utc_iso(getattr(item, "ReceivedTime", None))
                if since_dt and received_iso and self._parse_iso(received_iso) < since_dt:
                    break

                yielded += 1
                yield self._serialize_message(item, received_iso)

    def list_attachments(self, message_id: str) -> list[dict[str, Any]]:
        entry_id, store_id = self._split_message_id(message_id)
        item = self._namespace.GetItemFromID(entry_id, store_id)
        attachments = []
        for index in range(1, int(getattr(item.Attachments, "Count", 0)) + 1):
            attachment = item.Attachments.Item(index)
            snapshot = self._snapshot_attachment_metadata(attachment, index=index)
            try:
                attachments.append(self._serialize_attachment(attachment, owner_id=message_id, index=index, snapshot=snapshot))
            except Exception as exc:
                attachments.append(self._serialize_attachment_metadata_only(snapshot, owner_id=message_id, index=index, error=str(exc)))
        return attachments

    def list_historical_unread_messages(self, *, existing_keys: set[str], top: int) -> list[dict[str, Any]]:
        if top <= 0:
            return []

        current_year = datetime.now().year
        results: list[dict[str, Any]] = []
        for folder in self._resolve_target_folders():
            items = folder.Items
            items.Sort("[ReceivedTime]", False)

            index = 1
            while index <= items.Count:
                item = items.Item(index)
                index += 1

                if getattr(item, "Class", None) != MAIL_ITEM_CLASS:
                    continue

                received_value = getattr(item, "ReceivedTime", None)
                if not isinstance(received_value, datetime):
                    continue
                if received_value.year >= current_year:
                    continue
                if not bool(getattr(item, "UnRead", False)):
                    continue

                message_key = self._message_key_for_item(item)
                if message_key in existing_keys:
                    continue

                received_iso = self._to_utc_iso(received_value)
                results.append(self._serialize_message(item, received_iso))

        results.sort(key=lambda item: item.get("receivedDateTime") or "")
        return results[:top]

    def collect_historical_backlog_snapshot(self, *, existing_keys: set[str]) -> dict[str, Any]:
        current_year = datetime.now().year
        by_year: dict[int, dict[str, int]] = defaultdict(
            lambda: {
                "totalMessages": 0,
                "unreadMessages": 0,
                "notInRaw": 0,
                "unreadNotInRaw": 0,
            }
        )

        total_historical_messages = 0
        historical_unread = 0
        historical_unsynced = 0
        historical_unread_unsynced = 0
        historical_unread_with_attachments = 0
        historical_unread_attachment_count = 0
        historical_unread_unsynced_with_attachments = 0
        historical_unread_unsynced_attachment_count = 0
        oldest_year: int | None = None

        for folder in self._resolve_target_folders():
            items = folder.Items
            items.Sort("[ReceivedTime]", True)

            index = 1
            while index <= items.Count:
                item = items.Item(index)
                index += 1

                if getattr(item, "Class", None) != MAIL_ITEM_CLASS:
                    continue

                received_value = getattr(item, "ReceivedTime", None)
                if not isinstance(received_value, datetime):
                    continue

                message_year = received_value.year
                if message_year >= current_year:
                    continue

                total_historical_messages += 1
                year_bucket = by_year[message_year]
                year_bucket["totalMessages"] += 1
                oldest_year = message_year if oldest_year is None else min(oldest_year, message_year)

                is_unread = bool(getattr(item, "UnRead", False))
                attachment_count = int(getattr(getattr(item, "Attachments", None), "Count", 0) or 0)
                if is_unread:
                    historical_unread += 1
                    year_bucket["unreadMessages"] += 1
                    if attachment_count > 0:
                        historical_unread_with_attachments += 1
                        historical_unread_attachment_count += attachment_count

                message_key = self._message_key_for_item(item)
                not_in_raw = message_key not in existing_keys
                if not_in_raw:
                    historical_unsynced += 1
                    year_bucket["notInRaw"] += 1

                if is_unread and not_in_raw:
                    historical_unread_unsynced += 1
                    year_bucket["unreadNotInRaw"] += 1
                    if attachment_count > 0:
                        historical_unread_unsynced_with_attachments += 1
                        historical_unread_unsynced_attachment_count += attachment_count

        years = []
        for year in sorted(by_year.keys(), reverse=True):
            years.append(
                {
                    "year": year,
                    "totalMessages": by_year[year]["totalMessages"],
                    "unreadMessages": by_year[year]["unreadMessages"],
                    "notInRaw": by_year[year]["notInRaw"],
                    "unreadNotInRaw": by_year[year]["unreadNotInRaw"],
                }
            )

        return {
            "generatedUtc": datetime.now(UTC).replace(microsecond=0).isoformat().replace("+00:00", "Z"),
            "provider": "desktop",
            "mailbox": self.settings.mailbox_name or self.settings.user_id or "",
            "folder": "all-folders" if self.settings.all_folders else self.settings.folder,
            "currentYear": current_year,
            "historicalBacklog": {
                "totalHistoricalMessages": total_historical_messages,
                "historicalUnread": historical_unread,
                "historicalNotInRaw": historical_unsynced,
                "historicalUnreadNotInRaw": historical_unread_unsynced,
                "historicalUnreadWithAttachments": historical_unread_with_attachments,
                "historicalUnreadAttachmentCount": historical_unread_attachment_count,
                "historicalUnreadNotInRawWithAttachments": historical_unread_unsynced_with_attachments,
                "historicalUnreadNotInRawAttachmentCount": historical_unread_unsynced_attachment_count,
                "oldestYear": oldest_year,
                "years": years,
            },
        }

    def _resolve_folder(self):
        folder_id = OL_FOLDER_IDS.get(self.settings.folder.lower())
        if folder_id is None:
            raise ValueError(f"Unsupported Outlook desktop folder: {self.settings.folder}")

        matched_store = self._find_store()
        return matched_store.GetDefaultFolder(folder_id)

    def _resolve_target_folders(self) -> list[Any]:
        if not self.settings.all_folders:
            return [self._resolve_folder()]

        matched_store = self._find_store()
        root_folder = matched_store.GetRootFolder()
        results: list[Any] = []
        self._walk_mail_folders(root_folder, results)
        return results

    def _find_store(self):

        target_name = (self.settings.mailbox_name or self.settings.user_id or "").strip().lower()
        stores = self._namespace.Stores

        matched_store = None
        for index in range(1, stores.Count + 1):
            store = stores.Item(index)
            display_name = str(getattr(store, "DisplayName", "") or "").strip().lower()
            if not target_name or display_name == target_name or target_name in display_name:
                matched_store = store
                if target_name:
                    break

        if matched_store is None:
            raise RuntimeError(
                f"Could not find Outlook store for mailbox '{self.settings.mailbox_name or self.settings.user_id or 'default'}'."
            )

        return matched_store

    def _walk_mail_folders(self, folder, results: list[Any]) -> None:
        try:
            default_item_type = int(getattr(folder, "DefaultItemType", -1))
        except Exception:
            default_item_type = -1

        folder_path = self._normalize_folder_path(self._safe_str(getattr(folder, "FolderPath", "")))
        if folder_path and default_item_type == 0:
            results.append(folder)

        subfolders = getattr(folder, "Folders", None)
        if subfolders is None:
            return

        for index in range(1, int(getattr(subfolders, "Count", 0)) + 1):
            try:
                child = subfolders.Item(index)
            except Exception:
                continue
            self._walk_mail_folders(child, results)

    def _serialize_message(self, item, received_iso: str) -> dict[str, Any]:
        parent_folder = getattr(item, "Parent", None)
        store_id = self._safe_str(getattr(getattr(item, "Parent", None), "StoreID", ""))
        entry_id = self._safe_str(getattr(item, "EntryID", ""))
        subject = self._safe_str(getattr(item, "Subject", ""))
        body_text = self._safe_str(getattr(item, "Body", ""))
        body_html = self._safe_str(getattr(item, "HTMLBody", ""))
        sender_address = self._resolve_address(getattr(item, "Sender", None), fallback=self._safe_str(getattr(item, "SenderEmailAddress", "")))
        internet_message_id = self._get_property(item, "http://schemas.microsoft.com/mapi/proptag/0x1035001F")
        source_folder_name = self._safe_str(getattr(parent_folder, "Name", ""))
        source_folder_path = self._normalize_folder_path(self._safe_str(getattr(parent_folder, "FolderPath", "")))
        source_store_name = self._safe_str(getattr(getattr(parent_folder, "Store", None), "DisplayName", "")) or self.settings.mailbox_name or ""
        return {
            "id": self._compose_message_id(entry_id, store_id),
            "internetMessageId": internet_message_id or entry_id,
            "conversationId": self._safe_str(getattr(item, "ConversationID", "")),
            "subject": subject,
            "sender": {"emailAddress": {"address": sender_address}},
            "from": {"emailAddress": {"address": sender_address}},
            "toRecipients": self._serialize_recipients(getattr(item, "Recipients", None)),
            "ccRecipients": [],
            "bccRecipients": [],
            "receivedDateTime": received_iso,
            "sentDateTime": self._to_utc_iso(getattr(item, "SentOn", None)),
            "hasAttachments": bool(getattr(item, "Attachments", None) and item.Attachments.Count > 0),
            "importance": self._map_importance(getattr(item, "Importance", 1)),
            "isRead": bool(getattr(item, "UnRead", False) is False),
            "categories": self._parse_categories(self._safe_str(getattr(item, "Categories", ""))),
            "bodyPreview": body_text[:240],
            "body": {
                "contentType": "html" if body_html else "text",
                "content": body_html or body_text,
            },
            "webLink": "",
            "sourceFolderName": source_folder_name,
            "sourceFolderPath": source_folder_path,
            "sourceStoreName": source_store_name,
        }

    def _serialize_attachment(self, attachment, *, owner_id: str, index: int, snapshot: dict[str, Any] | None = None) -> dict[str, Any]:
        snapshot = snapshot or self._snapshot_attachment_metadata(attachment, index=index)
        with tempfile.TemporaryDirectory() as temp_dir:
            name = snapshot["name"]
            suffix = Path(name).suffix
            temp_path = Path(temp_dir) / f"attachment-{index}{suffix}"
            attachment.SaveAsFile(str(temp_path))
            content_bytes = temp_path.read_bytes()

        content_type = snapshot["contentType"]
        return {
            "id": f"{owner_id}-{index}",
            "name": name,
            "contentType": content_type,
            "size": len(content_bytes),
            "@odata.type": "#microsoft.graph.fileAttachment",
            "contentBytes": base64.b64encode(content_bytes).decode("ascii"),
            "isInline": snapshot["isInline"],
            "lastModifiedDateTime": "",
        }

    def _serialize_attachment_metadata_only(self, snapshot: dict[str, Any], *, owner_id: str, index: int, error: str) -> dict[str, Any]:
        name = snapshot["name"]
        return {
            "id": f"{owner_id}-{index}",
            "name": name,
            "contentType": snapshot["contentType"],
            "size": snapshot["size"],
            "@odata.type": snapshot["odataType"],
            "isInline": snapshot["isInline"],
            "lastModifiedDateTime": "",
            "extractionError": error,
        }

    def _snapshot_attachment_metadata(self, attachment, *, index: int) -> dict[str, Any]:
        name = self._safe_com_str(attachment, "FileName") or self._safe_com_str(attachment, "DisplayName") or f"attachment-{index}"
        content_type = mimetypes.guess_type(name)[0] or "application/octet-stream"
        size = self._safe_com_int(attachment, "Size")
        is_inline = bool(self._safe_com_int(attachment, "Position", default=-1) > 0)
        attachment_type = self._safe_com_int(attachment, "Type", default=1)
        odata_type = "#microsoft.graph.fileAttachment"
        if attachment_type == 5:
            odata_type = "#microsoft.graph.itemAttachment"
        elif attachment_type == 6:
            odata_type = "#microsoft.graph.referenceAttachment"

        return {
            "name": name,
            "contentType": content_type,
            "size": size,
            "isInline": is_inline,
            "odataType": odata_type,
        }

    def _serialize_recipients(self, recipients) -> list[dict[str, Any]]:
        result: list[dict[str, Any]] = []
        if recipients is None:
            return result
        for index in range(1, int(getattr(recipients, "Count", 0)) + 1):
            recipient = recipients.Item(index)
            address_entry = getattr(recipient, "AddressEntry", None)
            address = self._resolve_address(address_entry, fallback=self._safe_str(getattr(recipient, "Address", "")))
            recipient_type = int(getattr(recipient, "Type", 1))
            item = {"emailAddress": {"address": address}}
            if recipient_type == 1:
                result.append(item)
        return result

    def _resolve_address(self, address_entry, *, fallback: str = "") -> str:
        if address_entry is None:
            return fallback

        fallback = fallback or self._safe_str(getattr(address_entry, "Address", ""))
        address_type = self._safe_str(getattr(address_entry, "Type", ""))
        if fallback and "@" in fallback:
            return fallback

        if address_type == "EX":
            exchange_user = None
            try:
                exchange_user = address_entry.GetExchangeUser()
            except Exception:
                exchange_user = None
            if exchange_user is not None:
                smtp = self._safe_str(getattr(exchange_user, "PrimarySmtpAddress", ""))
                if smtp:
                    return smtp

        accessor_value = self._get_property(address_entry, "http://schemas.microsoft.com/mapi/proptag/0x39FE001E")
        return accessor_value or fallback

    @staticmethod
    def _safe_str(value: Any) -> str:
        if value is None:
            return ""
        return str(value).strip()

    @staticmethod
    def _safe_com_str(com_object: Any, attr_name: str, default: str = "") -> str:
        try:
            value = getattr(com_object, attr_name, default)
        except Exception:
            return default
        return OutlookDesktopClient._safe_str(value) or default

    @staticmethod
    def _safe_com_int(com_object: Any, attr_name: str, default: int = 0) -> int:
        try:
            value = getattr(com_object, attr_name, default)
        except Exception:
            return default
        try:
            return int(value or 0)
        except Exception:
            return default

    @staticmethod
    def _map_importance(value: int) -> str:
        if value >= 2:
            return "high"
        if value <= 0:
            return "low"
        return "normal"

    @staticmethod
    def _parse_categories(value: str) -> list[str]:
        return [item.strip() for item in value.split(",") if item.strip()]

    @staticmethod
    def _parse_since(since_iso: str | None) -> datetime | None:
        if not since_iso:
            return None
        return OutlookDesktopClient._parse_iso(since_iso)

    @staticmethod
    def _parse_iso(value: str) -> datetime:
        if value.endswith("Z"):
            value = value.replace("Z", "+00:00")
        return datetime.fromisoformat(value).astimezone(UTC)

    @staticmethod
    def _to_utc_iso(value: Any) -> str:
        if value is None:
            return ""
        if isinstance(value, datetime):
            if value.tzinfo is None:
                value = value.astimezone()
            return value.astimezone(UTC).replace(microsecond=0).isoformat().replace("+00:00", "Z")
        return ""

    @staticmethod
    def _get_property(com_object, property_tag: str) -> str:
        try:
            accessor = com_object.PropertyAccessor
            value = accessor.GetProperty(property_tag)
            return str(value).strip() if value is not None else ""
        except Exception:
            return ""

    @staticmethod
    def _compose_message_id(entry_id: str, store_id: str) -> str:
        if store_id:
            return f"{store_id}::{entry_id}"
        return entry_id

    @staticmethod
    def _split_message_id(value: str) -> tuple[str, str]:
        if "::" in value:
            store_id, entry_id = value.split("::", 1)
            return entry_id, store_id
        return value, ""

    def _message_key_for_item(self, item) -> str:
        store_id = self._safe_str(getattr(getattr(item, "Parent", None), "StoreID", ""))
        entry_id = self._safe_str(getattr(item, "EntryID", ""))
        internet_message_id = self._get_property(item, "http://schemas.microsoft.com/mapi/proptag/0x1035001F")
        subject = self._safe_str(getattr(item, "Subject", "")) or "message"
        seed = internet_message_id or self._compose_message_id(entry_id, store_id) or subject
        return MailStorage.message_key_from_seed(seed)

    @staticmethod
    def _normalize_folder_path(value: str) -> str:
        if not value:
            return ""
        normalized = value.replace("\\", "/").strip("/")
        while normalized.startswith("/"):
            normalized = normalized[1:]
        return normalized
