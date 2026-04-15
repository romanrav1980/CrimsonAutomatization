from __future__ import annotations

import base64
import hashlib
import json
import re
from dataclasses import dataclass
from datetime import UTC, datetime
from html.parser import HTMLParser
from pathlib import Path
from typing import Any


class _TextExtractor(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.parts: list[str] = []

    def handle_data(self, data: str) -> None:
        text = data.strip()
        if text:
            self.parts.append(text)

    def text(self) -> str:
        return "\n".join(self.parts)


def html_to_text(html: str) -> str:
    parser = _TextExtractor()
    parser.feed(html)
    return parser.text()


def slugify(value: str, *, fallback: str = "message") -> str:
    cleaned = re.sub(r"[^A-Za-z0-9]+", "-", value).strip("-").lower()
    return cleaned[:80] or fallback


def safe_path_segment(value: str, *, fallback: str = "_") -> str:
    cleaned = re.sub(r'[<>:"/\\\\|?*]+', "_", value).strip().strip(".")
    return cleaned or fallback


def classify_attachment_kind(name: str, content_type: str | None) -> str:
    lower_name = name.lower()
    lower_type = (content_type or "").lower()
    if lower_name.endswith((".xlsx", ".xls", ".xlsm", ".xlsb", ".csv")) or "spreadsheet" in lower_type or "excel" in lower_type:
        return "excel"
    if lower_name.endswith(".pdf") or lower_type == "application/pdf":
        return "pdf"
    if lower_name.endswith((".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tif", ".tiff", ".webp")) or lower_type.startswith("image/"):
        return "image"
    return "other"


def safe_json_dump(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def load_json(path: Path, *, default: Any) -> Any:
    if not path.exists():
        return default
    return json.loads(path.read_text(encoding="utf-8"))


@dataclass(slots=True)
class StoredMessage:
    key: str
    directory: Path


class MailStorage:
    def __init__(self, root: Path) -> None:
        self.root = root
        self.messages_dir = root / "messages"
        self.by_folder_dir = root / "by_folder"
        self.index_path = root / "index.json"
        self.state_path = root / "state.json"
        self.folders_path = root / "folders.json"

    def load_index(self) -> dict[str, Any]:
        return load_json(self.index_path, default={"messages": {}})

    def load_state(self) -> dict[str, Any]:
        return load_json(self.state_path, default={"last_sync_utc": None})

    def save_state(self, state: dict[str, Any]) -> None:
        safe_json_dump(self.state_path, state)

    def has_message(self, key: str) -> bool:
        index = self.load_index()
        return key in index.get("messages", {})

    def message_key(self, message: dict[str, Any]) -> str:
        seed = message.get("internetMessageId") or message.get("id") or message.get("subject", "message")
        return self.message_key_from_seed(seed)

    @staticmethod
    def message_key_from_seed(seed: str) -> str:
        return hashlib.sha1(seed.encode("utf-8")).hexdigest()[:16]

    def store_message(
        self,
        message: dict[str, Any],
        attachments: list[dict[str, Any]],
        *,
        save_attachments: bool,
        index: dict[str, Any] | None = None,
        rebuild_views: bool = True,
    ) -> StoredMessage:
        key = self.message_key(message)
        received = message.get("receivedDateTime") or datetime.now(UTC).isoformat()
        stamp = received.replace(":", "").replace("-", "")
        subject = message.get("subject") or "no-subject"
        folder_name = f"{stamp}__{slugify(subject)}__{key}"
        target_dir = self.messages_dir / folder_name
        target_dir.mkdir(parents=True, exist_ok=True)

        body = message.get("body") or {}
        body_html = body.get("content", "") if body.get("contentType") == "html" else ""
        body_text = body.get("content", "") if body.get("contentType") == "text" else ""
        if body_html and not body_text:
            body_text = html_to_text(body_html)
        if not body_text:
            body_text = message.get("bodyPreview", "")

        normalized = self._normalize_message(message, key=key, attachment_count=len(attachments))
        safe_json_dump(target_dir / "message.json", normalized)

        if body_html:
            (target_dir / "body.html").write_text(body_html, encoding="utf-8")
        (target_dir / "body.txt").write_text(body_text, encoding="utf-8")
        (target_dir / "source.md").write_text(self._source_markdown(normalized, body_text), encoding="utf-8")

        attachment_manifest: list[dict[str, Any]] = []
        if attachments:
            attachment_dir = target_dir / "attachments"
            attachment_dir.mkdir(exist_ok=True)
            for attachment in attachments:
                attachment_manifest.append(self._store_attachment(attachment_dir, attachment, save_attachments=save_attachments))
            safe_json_dump(target_dir / "attachments.json", attachment_manifest)

        index = index if index is not None else self.load_index()
        index.setdefault("messages", {})[key] = self._build_index_entry(
            normalized=normalized,
            target_dir=target_dir,
        )
        if rebuild_views:
            self.persist_index_and_views(index)
        return StoredMessage(key=key, directory=target_dir)

    def refresh_existing_message_metadata(
        self,
        message: dict[str, Any],
        *,
        index: dict[str, Any] | None = None,
        rebuild_views: bool = True,
    ) -> bool:
        key = self.message_key(message)
        index = index if index is not None else self.load_index()
        existing = index.get("messages", {}).get(key)
        if not existing:
            return False

        target_dir_value = existing.get("path")
        if not target_dir_value:
            return False

        target_dir = Path(target_dir_value)
        if not target_dir.is_absolute():
            target_dir = Path(target_dir_value)

        message_json_path = target_dir / "message.json"
        if not message_json_path.exists():
            return False

        payload = load_json(message_json_path, default={})
        attachment_count = int(payload.get("attachmentCount") or existing.get("attachmentCount") or 0)
        normalized = self._normalize_message(message, key=key, attachment_count=attachment_count)
        safe_json_dump(message_json_path, normalized)

        body_path = target_dir / "body.txt"
        body_text = body_path.read_text(encoding="utf-8") if body_path.exists() else (normalized.get("bodyPreview") or "")
        (target_dir / "source.md").write_text(self._source_markdown(normalized, body_text), encoding="utf-8")

        index.setdefault("messages", {})[key] = self._build_index_entry(
            normalized=normalized,
            target_dir=target_dir,
        )
        if rebuild_views:
            self.persist_index_and_views(index)
        return True

    def persist_index_and_views(self, index: dict[str, Any]) -> None:
        self.save_index(index)
        self.rebuild_views(index)

    def save_index(self, index: dict[str, Any]) -> None:
        safe_json_dump(self.index_path, index)

    def rebuild_views(self, index: dict[str, Any]) -> None:
        self._rebuild_folder_catalog(index)
        self._rebuild_folder_mirror(index)

    def _build_index_entry(self, *, normalized: dict[str, Any], target_dir: Path) -> dict[str, Any]:
        return {
            "subject": normalized["subject"],
            "receivedDateTime": normalized["receivedDateTime"],
            "sender": normalized["sender"],
            "path": str(target_dir).replace("\\", "/"),
            "hasAttachments": normalized["hasAttachments"],
            "attachmentCount": normalized["attachmentCount"],
            "sourceFolderName": normalized.get("sourceFolderName") or "",
            "sourceFolderPath": normalized.get("sourceFolderPath") or "",
            "sourceStoreName": normalized.get("sourceStoreName") or "",
        }
 
    def _rebuild_folder_catalog(self, index: dict[str, Any]) -> None:
        folders: dict[str, dict[str, Any]] = {}
        for key, entry in index.get("messages", {}).items():
            folder_path = str(entry.get("sourceFolderPath") or "").strip()
            if not folder_path:
                continue

            folder_name = str(entry.get("sourceFolderName") or folder_path.rsplit("/", 1)[-1]).strip()
            store_name = str(entry.get("sourceStoreName") or "").strip()
            bucket = folders.setdefault(
                folder_path,
                {
                    "folderName": folder_name,
                    "folderPath": folder_path,
                    "storeName": store_name,
                    "messageCount": 0,
                    "latestReceivedUtc": "",
                    "sampleMessageKeys": [],
                },
            )
            bucket["messageCount"] += 1
            received = str(entry.get("receivedDateTime") or "")
            if received and (not bucket["latestReceivedUtc"] or received > bucket["latestReceivedUtc"]):
                bucket["latestReceivedUtc"] = received
            if len(bucket["sampleMessageKeys"]) < 5:
                bucket["sampleMessageKeys"].append(key)

        payload = {
            "generatedUtc": datetime.now(UTC).replace(microsecond=0).isoformat().replace("+00:00", "Z"),
            "folders": [folders[path] for path in sorted(folders.keys())],
        }
        safe_json_dump(self.folders_path, payload)

    def _rebuild_folder_mirror(self, index: dict[str, Any]) -> None:
        if self.by_folder_dir.exists():
            for child in self.by_folder_dir.iterdir():
                if child.is_dir():
                    for nested in sorted(child.rglob("*"), reverse=True):
                        if nested.is_file():
                            nested.unlink()
                        elif nested.is_dir():
                            try:
                                nested.rmdir()
                            except OSError:
                                pass
                    try:
                        child.rmdir()
                    except OSError:
                        pass
                else:
                    child.unlink()
        self.by_folder_dir.mkdir(parents=True, exist_ok=True)

        folders: dict[str, list[dict[str, Any]]] = {}
        folder_meta: dict[str, dict[str, Any]] = {}

        for key, entry in index.get("messages", {}).items():
            folder_path = str(entry.get("sourceFolderPath") or "").strip()
            if not folder_path:
                folder_path = "_unknown"

            messages = folders.setdefault(folder_path, [])
            messages.append(
                {
                    "key": key,
                    "subject": entry.get("subject") or "",
                    "receivedDateTime": entry.get("receivedDateTime") or "",
                    "sender": entry.get("sender") or "",
                    "hasAttachments": bool(entry.get("hasAttachments")),
                    "attachmentCount": int(entry.get("attachmentCount") or 0),
                    "canonicalPath": entry.get("path") or "",
                }
            )
            folder_meta.setdefault(
                folder_path,
                {
                    "folderName": entry.get("sourceFolderName") or folder_path.rsplit("/", 1)[-1],
                    "folderPath": folder_path,
                    "storeName": entry.get("sourceStoreName") or "",
                },
            )

        for folder_path, messages in folders.items():
            folder_dir = self.by_folder_dir
            for segment in folder_path.split("/"):
                folder_dir = folder_dir / safe_path_segment(segment, fallback="_")
            folder_dir.mkdir(parents=True, exist_ok=True)
            payload = {
                "generatedUtc": datetime.now(UTC).replace(microsecond=0).isoformat().replace("+00:00", "Z"),
                "folderName": folder_meta[folder_path]["folderName"],
                "folderPath": folder_path,
                "storeName": folder_meta[folder_path]["storeName"],
                "messageCount": len(messages),
                "messages": sorted(messages, key=lambda item: item["receivedDateTime"], reverse=True),
            }
            safe_json_dump(folder_dir / "folder.json", payload)

    def _normalize_message(self, message: dict[str, Any], *, key: str, attachment_count: int) -> dict[str, Any]:
        from_address = ((message.get("from") or {}).get("emailAddress") or {}).get("address")
        sender_address = ((message.get("sender") or {}).get("emailAddress") or {}).get("address")
        return {
            "key": key,
            "id": message.get("id"),
            "internetMessageId": message.get("internetMessageId"),
            "conversationId": message.get("conversationId"),
            "subject": message.get("subject"),
            "sender": sender_address or from_address,
            "from": message.get("from"),
            "toRecipients": message.get("toRecipients") or [],
            "ccRecipients": message.get("ccRecipients") or [],
            "bccRecipients": message.get("bccRecipients") or [],
            "receivedDateTime": message.get("receivedDateTime"),
            "sentDateTime": message.get("sentDateTime"),
            "hasAttachments": bool(message.get("hasAttachments")),
            "attachmentCount": attachment_count,
            "importance": message.get("importance"),
            "isRead": message.get("isRead"),
            "categories": message.get("categories") or [],
            "bodyPreview": message.get("bodyPreview"),
            "webLink": message.get("webLink"),
            "sourceFolderName": message.get("sourceFolderName") or "",
            "sourceFolderPath": message.get("sourceFolderPath") or "",
            "sourceStoreName": message.get("sourceStoreName") or "",
        }

    def _store_attachment(self, attachment_dir: Path, attachment: dict[str, Any], *, save_attachments: bool) -> dict[str, Any]:
        attachment_type = attachment.get("@odata.type", "")
        name = attachment.get("name") or "attachment"
        content_type = attachment.get("contentType")
        attachment_kind = classify_attachment_kind(name, content_type)
        result = {
            "id": attachment.get("id"),
            "name": name,
            "originalName": name,
            "contentType": content_type,
            "size": attachment.get("size"),
            "type": attachment_type,
            "kind": attachment_kind,
            "isInline": attachment.get("isInline"),
            "saved": False,
        }

        if save_attachments and attachment_type.endswith("fileAttachment") and attachment.get("contentBytes"):
            content_bytes = base64.b64decode(attachment["contentBytes"])
            sha256 = hashlib.sha256(content_bytes).hexdigest()
            extension = Path(name).suffix
            stem = Path(name).stem
            safe_name = "{0}__{1}{2}".format(
                slugify(stem, fallback="attachment"),
                sha256[:10],
                extension if extension else "",
            )
            kind_dir = attachment_dir / attachment_kind
            kind_dir.mkdir(exist_ok=True)
            target_path = kind_dir / safe_name
            target_path.write_bytes(content_bytes)
            result["saved"] = True
            result["path"] = str(target_path).replace("\\", "/")
            result["storedName"] = safe_name
            result["sha256"] = sha256
            result["extension"] = extension.lower()
        else:
            metadata_path = attachment_dir / f"{slugify(name, fallback='attachment')}.json"
            safe_json_dump(metadata_path, attachment)
            result["path"] = str(metadata_path).replace("\\", "/")
            result["storedName"] = metadata_path.name
            result["sha256"] = ""
            result["extension"] = Path(name).suffix.lower()

        return result

    def _source_markdown(self, message: dict[str, Any], body_text: str) -> str:
        lines = [
            f"# {message.get('subject') or 'Untitled message'}",
            "",
            f"- Key: `{message.get('key')}`",
            f"- Received: `{message.get('receivedDateTime')}`",
            f"- Sender: `{message.get('sender') or 'unknown'}`",
            f"- Outlook folder: `{message.get('sourceFolderPath') or message.get('sourceFolderName') or 'unknown'}`",
            f"- Outlook store: `{message.get('sourceStoreName') or 'unknown'}`",
            f"- Attachment count: `{message.get('attachmentCount', 0)}`",
            f"- Web link: {message.get('webLink') or 'n/a'}",
            "",
            "## Body",
            "",
            body_text.strip() or "_No body content_",
            "",
        ]
        return "\n".join(lines)
