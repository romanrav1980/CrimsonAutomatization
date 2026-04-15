from __future__ import annotations

import json
import sqlite3
from datetime import UTC, datetime
from pathlib import Path

from services.mail_processing.models import ProcessedMailRecord


SCHEMA = """
CREATE TABLE IF NOT EXISTS mail_items (
    message_key TEXT PRIMARY KEY,
    subject TEXT NOT NULL,
    sender TEXT,
    sender_domain TEXT,
    source_folder_name TEXT,
    source_folder_path TEXT,
    source_store_name TEXT,
    received_utc TEXT,
    body_preview TEXT,
    body_path TEXT,
    source_path TEXT,
    raw_message_path TEXT,
    attachment_count INTEGER NOT NULL DEFAULT 0,
    has_attachments INTEGER NOT NULL DEFAULT 0,
    analyzed_attachment_count INTEGER NOT NULL DEFAULT 0,
    attachment_summary TEXT,
    attachment_kinds_json TEXT NOT NULL DEFAULT '[]',
    attachment_analysis_path TEXT,
    attachment_analysis_json TEXT NOT NULL DEFAULT '[]',
    process_type TEXT NOT NULL,
    confidence REAL NOT NULL,
    needs_action INTEGER NOT NULL DEFAULT 0,
    urgency TEXT NOT NULL,
    labels_json TEXT NOT NULL,
    decision_mode TEXT NOT NULL,
    recommended_action TEXT NOT NULL,
    decision_reason TEXT NOT NULL,
    status TEXT NOT NULL,
    service_level_state TEXT NOT NULL,
    owner TEXT,
    queue_name TEXT,
    web_link TEXT,
    created_utc TEXT NOT NULL,
    updated_utc TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS mail_audit_events (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    message_key TEXT NOT NULL,
    event_type TEXT NOT NULL,
    payload_json TEXT NOT NULL,
    created_utc TEXT NOT NULL
);
"""

REQUIRED_MAIL_ITEM_COLUMNS = {
    "source_folder_name": "ALTER TABLE mail_items ADD COLUMN source_folder_name TEXT;",
    "source_folder_path": "ALTER TABLE mail_items ADD COLUMN source_folder_path TEXT;",
    "source_store_name": "ALTER TABLE mail_items ADD COLUMN source_store_name TEXT;",
    "analyzed_attachment_count": "ALTER TABLE mail_items ADD COLUMN analyzed_attachment_count INTEGER NOT NULL DEFAULT 0;",
    "attachment_summary": "ALTER TABLE mail_items ADD COLUMN attachment_summary TEXT;",
    "attachment_kinds_json": "ALTER TABLE mail_items ADD COLUMN attachment_kinds_json TEXT NOT NULL DEFAULT '[]';",
    "attachment_analysis_path": "ALTER TABLE mail_items ADD COLUMN attachment_analysis_path TEXT;",
    "attachment_analysis_json": "ALTER TABLE mail_items ADD COLUMN attachment_analysis_json TEXT NOT NULL DEFAULT '[]';",
}


class MailMvpRepository:
    def __init__(self, db_path: Path) -> None:
        self.db_path = db_path
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        with self._connect() as connection:
            connection.executescript(SCHEMA)
            self._ensure_mail_item_columns(connection)

    def upsert_mail_item(self, record: ProcessedMailRecord) -> None:
        now = self._utc_now()
        with self._connect() as connection:
            connection.row_factory = sqlite3.Row
            existing = connection.execute(
                "SELECT created_utc, status, decision_mode, owner FROM mail_items WHERE message_key = ?",
                (record.message_key,),
            ).fetchone()
            created_utc = existing["created_utc"] if existing else now
            status_value = record.status
            decision_mode_value = record.decision_mode
            owner_value = record.owner

            if existing is not None:
                existing_status = (existing["status"] or "").strip()
                existing_decision_mode = (existing["decision_mode"] or "").strip()
                existing_owner = (existing["owner"] or "").strip()

                if existing_owner:
                    owner_value = existing_owner

                if existing_status in {"approved", "archived", "manual_review"}:
                    status_value = existing_status
                    decision_mode_value = existing_decision_mode or decision_mode_value

            connection.execute(
                """
                INSERT INTO mail_items (
                    message_key, subject, sender, sender_domain, source_folder_name, source_folder_path, source_store_name,
                    received_utc, body_preview, body_path, source_path,
                    raw_message_path, attachment_count, has_attachments, analyzed_attachment_count, attachment_summary,
                    attachment_kinds_json, attachment_analysis_path, attachment_analysis_json,
                    process_type, confidence, needs_action,
                    urgency, labels_json, decision_mode, recommended_action, decision_reason, status,
                    service_level_state, owner, queue_name, web_link, created_utc, updated_utc
                ) VALUES (
                    ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?
                )
                ON CONFLICT(message_key) DO UPDATE SET
                    subject = excluded.subject,
                    sender = excluded.sender,
                    sender_domain = excluded.sender_domain,
                    source_folder_name = excluded.source_folder_name,
                    source_folder_path = excluded.source_folder_path,
                    source_store_name = excluded.source_store_name,
                    received_utc = excluded.received_utc,
                    body_preview = excluded.body_preview,
                    body_path = excluded.body_path,
                    source_path = excluded.source_path,
                    raw_message_path = excluded.raw_message_path,
                    attachment_count = excluded.attachment_count,
                    has_attachments = excluded.has_attachments,
                    analyzed_attachment_count = excluded.analyzed_attachment_count,
                    attachment_summary = excluded.attachment_summary,
                    attachment_kinds_json = excluded.attachment_kinds_json,
                    attachment_analysis_path = excluded.attachment_analysis_path,
                    attachment_analysis_json = excluded.attachment_analysis_json,
                    process_type = excluded.process_type,
                    confidence = excluded.confidence,
                    needs_action = excluded.needs_action,
                    urgency = excluded.urgency,
                    labels_json = excluded.labels_json,
                    decision_mode = excluded.decision_mode,
                    recommended_action = excluded.recommended_action,
                    decision_reason = excluded.decision_reason,
                    status = excluded.status,
                    service_level_state = excluded.service_level_state,
                    owner = excluded.owner,
                    queue_name = excluded.queue_name,
                    web_link = excluded.web_link,
                    updated_utc = excluded.updated_utc
                """,
                (
                    record.message_key,
                    record.subject,
                    record.sender,
                    record.sender_domain,
                    record.source_folder_name,
                    record.source_folder_path,
                    record.source_store_name,
                    record.received_utc,
                    record.body_preview,
                    record.body_path,
                    record.source_path,
                    record.raw_message_path,
                    int(record.attachment_count),
                    int(record.has_attachments),
                    int(record.analyzed_attachment_count),
                    record.attachment_summary,
                    json.dumps(record.attachment_kinds, ensure_ascii=False),
                    record.attachment_analysis_path,
                    json.dumps(record.attachment_analysis, ensure_ascii=False),
                    record.process_type,
                    float(record.confidence),
                    int(record.needs_action),
                    record.urgency,
                    json.dumps(record.labels),
                    decision_mode_value,
                    record.recommended_action,
                    record.decision_reason,
                    status_value,
                    record.service_level_state,
                    owner_value,
                    record.queue_name,
                    record.web_link,
                    created_utc,
                    now,
                ),
            )
            connection.execute(
                """
                INSERT INTO mail_audit_events (message_key, event_type, payload_json, created_utc)
                VALUES (?, ?, ?, ?)
                """,
                (
                    record.message_key,
                    "processed",
                    json.dumps(
                        {
                            "processType": record.process_type,
                            "decisionMode": record.decision_mode,
                            "recommendedAction": record.recommended_action,
                            "status": record.status,
                        }
                    ),
                    now,
                ),
            )

    def fetch_dashboard_snapshot(self) -> dict:
        with self._connect() as connection:
            connection.row_factory = sqlite3.Row
            totals = connection.execute(
                """
                SELECT
                    COUNT(*) AS total_messages,
                    SUM(CASE WHEN status = 'needs_decision' THEN 1 ELSE 0 END) AS needs_decision,
                    SUM(CASE WHEN decision_mode = 'Auto' THEN 1 ELSE 0 END) AS auto_ready,
                    SUM(CASE WHEN decision_mode = 'Manual' THEN 1 ELSE 0 END) AS manual_review,
                    SUM(CASE WHEN urgency = 'high' THEN 1 ELSE 0 END) AS high_urgency
                FROM mail_items
                """
            ).fetchone()
            needs_decision_rows = connection.execute(
                """
                SELECT *
                FROM mail_items
                WHERE status = 'needs_decision'
                ORDER BY
                    CASE urgency WHEN 'high' THEN 0 WHEN 'medium' THEN 1 ELSE 2 END,
                    received_utc DESC
                LIMIT 100
                """
            ).fetchall()

        summary = {
            "totalMessages": int(totals["total_messages"] or 0),
            "needsDecision": int(totals["needs_decision"] or 0),
            "autoReady": int(totals["auto_ready"] or 0),
            "manualReview": int(totals["manual_review"] or 0),
            "highUrgency": int(totals["high_urgency"] or 0),
        }
        needs_decision = []
        for row in needs_decision_rows:
            needs_decision.append(
                {
                    "messageKey": row["message_key"],
                    "subject": row["subject"],
                    "sender": row["sender"],
                    "senderDomain": row["sender_domain"],
                    "sourceFolderName": row["source_folder_name"] or "",
                    "sourceFolderPath": row["source_folder_path"] or "",
                    "sourceStoreName": row["source_store_name"] or "",
                    "receivedUtc": row["received_utc"],
                    "processType": row["process_type"],
                    "confidence": round(float(row["confidence"] or 0), 2),
                    "urgency": row["urgency"],
                    "decisionMode": row["decision_mode"],
                    "recommendedAction": row["recommended_action"],
                    "decisionReason": row["decision_reason"],
                    "status": row["status"],
                    "serviceLevelState": row["service_level_state"],
                    "owner": row["owner"] or "",
                    "attachmentCount": int(row["attachment_count"] or 0),
                    "analyzedAttachmentCount": int(row["analyzed_attachment_count"] or 0),
                    "attachmentSummary": row["attachment_summary"] or "",
                    "attachmentKinds": json.loads(row["attachment_kinds_json"] or "[]"),
                    "attachmentAnalysisPath": row["attachment_analysis_path"] or "",
                    "attachmentAnalysis": json.loads(row["attachment_analysis_json"] or "[]"),
                    "labels": json.loads(row["labels_json"] or "[]"),
                    "bodyPath": row["body_path"],
                    "sourcePath": row["source_path"],
                    "rawMessagePath": row["raw_message_path"],
                    "webLink": row["web_link"],
                }
            )

        return {
            "generatedUtc": self._utc_now(),
            "summary": summary,
            "needsDecision": needs_decision,
        }

    def export_dashboard_snapshot(self, target_path: Path) -> dict:
        snapshot = self.fetch_dashboard_snapshot()
        target_path.parent.mkdir(parents=True, exist_ok=True)
        target_path.write_text(json.dumps(snapshot, ensure_ascii=False, indent=2), encoding="utf-8")
        return snapshot

    def apply_operator_action(
        self,
        *,
        message_key: str,
        action: str,
        actor: str,
        owner: str | None = None,
        notes: str | None = None,
    ) -> dict:
        normalized_action = action.strip().lower()
        owner_value = (owner or "").strip() or None
        notes_value = (notes or "").strip() or None
        now = self._utc_now()

        with self._connect() as connection:
            connection.row_factory = sqlite3.Row
            row = connection.execute(
                "SELECT * FROM mail_items WHERE message_key = ?",
                (message_key,),
            ).fetchone()

            if row is None:
                raise KeyError(f"Mail item '{message_key}' was not found.")

            updates: dict[str, object] = {
                "updated_utc": now,
            }

            if owner_value is not None:
                updates["owner"] = owner_value

            if normalized_action == "approve":
                updates["status"] = "approved"
            elif normalized_action == "archive":
                updates["status"] = "archived"
            elif normalized_action == "manual":
                updates["decision_mode"] = "Manual"
                updates["status"] = "manual_review"
            elif normalized_action == "assign_owner":
                if owner_value is None:
                    raise ValueError("Assign owner action requires a non-empty owner value.")
            else:
                raise ValueError(f"Unsupported operator action: {action}")

            assignments = ", ".join(f"{column} = ?" for column in updates.keys())
            connection.execute(
                f"UPDATE mail_items SET {assignments} WHERE message_key = ?",
                (*updates.values(), message_key),
            )

            payload = {
                "action": normalized_action,
                "actor": actor,
                "owner": owner_value,
                "notes": notes_value,
                "previousStatus": row["status"],
                "newStatus": updates.get("status", row["status"]),
                "previousOwner": row["owner"],
                "newOwner": updates.get("owner", row["owner"]),
            }
            connection.execute(
                """
                INSERT INTO mail_audit_events (message_key, event_type, payload_json, created_utc)
                VALUES (?, ?, ?, ?)
                """,
                (
                    message_key,
                    "operator_action",
                    json.dumps(payload, ensure_ascii=False),
                    now,
                ),
            )

            updated = connection.execute(
                "SELECT * FROM mail_items WHERE message_key = ?",
                (message_key,),
            ).fetchone()

        return {
            "messageKey": message_key,
            "action": normalized_action,
            "status": updated["status"] if updated is not None else "",
            "owner": updated["owner"] if updated is not None else "",
            "updatedUtc": now,
        }

    def _connect(self) -> sqlite3.Connection:
        connection = sqlite3.connect(self.db_path)
        connection.execute("PRAGMA journal_mode=WAL;")
        return connection

    @staticmethod
    def _ensure_mail_item_columns(connection: sqlite3.Connection) -> None:
        rows = connection.execute("PRAGMA table_info(mail_items)").fetchall()
        existing_columns = {row[1] for row in rows}
        for column_name, statement in REQUIRED_MAIL_ITEM_COLUMNS.items():
            if column_name not in existing_columns:
                connection.execute(statement)

    @staticmethod
    def _utc_now() -> str:
        return datetime.now(UTC).replace(microsecond=0).isoformat().replace("+00:00", "Z")
