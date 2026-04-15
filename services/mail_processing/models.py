from __future__ import annotations

from dataclasses import dataclass, field


@dataclass(slots=True)
class RawMailArtifact:
    message_key: str
    subject: str
    sender: str
    sender_domain: str
    source_folder_name: str
    source_folder_path: str
    source_store_name: str
    received_utc: str
    body_preview: str
    body_text: str
    body_path: str
    source_path: str
    raw_message_path: str
    message_dir_path: str
    attachments_manifest_path: str
    attachment_count: int
    has_attachments: bool
    web_link: str


@dataclass(slots=True)
class ClassificationResult:
    process_type: str
    confidence: float
    needs_action: bool
    urgency: str
    labels: list[str] = field(default_factory=list)


@dataclass(slots=True)
class DecisionResult:
    decision_mode: str
    recommended_action: str
    decision_reason: str
    status: str
    service_level_state: str


@dataclass(slots=True)
class ProcessedMailRecord:
    message_key: str
    subject: str
    sender: str
    sender_domain: str
    source_folder_name: str
    source_folder_path: str
    source_store_name: str
    received_utc: str
    body_preview: str
    body_path: str
    source_path: str
    raw_message_path: str
    attachment_count: int
    has_attachments: bool
    web_link: str
    analyzed_attachment_count: int
    attachment_summary: str
    attachment_kinds: list[str]
    attachment_analysis_path: str
    process_type: str
    confidence: float
    needs_action: bool
    urgency: str
    labels: list[str]
    decision_mode: str
    recommended_action: str
    decision_reason: str
    status: str
    service_level_state: str
    attachment_analysis: list[dict[str, object]] = field(default_factory=list)
    owner: str = "unassigned"
    queue_name: str = "triage"
