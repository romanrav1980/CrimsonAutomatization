from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path

from services.mail_processing.attachment_analysis import AttachmentAnalysisService
from services.mail_processing.classifier import ClassifierContext, HeuristicMailClassifier
from services.mail_processing.database import MailMvpRepository
from services.mail_processing.decisioning import DecisionEngine
from services.mail_processing.models import ProcessedMailRecord, RawMailArtifact


@dataclass(slots=True)
class MailProcessingResult:
    processed: int
    failed: int
    normalized_loaded: int
    attachments_analyzed: int
    classified: int
    needs_decision: int
    auto_ready: int
    manual_review: int
    exported_read_model_path: Path
    db_path: Path


class MailProcessingPipeline:
    def __init__(self, *, raw_root: Path, derived_root: Path, db_path: Path, read_model_path: Path, internal_domains: set[str]) -> None:
        self.raw_root = raw_root
        self.derived_root = derived_root
        self.db_path = db_path
        self.read_model_path = read_model_path
        self.repository = MailMvpRepository(db_path)
        self.classifier = HeuristicMailClassifier(ClassifierContext(internal_domains=internal_domains))
        self.decision_engine = DecisionEngine()
        self.attachment_analysis = AttachmentAnalysisService(raw_root=raw_root, derived_root=derived_root)

    def process(self) -> MailProcessingResult:
        processed = 0
        failed = 0
        self._emit_stage("normalize", "running", "Loading raw mail artifacts")
        artifacts = self._load_artifacts()
        self._emit_stage("normalize", "done", f"Loaded {len(artifacts)} raw mail artifacts")
        self._emit_metric("normalized_loaded", len(artifacts), "Raw artifacts loaded")

        self._emit_stage("attachments", "running", f"Analyzing attachments for {len(artifacts)} mail artifacts")
        attachment_results: dict[str, dict] = {}
        for artifact in artifacts:
            try:
                analysis = self.attachment_analysis.analyze_message(artifact)
                attachment_results[artifact.message_key] = {
                    "analyzedAttachmentCount": analysis.analyzed_attachment_count,
                    "attachmentSummary": analysis.summary,
                    "attachmentKinds": analysis.attachment_kinds,
                    "attachmentAnalysisPath": analysis.analysis_path,
                    "attachmentAnalysis": analysis.attachments,
                }
            except Exception as exc:
                failed += 1
                print(f"Failed to analyze attachments for {artifact.message_key}: {exc}")
        attachments_analyzed = sum(item["analyzedAttachmentCount"] for item in attachment_results.values())
        self._emit_stage("attachments", "done", f"Analyzed {attachments_analyzed} attachments")
        self._emit_metric("attachments_analyzed", attachments_analyzed, "Attachments analyzed")

        self._emit_stage("classify", "running", f"Classifying {len(artifacts)} mail artifacts")
        classified_items = []
        for artifact in artifacts:
            try:
                classification = self.classifier.classify(artifact)
                classified_items.append((artifact, classification))
            except Exception as exc:
                failed += 1
                print(f"Failed to classify {artifact.message_key}: {exc}")
        self._emit_stage("classify", "done", f"Classified {len(classified_items)} mail artifacts")
        self._emit_metric("classified", len(classified_items), "Mail artifacts classified")

        self._emit_stage("decision", "running", f"Applying decision matrix to {len(classified_items)} items")
        for artifact, classification in classified_items:
            try:
                attachment_context = attachment_results.get(
                    artifact.message_key,
                    {
                        "analyzedAttachmentCount": 0,
                        "attachmentSummary": "No attachments to analyze." if artifact.attachment_count <= 0 else "Attachment analysis unavailable.",
                        "attachmentKinds": [],
                        "attachmentAnalysisPath": "",
                        "attachmentAnalysis": [],
                    },
                )
                decision = self.decision_engine.decide(artifact, classification)
                labels = list(classification.labels)
                for kind in attachment_context["attachmentKinds"]:
                    label = f"attachment:{kind}"
                    if label not in labels:
                        labels.append(label)
                record = ProcessedMailRecord(
                    message_key=artifact.message_key,
                    subject=artifact.subject,
                    sender=artifact.sender,
                    sender_domain=artifact.sender_domain,
                    source_folder_name=artifact.source_folder_name,
                    source_folder_path=artifact.source_folder_path,
                    source_store_name=artifact.source_store_name,
                    received_utc=artifact.received_utc,
                    body_preview=artifact.body_preview,
                    body_path=artifact.body_path,
                    source_path=artifact.source_path,
                    raw_message_path=artifact.raw_message_path,
                    attachment_count=artifact.attachment_count,
                    has_attachments=artifact.has_attachments,
                    web_link=artifact.web_link,
                    analyzed_attachment_count=attachment_context["analyzedAttachmentCount"],
                    attachment_summary=attachment_context["attachmentSummary"],
                    attachment_kinds=attachment_context["attachmentKinds"],
                    attachment_analysis_path=attachment_context["attachmentAnalysisPath"],
                    process_type=classification.process_type,
                    confidence=classification.confidence,
                    needs_action=classification.needs_action,
                    urgency=classification.urgency,
                    labels=labels,
                    decision_mode=decision.decision_mode,
                    recommended_action=decision.recommended_action,
                    decision_reason=decision.decision_reason,
                    status=decision.status,
                    service_level_state=decision.service_level_state,
                    attachment_analysis=attachment_context["attachmentAnalysis"],
                )
                self.repository.upsert_mail_item(record)
                processed += 1
            except Exception as exc:
                failed += 1
                print(f"Failed to process {artifact.message_key}: {exc}")
        self._emit_stage("decision", "done", f"Processed {processed} items, failed {failed}")
        self._emit_metric("processed", processed, "Processed items")
        self._emit_metric("failed", failed, "Failed items")

        snapshot = self.repository.fetch_dashboard_snapshot()
        summary = snapshot.get("summary", {})
        needs_decision = int(summary.get("needsDecision", 0))
        auto_ready = int(summary.get("autoReady", 0))
        manual_review = int(summary.get("manualReview", 0))
        self._emit_metric("needs_decision", needs_decision, "Items in Needs Decision queue")
        self._emit_metric("auto_ready", auto_ready, "Items marked Auto ready")
        self._emit_metric("manual_review", manual_review, "Items requiring manual review")
        self.read_model_path.parent.mkdir(parents=True, exist_ok=True)
        self.read_model_path.write_text(json.dumps(snapshot, ensure_ascii=False, indent=2), encoding="utf-8")
        return MailProcessingResult(
            processed=processed,
            failed=failed,
            normalized_loaded=len(artifacts),
            attachments_analyzed=attachments_analyzed,
            classified=len(classified_items),
            needs_decision=needs_decision,
            auto_ready=auto_ready,
            manual_review=manual_review,
            exported_read_model_path=self.read_model_path,
            db_path=self.db_path,
        )

    def _load_artifacts(self) -> list[RawMailArtifact]:
        artifacts: list[RawMailArtifact] = []
        messages_root = self.raw_root / "messages"
        if not messages_root.exists():
            return artifacts

        for message_dir in sorted(messages_root.iterdir()):
            if not message_dir.is_dir():
                continue
            try:
                artifact = self._load_artifact(message_dir)
                if artifact is not None:
                    artifacts.append(artifact)
            except Exception as exc:
                print(f"Failed to load raw artifact {message_dir}: {exc}")
        return artifacts

    def _load_artifact(self, message_dir: Path) -> RawMailArtifact | None:
        message_json = message_dir / "message.json"
        body_txt = message_dir / "body.txt"
        source_md = message_dir / "source.md"
        if not message_json.exists():
            return None

        payload = json.loads(message_json.read_text(encoding="utf-8"))
        sender = payload.get("sender") or "unknown"
        sender_domain = sender.split("@", 1)[1].lower() if "@" in sender else ""
        body_text = body_txt.read_text(encoding="utf-8") if body_txt.exists() else ""
        return RawMailArtifact(
            message_key=payload.get("key") or message_dir.name,
            subject=payload.get("subject") or "Untitled message",
            sender=sender,
            sender_domain=sender_domain,
            source_folder_name=payload.get("sourceFolderName") or "",
            source_folder_path=payload.get("sourceFolderPath") or "",
            source_store_name=payload.get("sourceStoreName") or "",
            received_utc=payload.get("receivedDateTime") or "",
            body_preview=payload.get("bodyPreview") or body_text[:240],
            body_text=body_text,
            body_path=str(body_txt).replace("\\", "/"),
            source_path=str(source_md).replace("\\", "/"),
            raw_message_path=str(message_json).replace("\\", "/"),
            message_dir_path=str(message_dir.resolve()).replace("\\", "/"),
            attachments_manifest_path=str((message_dir / "attachments.json").resolve()).replace("\\", "/"),
            attachment_count=int(payload.get("attachmentCount") or 0),
            has_attachments=bool(payload.get("hasAttachments")),
            web_link=payload.get("webLink") or "",
        )

    @staticmethod
    def _emit_stage(stage: str, state: str, detail: str) -> None:
        print(f"[[STAGE|{stage}|{state}|{detail}]]")

    @staticmethod
    def _emit_metric(name: str, value: int, label: str) -> None:
        print(f"[[METRIC|{name}|{value}|{label}]]")
