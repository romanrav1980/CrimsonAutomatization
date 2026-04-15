from __future__ import annotations

from dataclasses import dataclass

from services.mail_processing.models import ClassificationResult, RawMailArtifact


@dataclass(slots=True)
class ClassifierContext:
    internal_domains: set[str]


class HeuristicMailClassifier:
    def __init__(self, context: ClassifierContext) -> None:
        self.context = context

    def classify(self, artifact: RawMailArtifact) -> ClassificationResult:
        text = f"{artifact.subject}\n{artifact.body_preview}\n{artifact.body_text}".lower()
        labels: list[str] = []
        if artifact.has_attachments:
            labels.append("contains_attachment")
        if self._is_external_sender(artifact.sender_domain):
            labels.append("customer_facing")
        else:
            labels.append("internal_only")

        scores = {
            "incident": self._score(text, ["incident", "outage", "failure", "error", "down", "broken", "urgent"]),
            "approval": self._score(text, ["approve", "approval", "sign off", "signoff", "authorize", "authorization"]),
            "document_review": self._score(text, ["review", "document", "contract", "proposal", "invoice", "attachment"]),
            "request": self._score(text, ["request", "please", "need", "could you", "can you", "help", "action required"]),
            "follow_up": self._score(text, ["follow up", "following up", "checking in", "reminder", "ping"]),
            "informational": self._score(text, ["fyi", "for your information", "notice", "update", "announcement"]),
            "noise": self._score(text, ["unsubscribe", "newsletter", "marketing", "promotion", "promo"]),
        }

        if artifact.has_attachments:
            scores["document_review"] += 1

        process_type, top_score = max(scores.items(), key=lambda item: item[1])
        if top_score <= 0:
            process_type = "informational"
            top_score = 1

        confidence = min(0.98, 0.46 + 0.12 * top_score)
        urgency = self._detect_urgency(text, artifact)
        needs_action = process_type in {"incident", "approval", "document_review", "request", "follow_up"}

        if process_type == "noise":
            needs_action = False
            labels.append("auto_archive_candidate")

        if process_type in {"incident", "approval", "request", "follow_up"}:
            labels.append("needs_action")

        if urgency == "high":
            labels.append("high_priority")

        if confidence < 0.58:
            labels.append("requires_human_review")

        return ClassificationResult(
            process_type=process_type,
            confidence=round(confidence, 2),
            needs_action=needs_action,
            urgency=urgency,
            labels=sorted(set(labels)),
        )

    def _detect_urgency(self, text: str, artifact: RawMailArtifact) -> str:
        high_terms = ["urgent", "asap", "today", "eod", "immediately", "outage", "down", "critical"]
        medium_terms = ["follow up", "request", "approval", "review", "tomorrow", "soon"]
        if any(term in text for term in high_terms) or str(artifact.subject).isupper():
            return "high"
        if any(term in text for term in medium_terms):
            return "medium"
        return "low"

    def _is_external_sender(self, domain: str) -> bool:
        if not domain:
            return True
        if not self.context.internal_domains:
            return True
        return domain.lower() not in self.context.internal_domains

    @staticmethod
    def _score(text: str, terms: list[str]) -> int:
        return sum(1 for term in terms if term in text)
