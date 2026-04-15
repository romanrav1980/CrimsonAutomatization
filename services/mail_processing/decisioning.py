from __future__ import annotations

from services.mail_processing.models import ClassificationResult, DecisionResult, RawMailArtifact


class DecisionEngine:
    def decide(self, artifact: RawMailArtifact, classification: ClassificationResult) -> DecisionResult:
        text = f"{artifact.subject}\n{artifact.body_preview}\n{artifact.body_text}".lower()

        if "legal" in text or "payment" in text or "contract" in text:
            return DecisionResult(
                decision_mode="Manual",
                recommended_action="Route to human review",
                decision_reason="Sensitive financial or legal wording detected.",
                status="needs_decision",
                service_level_state="at_risk" if classification.urgency == "high" else "on_track",
            )

        if classification.confidence < 0.58:
            return DecisionResult(
                decision_mode="Manual",
                recommended_action="Review classification manually",
                decision_reason="Low confidence classification should not auto-route.",
                status="needs_decision",
                service_level_state="on_track",
            )

        if classification.process_type == "noise":
            return DecisionResult(
                decision_mode="Auto",
                recommended_action="Archive as noise",
                decision_reason="Message looks like newsletter or low-value operational noise.",
                status="auto_ready",
                service_level_state="on_track",
            )

        if classification.process_type == "informational" and not classification.needs_action:
            return DecisionResult(
                decision_mode="Auto",
                recommended_action="Archive as informational",
                decision_reason="Informational message without explicit action markers.",
                status="auto_ready",
                service_level_state="on_track",
            )

        if classification.process_type == "incident":
            return DecisionResult(
                decision_mode="Suggest",
                recommended_action="Create or update incident case",
                decision_reason="Incident-like message with possible SLA impact.",
                status="needs_decision",
                service_level_state="at_risk" if classification.urgency == "high" else "on_track",
            )

        if classification.process_type == "approval":
            return DecisionResult(
                decision_mode="Manual",
                recommended_action="Review approval request",
                decision_reason="Approval requests should stay human-controlled in the MVP.",
                status="needs_decision",
                service_level_state="on_track",
            )

        if classification.process_type == "document_review":
            return DecisionResult(
                decision_mode="Suggest",
                recommended_action="Create review task and assign owner",
                decision_reason="Document-bearing message likely needs review workflow.",
                status="needs_decision",
                service_level_state="on_track",
            )

        if classification.process_type == "follow_up":
            return DecisionResult(
                decision_mode="Suggest",
                recommended_action="Attach to existing case or create follow-up task",
                decision_reason="Follow-up wording indicates pending work or waiting state.",
                status="needs_decision",
                service_level_state="at_risk" if classification.urgency != "low" else "on_track",
            )

        return DecisionResult(
            decision_mode="Suggest",
            recommended_action="Create routed task",
            decision_reason="Request-like message should be reviewed before routing.",
            status="needs_decision",
            service_level_state="at_risk" if classification.urgency == "high" else "on_track",
        )
