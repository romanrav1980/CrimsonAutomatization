# Mail Triage And Decisioning

## Purpose

This document defines how the project should regularly read Outlook mail, store it in `raw/`, analyze incoming messages, and decide what action should happen next.

The goal is to reduce manual work without losing control over business-critical decisions.

## Core principle

Mail must be handled in two layers:

1. `raw/` keeps the original message and attachment artifacts.
2. The automation layer creates normalized records, business context, and decision recommendations on top of those raw artifacts.

Do not let the LLM operate only on chat memory. Important mail knowledge and decision logic should be compiled into durable project artifacts.

## Recommended operational flow

```text
Outlook -> raw/mail -> normalization -> classification -> enrichment -> decision -> action -> wiki/Fabric/reporting
```

## Stage 1. Mail ingest

The system should check Outlook on a schedule and save all new mail into `raw/mail/`.

For each message, store:

- `message.json`
- `body.txt`
- `body.html` when available
- `source.md`
- `attachments/`

Important rule:

- `raw/mail/` is the archive of truth for incoming mail.
- Do not overwrite the original message content.
- Any later classification or correction should live outside the raw layer.

## Stage 2. Normalization

Each new email should be converted into a structured operational record.

Recommended extracted fields:

- message id
- internet message id
- conversation id
- sender
- recipients
- subject
- received datetime
- attachments present
- attachment count
- body preview
- detected language
- possible customer or counterparty
- process type candidate
- action required flag
- urgency candidate
- referenced system or service

At this stage, the system should answer:

- Is this a message that requires action?
- Is this a new case or part of an existing thread?
- Does it contain a document or evidence that should be tracked?

## Stage 3. Classification

Each message should be classified into one of a small set of process-oriented categories.

Recommended initial categories:

- `informational`
- `request`
- `incident`
- `approval`
- `document_review`
- `follow_up`
- `noise`

Recommended supporting labels:

- `needs_action`
- `customer_facing`
- `internal_only`
- `high_priority`
- `possible_duplicate`
- `contains_attachment`
- `requires_human_review`

Important rule:

- Classification should be conservative.
- If confidence is low, route the message to review instead of auto-deciding.

## Stage 4. Enrichment

After classification, the system should call business APIs, skill servers, or local rules to add business context.

Examples of enrichment:

- service or product mapping
- queue assignment
- owner lookup
- SLA policy lookup
- current case lookup
- duplicate detection
- customer tier
- process state
- escalation path

This stage should answer:

- Which business process owns this message?
- What SLA applies?
- Is this attached to an already open case?
- Who should be responsible next?

## Stage 5. Decision engine

The project should not jump directly from classification to autonomous action. It should use explicit decision modes.

### Decision modes

- `Auto`
  - safe actions with low business risk;
- `Suggest`
  - the system proposes the next action and waits for confirmation;
- `Manual`
  - the system only prepares context for human review.

### Recommended rule

Use:

- `Auto` for safe and reversible actions;
- `Suggest` for operationally important but still common decisions;
- `Manual` for high-risk, external, financial, legal, or ambiguous cases.

## Decision matrix

| Situation | Recommended mode | Action |
| --- | --- | --- |
| Newsletter, noise, duplicate alerts | Auto | archive, label, ignore |
| Follow-up to an existing known case | Auto or Suggest | attach to case, update timeline |
| Standard internal request with known routing | Suggest | propose queue, owner, due time |
| Incident with possible SLA impact | Suggest | create or update incident record, highlight risk |
| Customer complaint or escalation | Manual | human review with prepared context |
| Legal, financial, or contract-sensitive message | Manual | no automatic action |
| Ambiguous message with low model confidence | Manual | route to review queue |

## SLA handling

SLA should be attached after enrichment, not guessed only from message text.

For each mail-derived case, track:

- first response due
- resolution due
- current risk level
- queue owner
- current status
- last meaningful update

Recommended SLA states:

- `on_track`
- `at_risk`
- `breached`
- `waiting_external`
- `waiting_internal`
- `resolved`

## Actions the system may take

Recommended downstream actions:

- archive message
- tag and route message
- create task
- attach message to existing case
- create new process case
- create incident
- prepare reply draft
- prepare escalation summary
- request human review

Every action should produce an audit trail with:

- timestamp
- source message key
- chosen action
- reason
- actor type
- confidence

## Human-in-the-loop policy

The project should keep human approval where the cost of a wrong action is high.

Human review is strongly recommended when:

- the message affects an external customer;
- the decision may trigger escalation;
- the message is financially sensitive;
- the wording is ambiguous;
- the model confidence is low;
- a new exception pattern is detected.

## Recommended cadence

Suggested startup cadence:

- poll Outlook every 5 to 10 minutes;
- run classification immediately after ingest;
- run enrichment right after classification;
- refresh SLA risk view every few minutes;
- build a daily digest for the human operator.

Later optimization:

- move from full polling to Microsoft Graph delta query;
- move to push/webhook model when near-real-time behavior becomes necessary.

## Daily operator outputs

The project should produce useful working outputs, not just background processing.

Recommended outputs:

- inbox triage queue
- cases needing decision
- messages with low confidence
- SLA at-risk list
- overdue items
- daily digest
- weekly bottleneck report

## Recommended data model

Useful canonical entities:

- `MailMessage`
- `MailThread`
- `ProcessCase`
- `Task`
- `Incident`
- `Approval`
- `SlaPolicy`
- `SlaCheckpoint`
- `Decision`
- `OperatorAction`
- `AuditEvent`

## Recommended project rule

Do not let the LLM be the only source of decision logic.

Preferred architecture:

- explicit rules for deterministic cases;
- LLM classification and summarization for ambiguous text-heavy cases;
- human approval for high-risk actions.

This keeps the system explainable, auditable, and safer to automate over time.

## MVP for this repository

Recommended implementation order:

1. Stable Outlook ingest into `raw/mail/`
2. Normalize mail into a table or database record store
3. Classify mail into 5 to 7 business-process categories
4. Add the decision matrix with `Auto`, `Suggest`, and `Manual`
5. Build the `Needs decision` review screen in `Automation.Web`
6. Add SLA lookup and process enrichment
7. Sync facts and metrics into Microsoft Fabric

## Approved rollout order

The approved launch order for this project is:

1. Stable mail ingest into `raw/`
2. Normalize mail into a table or database
3. Build classification for 5 to 7 process types
4. Add the decision matrix
5. Build the `Needs decision` screen

This order is intentional.

Reasoning:

- first create a reliable archive of truth;
- then create structured operational records;
- then add process meaning through classification;
- then add explicit decision policy;
- only after that expose a human review surface.

This sequence should guide implementation planning unless the user explicitly reprioritizes.

## Immediate next steps

The next technical steps for this repository should be:

1. Add a normalized mail record store outside `raw/`
2. Implement a simple classifier for 5 to 7 process categories
3. Add the decision matrix and decision logging
4. Create the `Needs decision` review queue screen in the MVC portal
5. Add SLA policy mapping and process enrichment
6. Add a daily digest job
7. Add Fabric export for metrics and case facts

## Summary

The recommended design is:

- Outlook is the source of incoming operational events;
- `raw/` is the archive of truth;
- structured records carry business meaning;
- enrichment adds service and process context;
- decisions happen through explicit modes;
- high-risk decisions stay human-reviewed;
- durable knowledge should be stored in project docs, logs, and analytics outputs.
