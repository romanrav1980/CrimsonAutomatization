from __future__ import annotations

import argparse
import json
import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from services.mail_ingest.config import load_env_file
from services.mail_processing.database import MailMvpRepository


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Apply an operator decision to a mail item and refresh the dashboard read model.")
    parser.add_argument("--message-key", required=True, help="Target mail item key.")
    parser.add_argument("--action", required=True, choices=["approve", "archive", "manual", "assign_owner"], help="Operator action to apply.")
    parser.add_argument("--owner", default="", help="Owner value for assign_owner or ownership updates.")
    parser.add_argument("--notes", default="", help="Optional operator notes.")
    parser.add_argument("--actor", default="", help="Operator identity.")
    parser.add_argument("--db-path", default=os.getenv("AUTOMATION_DB_PATH", "data/automation.db"), help="Path to the SQLite database.")
    parser.add_argument(
        "--read-model-path",
        default=os.getenv("AUTOMATION_READ_MODEL_PATH", "data/mail_triage_read_model.json"),
        help="Path to the dashboard read model JSON.",
    )
    return parser


def main() -> int:
    load_env_file()
    parser = build_parser()
    args = parser.parse_args()

    actor = args.actor.strip() or os.getenv("USERNAME", "") or "local-operator"
    repository = MailMvpRepository(Path(args.db_path))
    result = repository.apply_operator_action(
        message_key=args.message_key,
        action=args.action,
        actor=actor,
        owner=args.owner,
        notes=args.notes,
    )
    repository.export_dashboard_snapshot(Path(args.read_model_path))

    print(json.dumps(result, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
