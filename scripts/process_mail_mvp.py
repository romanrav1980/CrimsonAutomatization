from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from services.mail_ingest.config import load_env_file
from services.mail_processing.pipeline import MailProcessingPipeline


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Process raw mail into the MVP database and dashboard read model.")
    parser.add_argument("--raw-root", default=os.getenv("OUTLOOK_OUTPUT_DIR", "raw/mail"), help="Path to the raw mail root.")
    parser.add_argument("--db-path", default=os.getenv("AUTOMATION_DB_PATH", "data/automation.db"), help="Path to the SQLite database.")
    parser.add_argument(
        "--derived-root",
        default=os.getenv("AUTOMATION_DERIVED_ROOT", "derived/mail"),
        help="Path to the derived mail analysis root.",
    )
    parser.add_argument(
        "--read-model-path",
        default=os.getenv("AUTOMATION_READ_MODEL_PATH", "data/mail_triage_read_model.json"),
        help="Path to the exported dashboard read model JSON.",
    )
    return parser


def parse_internal_domains() -> set[str]:
    raw_value = os.getenv("MAIL_INTERNAL_DOMAINS", "")
    return {item.strip().lower() for item in raw_value.split(",") if item.strip()}


def emit_stage(stage: str, state: str, detail: str) -> None:
    print(f"[[STAGE|{stage}|{state}|{detail}]]")


def emit_metric(name: str, value: int, label: str) -> None:
    print(f"[[METRIC|{name}|{value}|{label}]]")


def main() -> int:
    load_env_file()
    parser = build_parser()
    args = parser.parse_args()

    pipeline = MailProcessingPipeline(
        raw_root=Path(args.raw_root),
        derived_root=Path(args.derived_root),
        db_path=Path(args.db_path),
        read_model_path=Path(args.read_model_path),
        internal_domains=parse_internal_domains(),
    )
    emit_stage("process", "running", "Starting raw mail processing pipeline")
    result = pipeline.process()
    emit_stage("process", "done", f"Updated database and read model with {result.processed} items")
    emit_metric("normalized_loaded", result.normalized_loaded, "Raw artifacts loaded")
    emit_metric("attachments_analyzed", result.attachments_analyzed, "Attachments analyzed")
    emit_metric("classified", result.classified, "Mail artifacts classified")
    emit_metric("processed", result.processed, "Processed items")
    emit_metric("failed", result.failed, "Failed items")
    emit_metric("needs_decision", result.needs_decision, "Items in Needs Decision queue")
    print("Mail MVP processing complete")
    print(f"Processed raw messages: {result.processed}")
    print(f"Failed raw messages: {result.failed}")
    print(f"Database: {result.db_path}")
    print(f"Read model: {result.exported_read_model_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
