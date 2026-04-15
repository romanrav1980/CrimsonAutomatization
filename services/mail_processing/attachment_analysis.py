from __future__ import annotations

import csv
import hashlib
import io
import json
import re
import struct
import zipfile
from dataclasses import asdict, dataclass, field
from datetime import UTC, datetime
from pathlib import Path
from typing import Any
from xml.etree import ElementTree

from services.mail_ingest.storage import classify_attachment_kind
from services.mail_processing.models import RawMailArtifact


WORKBOOK_NS = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
PDF_PAGE_PATTERN = re.compile(rb"/Type\s*/Page\b")


@dataclass(slots=True)
class AttachmentInsight:
    name: str
    stored_path: str
    kind: str
    content_type: str
    extension: str
    size: int
    sha256: str
    saved: bool
    analysis_status: str
    summary: str
    signals: list[str] = field(default_factory=list)
    worksheet_count: int = 0
    sheet_names: list[str] = field(default_factory=list)
    page_count: int = 0
    text_excerpt: str = ""
    image_width: int = 0
    image_height: int = 0


@dataclass(slots=True)
class MessageAttachmentAnalysis:
    generated_utc: str
    message_key: str
    message_subject: str
    attachment_count: int
    analyzed_attachment_count: int
    attachment_kinds: list[str]
    summary: str
    analysis_path: str
    attachments: list[dict[str, Any]] = field(default_factory=list)


class AttachmentAnalysisService:
    def __init__(self, *, raw_root: Path, derived_root: Path) -> None:
        self.raw_root = raw_root.resolve()
        self.derived_root = derived_root.resolve()
        self.workspace_root = self.raw_root.parents[1] if len(self.raw_root.parents) >= 2 else self.raw_root.parent

    def analyze_message(self, artifact: RawMailArtifact) -> MessageAttachmentAnalysis:
        message_dir = Path(artifact.message_dir_path)
        manifest_path = Path(artifact.attachments_manifest_path) if artifact.attachments_manifest_path else (message_dir / "attachments.json")
        payload = self._load_manifest(manifest_path)
        insights: list[AttachmentInsight] = []

        for item in payload:
            try:
                insight = self._analyze_attachment(item, message_dir=message_dir)
                insights.append(insight)
            except Exception as exc:
                name = str(item.get("originalName") or item.get("name") or "attachment")
                kind = classify_attachment_kind(name, item.get("contentType"))
                insights.append(
                    AttachmentInsight(
                        name=name,
                        stored_path=str(item.get("path") or ""),
                        kind=kind,
                        content_type=str(item.get("contentType") or ""),
                        extension=str(item.get("extension") or Path(name).suffix.lower()),
                        size=int(item.get("size") or 0),
                        sha256=str(item.get("sha256") or ""),
                        saved=bool(item.get("saved")),
                        analysis_status="error",
                        summary=f"{kind.capitalize()} attachment analysis failed: {exc}",
                    )
                )

        attachment_kinds = sorted({item.kind for item in insights if item.kind})
        summary = self._build_summary(insights)
        generated_utc = self._utc_now()
        target_dir = self.derived_root / artifact.message_key
        target_dir.mkdir(parents=True, exist_ok=True)
        analysis_path = target_dir / "attachment_analysis.json"
        analysis_payload = {
            "generatedUtc": generated_utc,
            "messageKey": artifact.message_key,
            "subject": artifact.subject,
            "attachmentCount": len(payload),
            "analyzedAttachmentCount": len(insights),
            "attachmentKinds": attachment_kinds,
            "summary": summary,
            "attachments": [asdict(item) for item in insights],
        }
        analysis_path.write_text(json.dumps(analysis_payload, ensure_ascii=False, indent=2), encoding="utf-8")

        return MessageAttachmentAnalysis(
            generated_utc=generated_utc,
            message_key=artifact.message_key,
            message_subject=artifact.subject,
            attachment_count=len(payload),
            analyzed_attachment_count=len(insights),
            attachment_kinds=attachment_kinds,
            summary=summary,
            analysis_path=str(analysis_path).replace("\\", "/"),
            attachments=analysis_payload["attachments"],
        )

    @staticmethod
    def _load_manifest(manifest_path: Path) -> list[dict[str, Any]]:
        if not manifest_path.exists():
            return []
        return json.loads(manifest_path.read_text(encoding="utf-8"))

    def _analyze_attachment(self, entry: dict[str, Any], *, message_dir: Path) -> AttachmentInsight:
        name = str(entry.get("originalName") or entry.get("name") or "attachment")
        stored_path = self._resolve_attachment_path(entry, message_dir=message_dir)
        content_type = str(entry.get("contentType") or "")
        kind = str(entry.get("kind") or classify_attachment_kind(name, content_type))
        extension = str(entry.get("extension") or stored_path.suffix.lower() or Path(name).suffix.lower())
        size = int(entry.get("size") or (stored_path.stat().st_size if stored_path.exists() else 0))
        sha256 = str(entry.get("sha256") or "")
        saved = bool(entry.get("saved")) and stored_path.exists()

        if saved and not sha256:
            sha256 = hashlib.sha256(stored_path.read_bytes()).hexdigest()

        if not stored_path.exists():
            return AttachmentInsight(
                name=name,
                stored_path=str(stored_path).replace("\\", "/"),
                kind=kind,
                content_type=content_type,
                extension=extension,
                size=size,
                sha256=sha256,
                saved=False,
                analysis_status="missing_file",
                summary=f"{kind.capitalize()} attachment metadata exists, but the saved file is missing.",
            )

        if kind == "excel":
            return self._analyze_excel(
                name=name,
                stored_path=stored_path,
                content_type=content_type,
                extension=extension,
                size=size,
                sha256=sha256,
            )

        if kind == "pdf":
            return self._analyze_pdf(
                name=name,
                stored_path=stored_path,
                content_type=content_type,
                extension=extension,
                size=size,
                sha256=sha256,
            )

        if kind == "image":
            return self._analyze_image(
                name=name,
                stored_path=stored_path,
                content_type=content_type,
                extension=extension,
                size=size,
                sha256=sha256,
            )

        return AttachmentInsight(
            name=name,
            stored_path=str(stored_path).replace("\\", "/"),
            kind=kind,
            content_type=content_type,
            extension=extension,
            size=size,
            sha256=sha256,
            saved=True,
            analysis_status="metadata_only",
            summary=f"Stored {kind} attachment with no specialized analyzer yet.",
            signals=[f"ext {extension or 'n/a'}", self._format_size(size)],
        )

    def _analyze_excel(
        self,
        *,
        name: str,
        stored_path: Path,
        content_type: str,
        extension: str,
        size: int,
        sha256: str,
    ) -> AttachmentInsight:
        extension = extension.lower()
        signals = [self._format_size(size)]
        worksheet_count = 0
        sheet_names: list[str] = []
        text_excerpt = ""
        analysis_status = "metadata_only"

        if extension == ".csv":
            rows = self._load_csv_preview(stored_path)
            worksheet_count = 1
            sheet_names = ["csv"]
            analysis_status = "preview_ready"
            if rows:
                text_excerpt = "\n".join(" | ".join(cell for cell in row if cell) for row in rows[:3])
                signals.append(f"preview rows {len(rows)}")
        elif extension in {".xlsx", ".xlsm", ".xltx", ".xltm"}:
            worksheet_count, sheet_names = self._read_xlsx_sheet_names(stored_path)
            analysis_status = "metadata_ready"
            if worksheet_count > 0:
                signals.append(f"sheets {worksheet_count}")
        elif extension in {".xls", ".xlsb"}:
            signals.append("legacy workbook format")
        else:
            signals.append(f"ext {extension or 'n/a'}")

        if sheet_names:
            signals.append(", ".join(sheet_names[:3]))

        summary = self._summarize_excel(worksheet_count, sheet_names, extension)
        return AttachmentInsight(
            name=name,
            stored_path=str(stored_path).replace("\\", "/"),
            kind="excel",
            content_type=content_type,
            extension=extension,
            size=size,
            sha256=sha256,
            saved=True,
            analysis_status=analysis_status,
            summary=summary,
            signals=signals,
            worksheet_count=worksheet_count,
            sheet_names=sheet_names,
            text_excerpt=text_excerpt,
        )

    def _analyze_pdf(
        self,
        *,
        name: str,
        stored_path: Path,
        content_type: str,
        extension: str,
        size: int,
        sha256: str,
    ) -> AttachmentInsight:
        page_count = 0
        text_excerpt = ""
        analysis_status = "metadata_only"
        signals = [self._format_size(size)]

        try:
            from pypdf import PdfReader  # type: ignore

            reader = PdfReader(str(stored_path))
            page_count = len(reader.pages)
            excerpt_parts: list[str] = []
            for page in reader.pages[:2]:
                try:
                    excerpt = (page.extract_text() or "").strip()
                except Exception:
                    excerpt = ""
                if excerpt:
                    excerpt_parts.append(excerpt)
            text_excerpt = "\n".join(excerpt_parts)[:800]
            analysis_status = "text_ready" if text_excerpt else "metadata_ready"
        except Exception:
            data = stored_path.read_bytes()
            page_count = len(PDF_PAGE_PATTERN.findall(data))
            analysis_status = "metadata_ready"

        if page_count > 0:
            signals.append(f"pages {page_count}")
        if not text_excerpt:
            signals.append("text extraction optional")

        summary = "PDF attachment"
        if page_count > 0:
            summary += f", {page_count} pages"
        if text_excerpt:
            summary += ", text preview ready"

        return AttachmentInsight(
            name=name,
            stored_path=str(stored_path).replace("\\", "/"),
            kind="pdf",
            content_type=content_type,
            extension=extension,
            size=size,
            sha256=sha256,
            saved=True,
            analysis_status=analysis_status,
            summary=summary,
            signals=signals,
            page_count=page_count,
            text_excerpt=text_excerpt,
        )

    def _analyze_image(
        self,
        *,
        name: str,
        stored_path: Path,
        content_type: str,
        extension: str,
        size: int,
        sha256: str,
    ) -> AttachmentInsight:
        width = 0
        height = 0
        try:
            width, height = self._read_image_dimensions(stored_path)
        except Exception:
            width, height = 0, 0

        signals = [self._format_size(size)]
        if width > 0 and height > 0:
            signals.append(f"{width}x{height}")
        signals.append("ocr not configured")

        summary = "Image attachment"
        if width > 0 and height > 0:
            summary += f", {width}x{height}"

        return AttachmentInsight(
            name=name,
            stored_path=str(stored_path).replace("\\", "/"),
            kind="image",
            content_type=content_type,
            extension=extension,
            size=size,
            sha256=sha256,
            saved=True,
            analysis_status="metadata_ready",
            summary=summary,
            signals=signals,
            image_width=width,
            image_height=height,
        )

    def _resolve_attachment_path(self, entry: dict[str, Any], *, message_dir: Path) -> Path:
        raw_value = str(entry.get("path") or "").strip()
        if raw_value:
            candidate = Path(raw_value)
            if candidate.is_absolute() and candidate.exists():
                return candidate
            repo_relative = (self.workspace_root / candidate).resolve()
            if repo_relative.exists():
                return repo_relative

        stored_name = str(entry.get("storedName") or "")
        if stored_name:
            nested = message_dir / "attachments" / stored_name
            if nested.exists():
                return nested.resolve()
            for kind_dir in ("excel", "pdf", "image", "other"):
                candidate = message_dir / "attachments" / kind_dir / stored_name
                if candidate.exists():
                    return candidate.resolve()

        original_name = str(entry.get("originalName") or entry.get("name") or "")
        if original_name:
            candidate = message_dir / "attachments" / original_name
            if candidate.exists():
                return candidate.resolve()

        return (message_dir / "attachments" / "missing").resolve()

    @staticmethod
    def _load_csv_preview(path: Path) -> list[list[str]]:
        data = path.read_bytes()
        text = AttachmentAnalysisService._decode_text(data)
        reader = csv.reader(io.StringIO(text))
        rows: list[list[str]] = []
        for row in reader:
            rows.append([cell.strip()[:80] for cell in row[:8]])
            if len(rows) >= 5:
                break
        return rows

    @staticmethod
    def _read_xlsx_sheet_names(path: Path) -> tuple[int, list[str]]:
        with zipfile.ZipFile(path) as archive:
            workbook_xml = archive.read("xl/workbook.xml")
        root = ElementTree.fromstring(workbook_xml)
        sheet_names = [sheet.attrib.get("name", "").strip() for sheet in root.findall(".//main:sheets/main:sheet", WORKBOOK_NS)]
        sheet_names = [name for name in sheet_names if name]
        return len(sheet_names), sheet_names

    @staticmethod
    def _read_image_dimensions(path: Path) -> tuple[int, int]:
        data = path.read_bytes()
        if data.startswith(b"\x89PNG\r\n\x1a\n") and len(data) >= 24:
            width = struct.unpack(">I", data[16:20])[0]
            height = struct.unpack(">I", data[20:24])[0]
            return width, height

        if data.startswith(b"GIF87a") or data.startswith(b"GIF89a"):
            width, height = struct.unpack("<HH", data[6:10])
            return width, height

        if data.startswith(b"BM") and len(data) >= 26:
            width = struct.unpack("<I", data[18:22])[0]
            height = struct.unpack("<I", data[22:26])[0]
            return width, abs(height)

        if data.startswith(b"\xff\xd8"):
            return AttachmentAnalysisService._read_jpeg_dimensions(data)

        return 0, 0

    @staticmethod
    def _read_jpeg_dimensions(data: bytes) -> tuple[int, int]:
        offset = 2
        while offset + 9 < len(data):
            if data[offset] != 0xFF:
                offset += 1
                continue
            marker = data[offset + 1]
            offset += 2
            if marker in {0xD8, 0xD9}:
                continue
            if offset + 2 > len(data):
                break
            length = struct.unpack(">H", data[offset:offset + 2])[0]
            if length < 2 or offset + length > len(data):
                break
            if marker in {
                0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7,
                0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF,
            }:
                if offset + 7 <= len(data):
                    height = struct.unpack(">H", data[offset + 3:offset + 5])[0]
                    width = struct.unpack(">H", data[offset + 5:offset + 7])[0]
                    return width, height
                break
            offset += length
        return 0, 0

    @staticmethod
    def _decode_text(data: bytes) -> str:
        for encoding in ("utf-8-sig", "utf-8", "cp1251", "cp866", "latin-1"):
            try:
                return data.decode(encoding)
            except UnicodeDecodeError:
                continue
        return data.decode("utf-8", errors="replace")

    @staticmethod
    def _summarize_excel(worksheet_count: int, sheet_names: list[str], extension: str) -> str:
        if extension == ".csv":
            return "CSV attachment with preview rows ready"
        if worksheet_count > 0:
            shown = ", ".join(sheet_names[:3])
            if shown:
                return f"Excel workbook, {worksheet_count} sheet(s): {shown}"
            return f"Excel workbook, {worksheet_count} sheet(s)"
        return "Excel attachment stored for later parsing"

    @staticmethod
    def _build_summary(insights: list[AttachmentInsight]) -> str:
        if not insights:
            return "No attachments to analyze."

        counts: dict[str, int] = {}
        for insight in insights:
            counts[insight.kind] = counts.get(insight.kind, 0) + 1

        parts = [f"{counts[kind]} {kind}" for kind in sorted(counts.keys())]
        return "Analyzed {0} attachment(s): {1}".format(len(insights), ", ".join(parts))

    @staticmethod
    def _format_size(size: int) -> str:
        if size <= 0:
            return "size n/a"

        units = ["B", "KB", "MB", "GB"]
        value = float(size)
        unit_index = 0
        while value >= 1024 and unit_index < len(units) - 1:
            value /= 1024
            unit_index += 1
        if unit_index == 0:
            return f"{int(value)} {units[unit_index]}"
        return f"{value:.1f} {units[unit_index]}"

    @staticmethod
    def _utc_now() -> str:
        return datetime.now(UTC).replace(microsecond=0).isoformat().replace("+00:00", "Z")
