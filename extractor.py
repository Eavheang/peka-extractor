"""Extract source files into neutral text for a later LLM/OKF stage.

This module deliberately does not classify, summarize, categorize, or write OKF.
It only extracts content and records source identity.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import zipfile
from datetime import UTC, datetime
from pathlib import Path
from xml.etree import ElementTree

import fitz


SUPPORTED = {".pdf", ".docx", ".xlsx", ".pptx", ".txt", ".md"}


def clean_text(text: str) -> str:
    text = text.replace("\u00a0", " ")
    text = re.sub(r"[ \t]+", " ", text)
    return re.sub(r"\n{3,}", "\n\n", text).strip()


def xml_text(data: bytes) -> str:
    root = ElementTree.fromstring(data)
    return clean_text("\n".join(value for value in root.itertext() if value.strip()))


def extract_pdf(path: Path) -> list[dict[str, object]]:
    with fitz.open(path) as document:
        return [
            {"page": number, "text": clean_text(page.get_text())}
            for number, page in enumerate(document, start=1)
            if clean_text(page.get_text())
        ]


def extract_office(path: Path) -> list[dict[str, object]]:
    """Extract readable XML text from Office Open XML files."""
    with zipfile.ZipFile(path) as archive:
        names = sorted(name for name in archive.namelist() if name.endswith(".xml"))
        pages = []
        for name in names:
            text = xml_text(archive.read(name))
            if text:
                pages.append({"source_part": name, "text": text})
        return pages


def extract_file(path: Path) -> dict[str, object]:
    suffix = path.suffix.lower()
    if suffix not in SUPPORTED:
        raise ValueError(f"Unsupported file type: {path.suffix or '(none)'}")
    content_hash = hashlib.sha256(path.read_bytes()).hexdigest()
    if suffix == ".pdf":
        sections = extract_pdf(path)
    elif suffix in {".docx", ".xlsx", ".pptx"}:
        sections = extract_office(path)
    else:
        text = clean_text(path.read_text(encoding="utf-8", errors="replace"))
        sections = [{"text": text}] if text else []
    return {
        "source_id": content_hash[:16],
        "source_sha256": content_hash,
        "source_filename": path.name,
        "source_path": str(path.resolve()),
        "source_extension": suffix,
        "extracted_at": datetime.now(UTC).isoformat(),
        "sections": sections,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Extract files for LLM processing")
    parser.add_argument("files", type=Path, nargs="+", help="Source files")
    parser.add_argument("--output", type=Path, default=Path("extracted"))
    args = parser.parse_args()
    args.output.mkdir(parents=True, exist_ok=True)
    registry_path = args.output / "registry.json"
    registry = json.loads(registry_path.read_text()) if registry_path.exists() else {}

    for path in args.files:
        if not path.is_file():
            raise SystemExit(f"File not found: {path}")
        record = extract_file(path)
        source_key = record["source_path"]
        previous = registry.get(source_key)
        if previous and previous["source_sha256"] == record["source_sha256"]:
            print(f"Skipped unchanged: {path}")
            continue
        output_file = args.output / f"{record['source_id']}.json"
        output_file.write_text(json.dumps(record, indent=2) + "\n", encoding="utf-8")
        registry[source_key] = {
            "source_id": record["source_id"],
            "source_sha256": record["source_sha256"],
            "output": str(output_file),
        }
        print(f"Extracted: {path} -> {output_file}")

    registry_path.write_text(json.dumps(registry, indent=2) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
