"""JSON file discovery and loading for export run inputs."""

from __future__ import annotations

import json
from pathlib import Path


def find_ae_pages(input_dir: Path) -> list[Path]:
    """Find Activity Explorer page files (JSON or JSONL)."""
    ae_dir = input_dir / "Data" / "ActivityExplorer"
    if not ae_dir.exists():
        return []
    pages = list(ae_dir.rglob("Page-*.json")) + list(ae_dir.rglob("Page-*.jsonl"))
    return sorted(pages)


def find_ce_pages(input_dir: Path) -> list[Path]:
    """Find Content Explorer page files (JSON or JSONL, new and old formats)."""
    pages = []

    # New format: Data/ContentExplorer/TagType/TagName/{Workload}-NNN.{json,jsonl}
    ce_dir = input_dir / "Data" / "ContentExplorer"
    if ce_dir.exists():
        for ext in ("*.json", "*.jsonl"):
            for f in ce_dir.rglob(ext):
                if f.name.startswith("_") or f.name.startswith("agg-"):
                    continue
                pages.append(f)

    # Old format: Worker-PID/detail-*.json
    for worker_dir in input_dir.glob("Worker-*"):
        if worker_dir.is_dir():
            for ext in ("detail-*.json", "detail-*.jsonl"):
                for f in worker_dir.glob(ext):
                    pages.append(f)

    return sorted(pages)


def _load_jsonl_records(path: Path) -> list[dict]:
    """Load JSONL: one JSON object per line; tolerate blank lines."""
    records = []
    try:
        with open(path, "r", encoding="utf-8-sig") as f:
            for line_num, raw in enumerate(f, start=1):
                line = raw.strip()
                if not line:
                    continue
                try:
                    rec = json.loads(line)
                except json.JSONDecodeError as exc:
                    print(f"  WARNING: Skipping malformed JSONL line {line_num} in {path.name}: {exc}")
                    continue
                if isinstance(rec, dict):
                    records.append(rec)
    except UnicodeDecodeError as exc:
        print(f"  WARNING: Skipping unreadable JSONL file: {path.name} ({exc})")
    return records


def load_page_records(path: Path) -> list[dict]:
    """Load records from a page file (.json wrapper, .json flat list, or .jsonl)."""
    if path.suffix.lower() == ".jsonl":
        return _load_jsonl_records(path)

    try:
        with open(path, "r", encoding="utf-8-sig") as f:
            data = json.load(f)
    except (json.JSONDecodeError, UnicodeDecodeError) as exc:
        print(f"  WARNING: Skipping malformed JSON file: {path.name} ({exc})")
        return []

    if isinstance(data, dict) and "Records" in data:
        records = data["Records"]
        if isinstance(records, list):
            return [rec for rec in records if isinstance(rec, dict)]
        if isinstance(records, dict):
            return [records]
        return []
    if isinstance(data, list):
        return [rec for rec in data if isinstance(rec, dict)]
    return []


def load_json_config(path: Path) -> dict | None:
    """Load a JSON config export file."""
    if not path.exists():
        return None
    with open(path, "r", encoding="utf-8-sig") as f:
        return json.load(f)
