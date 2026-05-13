"""JSON file discovery and loading for export run inputs."""

from __future__ import annotations

import json
from pathlib import Path


def find_ae_pages(input_dir: Path) -> list[Path]:
    """Find Activity Explorer page JSON files (new format)."""
    ae_dir = input_dir / "Data" / "ActivityExplorer"
    if not ae_dir.exists():
        return []
    return sorted(ae_dir.rglob("Page-*.json"))


def find_ce_pages(input_dir: Path) -> list[Path]:
    """Find Content Explorer page JSON files (new and old formats)."""
    pages = []

    # New format: Data/ContentExplorer/TagType/TagName/{Workload}-NNN.json
    ce_dir = input_dir / "Data" / "ContentExplorer"
    if ce_dir.exists():
        for f in ce_dir.rglob("*.json"):
            if f.name.startswith("_") or f.name.startswith("agg-"):
                continue
            pages.append(f)

    # Old format: Worker-PID/detail-*.json
    for worker_dir in input_dir.glob("Worker-*"):
        if worker_dir.is_dir():
            for f in worker_dir.glob("detail-*.json"):
                pages.append(f)

    return sorted(pages)


def load_page_records(path: Path) -> list[dict]:
    """Load records from a page JSON file (handles both wrapped and flat formats)."""
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
