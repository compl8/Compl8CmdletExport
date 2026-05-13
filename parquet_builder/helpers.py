"""Small pure utilities shared by every stage of the pipeline."""

from __future__ import annotations

import hashlib
import json
import re
from datetime import datetime, timezone
from pathlib import Path


def _now_iso() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _safe_str(val) -> str | None:
    if val is None:
        return None
    if isinstance(val, (list, dict)):
        return json.dumps(val, default=str)
    return str(val)


def _safe_int(val) -> int:
    if val is None or val == "":
        return 0
    try:
        return int(float(str(val).strip()))
    except (TypeError, ValueError):
        return 0


def _first_present(*values):
    for val in values:
        if val is not None and val != "":
            return val
    return None


def _sha1_text(val: str | None) -> str | None:
    if not val:
        return None
    # SHA-1 is used here as a content-addressing hash (doc_id derivation),
    # not for any security or integrity purpose.
    return hashlib.sha1(
        val.encode("utf-8", errors="ignore"), usedforsecurity=False
    ).hexdigest()


def _parse_nested_json(val) -> list[dict] | None:
    """Parse a nested JSON blob that may be a string, list, or None."""
    if val is None:
        return None
    if isinstance(val, str):
        if not val.strip():
            return None
        try:
            parsed = json.loads(val)
            if isinstance(parsed, list):
                return parsed
            if isinstance(parsed, dict):
                return [parsed]
            return None
        except (json.JSONDecodeError, ValueError):
            return None
    if isinstance(val, list):
        return val
    if isinstance(val, dict):
        return [val]
    return None


def _split_sit_ids(val) -> list[str]:
    if val is None:
        return []
    if isinstance(val, list):
        raw_values = val
    else:
        raw_values = re.split(r"[,;]", str(val))
    sit_ids = []
    seen = set()
    for raw in raw_values:
        sit_id = str(raw).strip().lower()
        if not sit_id or sit_id in seen:
            continue
        sit_ids.append(sit_id)
        seen.add(sit_id)
    return sit_ids


def _extract_file_name(file_path: str | None) -> str | None:
    if not file_path:
        return None
    # Handle both URL paths and file system paths
    name = file_path.rsplit("/", 1)[-1].rsplit("\\", 1)[-1]
    return name if name else None


def _rename_record(record: dict, rename_map: dict, excluded_keys: set | None = None) -> tuple[dict, dict]:
    """Rename keys per map, returning (renamed_dict, extra_fields_dict)."""
    renamed = {}
    extra = {}
    excluded = excluded_keys or set()
    mapped_sources = set(rename_map.keys())

    for key, val in record.items():
        if key in excluded:
            continue
        if key in rename_map:
            target = rename_map[key]
            if target not in renamed:  # first match wins
                renamed[target] = val
        elif key not in mapped_sources:
            extra[key] = val

    return renamed, extra


def _run_stamp(input_dir: Path) -> str:
    """Extract run_YYYYMMDD_HHMMSS from directory name like Export-20260307-123456."""
    name = input_dir.name
    m = re.search(r"(\d{8})[_-](\d{6})", name)
    if m:
        return f"run_{m.group(1)}_{m.group(2)}"
    return f"run_{datetime.now(timezone.utc).strftime('%Y%m%d_%H%M%S')}"
