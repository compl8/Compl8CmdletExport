"""Content Explorer page processing."""

from __future__ import annotations

import json
from pathlib import Path

from .constants import CE_METADATA_FIELDS, CONTENT_RENAMES
from .helpers import (
    _first_present,
    _now_iso,
    _parse_nested_json,
    _rename_record,
    _safe_int,
    _safe_str,
    _sha1_text,
    _split_sit_ids,
)
from .loaders import find_ce_pages, load_page_records


def process_content(input_dir: Path, drift_tracker=None) -> tuple[list[dict], list[dict]]:
    """Process CE pages -> content_files and sit_detections lists."""
    pages = find_ce_pages(input_dir)
    if not pages:
        return [], []

    ingested_at = _now_iso()
    content_files = []
    sit_detections_by_key: dict[tuple[str, str], dict] = {}

    for page_path in pages:
        records = load_page_records(page_path)

        # Try to get tag info from page wrapper (.json only; JSONL has no wrapper
        # but records carry _ExportTagType / _ExportTagName)
        page_tag_type = None
        page_tag_name = None
        if page_path.suffix.lower() == ".json":
            try:
                with open(page_path, "r", encoding="utf-8-sig") as f:
                    wrapper = json.load(f)
                if isinstance(wrapper, dict):
                    page_tag_type = wrapper.get("TagType")
                    page_tag_name = wrapper.get("TagName")
            except Exception:
                pass

        for raw in records:
            renamed, extra = _rename_record(raw, CONTENT_RENAMES, excluded_keys=CE_METADATA_FIELDS)
            if drift_tracker is not None:
                drift_tracker.record("content_files", extra)

            # Add tag metadata from record or page wrapper
            renamed["tag_type"] = raw.get("_ExportTagType") or page_tag_type
            renamed["tag_name"] = raw.get("_ExportTagName") or page_tag_name
            renamed["_source_tool"] = "cmdletexport"
            renamed["_ingested_at"] = ingested_at

            if not renamed.get("doc_id"):
                renamed["doc_id"] = _sha1_text(
                    renamed.get("file_url")
                    or "|".join([
                        renamed.get("source_url") or "",
                        renamed.get("file_name") or "",
                    ])
                )

            # Serialize matches_json if it's still a complex type
            if "matches_json" in renamed and isinstance(renamed["matches_json"], (list, dict)):
                renamed["matches_json"] = json.dumps(renamed["matches_json"], default=str)

            renamed["extra_fields"] = json.dumps(extra, default=str) if extra else None

            content_files.append(renamed)

            doc_id = renamed.get("doc_id")
            if not doc_id:
                continue

            # Each CE record carries a per-document SensitiveInfoTypesData payload listing
            # every SIT detected on that doc (not just the export's tag). When the same doc
            # appears in multiple SIT tag buckets the payload is byte-identical, so max()
            # dedupes correctly; sum() would double-count. Verified against real exports.
            parsed_matches = _parse_nested_json(renamed.get("matches_json")) or []
            for sit in parsed_matches:
                sit_id = _safe_str(
                    sit.get("Id")
                    or sit.get("SensitiveInfoTypeId")
                    or sit.get("SensitiveType")
                    or sit.get("sit_id")
                )
                if not sit_id:
                    continue
                sit_id = sit_id.strip().lower()
                low = _safe_int(_first_present(
                    sit.get("LowConfidenceMatch"),
                    sit.get("LowCount"),
                    sit.get("Low"),
                    sit.get("low_count"),
                ))
                medium = _safe_int(_first_present(
                    sit.get("MediumConfidenceMatch"),
                    sit.get("MediumCount"),
                    sit.get("Medium"),
                    sit.get("medium_count"),
                ))
                high = _safe_int(_first_present(
                    sit.get("HighConfidenceMatch"),
                    sit.get("HighCount"),
                    sit.get("High"),
                    sit.get("high_count"),
                ))
                key = (doc_id, sit_id)
                existing = sit_detections_by_key.get(key)
                if existing is None:
                    sit_detections_by_key[key] = {
                        "doc_id": doc_id,
                        "sit_id": sit_id,
                        "low_count": low,
                        "medium_count": medium,
                        "high_count": high,
                        "total_count": low + medium + high,
                        "_source_tool": "cmdletexport",
                        "_ingested_at": ingested_at,
                    }
                else:
                    existing["low_count"] = max(existing["low_count"], low)
                    existing["medium_count"] = max(existing["medium_count"], medium)
                    existing["high_count"] = max(existing["high_count"], high)
                    existing["total_count"] = (
                        existing["low_count"]
                        + existing["medium_count"]
                        + existing["high_count"]
                    )

            if not parsed_matches:
                for sit_id in _split_sit_ids(renamed.get("sensitive_info_type_ids")):
                    key = (doc_id, sit_id)
                    if key not in sit_detections_by_key:
                        sit_detections_by_key[key] = {
                            "doc_id": doc_id,
                            "sit_id": sit_id,
                            "low_count": 0,
                            "medium_count": 0,
                            "high_count": 0,
                            "total_count": 0,
                            "_source_tool": "cmdletexport",
                            "_ingested_at": ingested_at,
                        }

    sit_detections = list(sit_detections_by_key.values())
    print(f"  Content files: {len(content_files)} records")
    print(f"  Content SIT detections: {len(sit_detections)} records")
    return content_files, sit_detections
