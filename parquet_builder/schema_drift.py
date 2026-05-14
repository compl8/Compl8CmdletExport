"""Schema drift detection.

When Microsoft adds fields to a Purview cmdlet's output, they fall through
the rename maps into the per-record extra_fields blob. This module aggregates
those unknown keys across all records processed during a parquet build and
writes a SchemaDrift.json report at the end so operators can update the
rename maps before the unknown fields silently accumulate.
"""

from __future__ import annotations

import json
import re
from pathlib import Path

from .helpers import _now_iso, _safe_str


def _to_snake_case(name: str) -> str:
    """Convert PascalCase/camelCase to snake_case for rename-map suggestions."""
    # Handle acronyms like 'MDATPDeviceId' -> 'mdatp_device_id' and
    # camelCase like 'createdDateTime' -> 'created_date_time'
    s1 = re.sub(r"(.)([A-Z][a-z]+)", r"\1_\2", name)
    s2 = re.sub(r"([a-z0-9])([A-Z])", r"\1_\2", s1)
    return s2.replace("__", "_").lower().strip("_")


class SchemaDriftTracker:
    """Accumulates unknown field counts and samples per table."""

    def __init__(self) -> None:
        # table_name -> field_name -> {count, sample_value}
        self._by_table: dict[str, dict[str, dict]] = {}

    def record(self, table: str, extras: dict) -> None:
        """Note any unknown keys from one record's extra_fields dict."""
        if not extras:
            return
        bucket = self._by_table.setdefault(table, {})
        for key, value in extras.items():
            if not key:
                continue
            entry = bucket.setdefault(key, {"count": 0, "sample_value": None})
            entry["count"] += 1
            if entry["sample_value"] is None and value not in (None, "", [], {}):
                sample = _safe_str(value)
                if sample is not None and len(sample) > 200:
                    sample = sample[:197] + "..."
                entry["sample_value"] = sample

    def to_report(self, input_dir: Path) -> dict:
        unknown_by_table: dict[str, dict[str, dict]] = {}
        total_unknown_fields = 0
        for table, fields in self._by_table.items():
            if not fields:
                continue
            table_report: dict[str, dict] = {}
            for field_name, entry in sorted(fields.items()):
                table_report[field_name] = {
                    "count": entry["count"],
                    "sample_value": entry["sample_value"],
                    "rename_target_suggestion": _to_snake_case(field_name),
                }
            unknown_by_table[table] = table_report
            total_unknown_fields += len(table_report)

        return {
            "scan_time": _now_iso(),
            "input_dir": str(input_dir),
            "summary": {
                "tables_with_drift": len(unknown_by_table),
                "total_unknown_fields": total_unknown_fields,
            },
            "unknown_fields_by_table": unknown_by_table,
        }

    def has_drift(self) -> bool:
        return any(fields for fields in self._by_table.values())


def write_schema_drift_report(output_dir: Path, tracker: SchemaDriftTracker, input_dir: Path) -> Path | None:
    """Write SchemaDrift.json next to the parquet output. Returns the path or None."""
    if not tracker.has_drift():
        return None
    output_dir.mkdir(parents=True, exist_ok=True)
    report_path = output_dir / "SchemaDrift.json"
    report = tracker.to_report(input_dir)
    report_path.write_text(json.dumps(report, indent=2), encoding="utf-8")
    print(
        f"  Schema drift: {report['summary']['total_unknown_fields']} unknown field(s) "
        f"across {report['summary']['tables_with_drift']} table(s) -> {report_path}"
    )
    return report_path
