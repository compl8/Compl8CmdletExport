"""Dimension and aggregate materialization for the star pipeline."""

from __future__ import annotations

from datetime import date as date_type, timedelta
from pathlib import Path
from typing import Any

from .enrich import GUID_RE, RiskLookup
from .registry import IdRegistry
from .sinks import write_snapshot_table

AGGREGATE_KEYS = {
    "agg_department_sit_day": ("date_key", "department_id", "sit_key", "workload_id", "activity_type_id"),
    "agg_user_sit_day": ("date_key", "user_id", "department_id", "sit_key", "workload_id", "activity_type_id"),
    "agg_location_sit_day": ("date_key", "source_location_id", "sit_key", "workload_id", "activity_type_id"),
    "agg_activity_type_sit_day": ("date_key", "activity_type_id", "sit_key", "workload_id"),
    "agg_domain_sit_day": ("date_key", "target_domain_id", "sit_key", "workload_id", "activity_type_id"),
}


def build_date_rows(dates: dict[int, date_type]) -> list[dict[str, Any]]:
    """Continuous calendar rows from min..max observed dates (no holes)."""
    rows: list[dict[str, Any]] = []
    if not dates:
        return rows
    current = min(dates.values())
    last = max(dates.values())
    while current <= last:
        rows.append({
            "date_key": int(current.strftime("%Y%m%d")),
            "date": current,
            "year": current.year,
            "month": current.month,
            "month_name": current.strftime("%B"),
            "month_short": current.strftime("%b"),
            "quarter": (current.month - 1) // 3 + 1,
            "week_of_year": current.isocalendar()[1],
            "day_of_week": current.strftime("%A"),
            "day_of_week_num": current.weekday(),
            "is_weekend": current.weekday() >= 5,
        })
        current += timedelta(days=1)
    return rows


_NAME_SOURCE_SHEETS = {
    "raw_payload": "Generated (name from AE raw payload)",
    "tenant_map": "Generated (name from tenant SIT map)",
}


def build_sit_rows(
    risk: RiskLookup, seen_sit_keys: set[str],
    resolved_sit_names: dict[str, tuple[str, str]] | None = None,
) -> list[dict[str, Any]]:
    """All workbook rows (observed flag) plus generated rows for unknown SITs.

    Generated rows take their display name from ``resolved_sit_names``
    (raw-payload / tenant-map resolution recorded by the pipeline) and fall
    back to the GUID; source_sheet records which source named the row.
    """
    resolved_sit_names = resolved_sit_names or {}
    sit_rows_by_key: dict[str, dict[str, Any]] = {}
    for row in risk.rows:
        dim_row = dict(row)
        dim_row["observed"] = row["sit_key"] in seen_sit_keys
        sit_rows_by_key.setdefault(row["sit_key"], dim_row)
    for sit_key in sorted(seen_sit_keys):
        if sit_key in sit_rows_by_key:
            continue
        sit_id = sit_key if GUID_RE.match(sit_key) else None
        name, source = resolved_sit_names.get(sit_key, (None, None))
        if name:
            source_sheet = _NAME_SOURCE_SHEETS.get(
                source, "Generated from Activity Explorer export")
        else:
            name = sit_key if sit_key != "unknown" else "Unknown SIT"
            source_sheet = "Generated from Activity Explorer export"
        sit_rows_by_key[sit_key] = {
            "sit_key": sit_key,
            "sit_name": name,
            "sit_id": sit_id,
            "sit_slug": None if sit_id else sit_key,
            "risk_band": "Unrated",
            "source_sheet": source_sheet,
            "is_unrated": True,
            "observed": True,
        }
    return [sit_rows_by_key[key] for key in sorted(sit_rows_by_key)]


def write_dimensions(output_dir: Path, registry: IdRegistry,
                     department_mappings: dict[str, dict[str, Any]],
                     dates: dict[int, date_type], risk: RiskLookup,
                     seen_sit_keys: set[str],
                     resolved_sit_names: dict[str, tuple[str, str]] | None = None,
                     ) -> dict[str, int]:
    """Materialize every dimension table; returns row counts."""
    # Union the full GAL population into dim_user (has_activity=False).
    # Mail-alias keys are skipped: one dim_user row per person (UPN), not one
    # per address — aliases exist only so activity records that identify the
    # user by primary SMTP resolve to the same department mapping.
    for upn_key, mapping in department_mappings.items():
        if mapping.get("is_alias"):
            continue
        registry.get_user(upn_key, department_mappings, has_activity=False)

    dim_rows = {
        "dim_date": build_date_rows(dates),
        "dim_department": [registry.department_rows[k] for k in sorted(registry.department_rows)],
        "dim_user": [registry.user_rows[k] for k in sorted(registry.user_rows)],
        "dim_sit": build_sit_rows(risk, seen_sit_keys, resolved_sit_names),
        "dim_activity_type": [registry.activity_type_rows[k] for k in sorted(registry.activity_type_rows)],
        "dim_workload": [registry.workload_rows[k] for k in sorted(registry.workload_rows)],
        "dim_location": [registry.location_rows[k] for k in sorted(registry.location_rows)],
        "dim_file": [registry.file_rows[k] for k in sorted(registry.file_rows)],
        "dim_policy": [registry.policy_rows[k] for k in sorted(registry.policy_rows)],
        "dim_domain": [registry.domain_rows[k] for k in sorted(registry.domain_rows)],
        "dim_email_address": [registry.email_rows[k] for k in sorted(registry.email_rows)],
        "dim_app_identity": [registry.app_identity_rows[k] for k in sorted(registry.app_identity_rows)],
        "dim_source_page": [registry.page_rows[k] for k in sorted(registry.page_rows)],
    }
    return {
        name: write_snapshot_table(output_dir, name, rows)
        for name, rows in dim_rows.items()
    }


def write_aggregates(output_dir: Path,
                     aggregates: dict[str, dict[tuple, dict[str, int]]]) -> dict[str, int]:
    """Materialize the five agg_*_sit_day tables; returns row counts."""
    counts = {}
    for table_name, key_names in AGGREGATE_KEYS.items():
        rows = []
        for key, metrics in aggregates[table_name].items():
            row = dict(zip(key_names, key))
            row.update(metrics)
            rows.append(row)
        counts[table_name] = write_snapshot_table(output_dir, table_name, rows)
    return counts
