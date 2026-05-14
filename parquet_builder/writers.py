"""PyArrow table construction and Parquet writers."""

from __future__ import annotations

import json
from pathlib import Path

try:
    import pyarrow as pa
    import pyarrow.parquet as pq
except ModuleNotFoundError as exc:
    pa = None
    pq = None
    PYARROW_IMPORT_ERROR = exc
else:
    PYARROW_IMPORT_ERROR = None

from .helpers import _now_iso, _safe_str

PARQUET_WRITE_OPTS = {
    "compression": "ZSTD",
    "compression_level": 3,
    "use_dictionary": True,
    "write_statistics": True,
}


def _records_to_table(records: list[dict]):
    """Convert list of dicts to a PyArrow Table with all-string schema."""
    if not records:
        return None

    # Collect all keys across all records
    all_keys: list[str] = []
    seen: set[str] = set()
    for rec in records:
        for k in rec:
            if k not in seen:
                all_keys.append(k)
                seen.add(k)

    # Build columnar data — coerce everything to string for schema flexibility
    # except booleans and integers which we keep typed
    bool_cols = {"is_egress", "is_copilot", "is_dlp", "has_sensitive_data", "is_service_account"}
    int_cols = {
        "happened_hour",
        "file_size",
        "match_count",
        "confidence_score",
        "low_count",
        "medium_count",
        "high_count",
        "total_count",
    }

    columns: dict[str, list] = {k: [] for k in all_keys}
    for rec in records:
        for k in all_keys:
            val = rec.get(k)
            if k in bool_cols:
                columns[k].append(bool(val) if val is not None else None)
            elif k in int_cols:
                if val is not None:
                    try:
                        columns[k].append(int(val))
                    except (ValueError, TypeError):
                        columns[k].append(None)
                else:
                    columns[k].append(None)
            else:
                columns[k].append(_safe_str(val))

    arrays = {}
    for k in all_keys:
        if k in bool_cols:
            arrays[k] = pa.array(columns[k], type=pa.bool_())
        elif k in int_cols:
            arrays[k] = pa.array(columns[k], type=pa.int64())
        else:
            arrays[k] = pa.array(columns[k], type=pa.string())

    return pa.table(arrays)


def write_parquet(table, output_path: Path) -> bool:
    """Write a PyArrow table to a single Parquet file."""
    if table is None or table.num_rows == 0:
        return False
    output_path.parent.mkdir(parents=True, exist_ok=True)
    pq.write_table(table, str(output_path), **PARQUET_WRITE_OPTS)
    print(f"  Wrote {table.num_rows} rows -> {output_path}")
    return True


def write_hive_partitioned(records: list[dict], base_dir: Path,
                           run_stamp: str, partition_key: str = "happened_date") -> bool:
    """Write activity records as Hive-partitioned Parquet (source=cmdletexport/year=YYYY/month=MM/)."""
    if not records:
        return False

    # Group by year/month from the partition key
    buckets: dict[tuple[str, str], list[dict]] = {}
    for rec in records:
        date_str = rec.get(partition_key)
        if date_str:
            try:
                parts = date_str.split("-")
                year, month = parts[0], parts[1]
            except (IndexError, AttributeError):
                year, month = "unknown", "unknown"
        else:
            year, month = "unknown", "unknown"
        key = (year, month)
        buckets.setdefault(key, []).append(rec)

    source_dir = base_dir / "source=cmdletexport"
    total_written = 0

    for (year, month), bucket_records in buckets.items():
        partition_dir = source_dir / f"year={year}" / f"month={month}"
        file_path = partition_dir / f"{run_stamp}.parquet"
        table = _records_to_table(bucket_records)
        if write_parquet(table, file_path):
            total_written += len(bucket_records)

    return total_written > 0


def write_c8_tuning_manifest(
    output_dir: Path,
    input_dir: Path,
    run_stamp: str,
    row_counts: dict[str, int],
    users_csv: list[str],
    schema_drift_path: Path | None = None,
) -> None:
    """Write a small manifest for downstream run pickers."""
    output_dir.mkdir(parents=True, exist_ok=True)
    manifest = {
        "manifest_version": 1,
        "schema_profile": "c8_tuning_input",
        "producer": "Compl8CmdletExport",
        "generated_at_utc": _now_iso(),
        "source_export_dir": str(input_dir),
        "run_stamp": run_stamp,
        "paths": {
            "content_files": "content/content_files",
            "sit_detections": "content/sit_detections",
            "activities": "activities",
            "activity_sit_matches": "activity_sit_matches",
            "activity_policy_matches": "activity_policy_matches",
            "activity_email_details": "activity_email_details",
            "users": "identity/users",
        },
        "capabilities": {
            "content_inventory": row_counts.get("content_files", 0) > 0,
            "sit_detections": row_counts.get("content_sit_detections", 0) > 0,
            "matched_value_enrichment": False,
        },
        "row_counts": row_counts,
        "users_csv": users_csv,
        "schema_drift_report": schema_drift_path.name if schema_drift_path is not None else None,
    }
    manifest_path = output_dir / "c8_tuning_input_manifest.json"
    manifest_path.write_text(json.dumps(manifest, indent=2), encoding="utf-8")
    print(f"  Wrote manifest -> {manifest_path}")
