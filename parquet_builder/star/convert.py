"""CLI: convert an Activity Explorer export to the star-schema v6 parquet model.

Usage:
    py -m parquet_builder.star.convert --input-dir <export-dir>
        [--output-dir <dir>] [--risk-workbook X.xlsx] [--department-csv Y.csv]
        [--archive-raw | --no-archive-raw] [--allow-unenriched]
        [--derive-target-domain | --no-derive-target-domain]
        [--sit-exclusions <json>] [--batch-size N]

Default output: <input-dir>/PowerBI-AE-Parquet-v6. Writes schema.json (from
the SSOT), manifest.json (row counts, enrichment provenance, exclusion and
drift summaries) and SchemaDrift.json (when unknown raw keys were seen).
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from ..helpers import _now_iso
from ..schema_drift import SchemaDriftTracker, write_schema_drift_report
from .enrich import (
    EnrichmentError,
    RiskLookup,
    load_department_mapping,
    load_risk_workbook,
    resolve_enrichment_inputs,
)
from .pipeline import StarPipeline
from .schema import SCHEMA_PROFILE, SCHEMA_VERSION, emit_schema_json

_DEFAULT_OUTPUT_NAME = "PowerBI-AE-Parquet-v6"
_DEFAULT_EXCLUSIONS = Path(__file__).resolve().parents[2] / "ConfigFiles" / "AEStarSITExclusions.json"


def load_sit_exclusions(path: Path | None) -> tuple[list[str], Path | None]:
    """Load excluded SIT names; explicit path must exist, default is optional."""
    if path is not None:
        if not path.exists():
            raise EnrichmentError(f"--sit-exclusions does not exist: {path}")
        target = path
    elif _DEFAULT_EXCLUSIONS.exists():
        target = _DEFAULT_EXCLUSIONS
    else:
        print("  WARNING: no SIT exclusion config found; no SIT matches will be excluded.")
        return [], None
    with open(target, "r", encoding="utf-8-sig") as handle:
        payload = json.load(handle)
    names = payload.get("ExcludedSITNames") if isinstance(payload, dict) else payload
    if not isinstance(names, list):
        raise EnrichmentError(f"SIT exclusion config has no ExcludedSITNames list: {target}")
    return [str(name) for name in names], target


def convert(input_dir: Path, output_dir: Path | None = None, *,
            risk_workbook: Path | None = None, department_csv: Path | None = None,
            archive_raw: bool = True, allow_unenriched: bool = False,
            derive_target_domain: bool = True, sit_exclusions: Path | None = None,
            batch_size: int = 50_000) -> dict:
    """Run the conversion; returns the manifest dict."""
    input_dir = input_dir.resolve()
    if output_dir is None:
        output_dir = input_dir / _DEFAULT_OUTPUT_NAME
    output_dir = output_dir.resolve()

    risk_path, dept_path = resolve_enrichment_inputs(
        input_dir, risk_workbook, department_csv, allow_unenriched
    )
    if risk_path is None or dept_path is None:
        print("=" * 70)
        print("WARNING: producing an UNENRICHED model (--allow-unenriched).")
        print("  Risk scores will be 0 and departments unmapped. Do not ship")
        print("  this output to reporting consumers.")
        print("=" * 70)

    risk = load_risk_workbook(risk_path) if risk_path else RiskLookup()
    department_mappings = load_department_mapping(dept_path) if dept_path else {}
    excluded_names, exclusions_path = load_sit_exclusions(sit_exclusions)

    primary_mappings = sum(
        1 for entry in department_mappings.values() if not entry.get("is_alias"))
    alias_mappings = len(department_mappings) - primary_mappings

    print(f"Input:      {input_dir}")
    print(f"Output:     {output_dir}")
    print(f"Risk:       {risk_path} ({len(risk.rows)} SIT reference rows)")
    print(f"GAL:        {dept_path} ({primary_mappings} user mappings, "
          f"{alias_mappings} mail aliases)")
    print(f"Exclusions: {exclusions_path} ({len(excluded_names)} SIT names)")
    print()

    drift_tracker = SchemaDriftTracker()
    pipeline = StarPipeline(
        input_dir=input_dir, output_dir=output_dir, risk=risk,
        department_mappings=department_mappings,
        excluded_sit_names=excluded_names, archive_raw=archive_raw,
        derive_domains=derive_target_domain, batch_size=batch_size,
        drift_tracker=drift_tracker,
    )
    stats = pipeline.run()

    emit_schema_json(output_dir / "schema.json")
    drift_path = write_schema_drift_report(output_dir, drift_tracker, input_dir)
    drift_report = drift_tracker.to_report(input_dir) if drift_tracker.has_drift() else None

    enrichment = None
    if risk_path is not None or dept_path is not None:
        enrichment = {
            "risk_workbook": str(risk_path) if risk_path else None,
            "department_csv": str(dept_path) if dept_path else None,
            "sit_reference_rows": len(risk.rows),
            "department_mappings": primary_mappings,
            "department_mail_aliases": alias_mappings,
        }

    manifest = {
        "schema_version": SCHEMA_VERSION,
        "profile": SCHEMA_PROFILE,
        "producer": "Compl8CmdletExport parquet_builder.star",
        "generated_at_utc": _now_iso(),
        "source_export_dir": str(input_dir),
        "raw_records_scanned": stats["raw_records_scanned"],
        "duplicates_skipped": stats["duplicates_skipped"],
        "missing_record_identity": stats["missing_record_identity"],
        "row_counts": stats["row_counts"],
        "enrichment": enrichment,
        "sit_exclusions": {
            "config": str(exclusions_path) if exclusions_path else None,
            "excluded_sit_names": len(excluded_names),
            "sit_rows_before_exclusions": stats["sit_rows_before_exclusions"],
            "excluded_sit_match_rows": stats["excluded_sit_rows"],
        },
        "schema_drift": {
            "report_file": drift_path.name if drift_path else None,
            "tables_with_drift": (drift_report or {}).get("summary", {}).get("tables_with_drift", 0),
            "total_unknown_fields": (drift_report or {}).get("summary", {}).get("total_unknown_fields", 0),
        },
        "options": {
            "archive_raw": archive_raw,
            "derive_target_domain": derive_target_domain,
        },
        "notes": [
            "Overview visuals should use the agg_*_sit_day tables.",
            "fact_activity metrics (activity_risk_score, sit_type_count) include excluded SITs;"
            " exclusions filter fact_activity_sit and agg tables only.",
            "archive_raw is pipeline-only; do not load it into Power BI.",
        ],
    }
    (output_dir / "manifest.json").write_text(
        json.dumps(manifest, indent=2), encoding="utf-8"
    )
    return manifest


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Convert an Activity Explorer export to star-schema v6 parquet."
    )
    parser.add_argument("--input-dir", required=True,
                        help="Export root containing Data/ActivityExplorer pages.")
    parser.add_argument("--output-dir", default=None,
                        help=f"Output folder (default: <input-dir>/{_DEFAULT_OUTPUT_NAME}).")
    parser.add_argument("--risk-workbook", default=None,
                        help="SIT risk workbook (.xlsx). Searched in the export root and one level below when omitted.")
    parser.add_argument("--department-csv", default=None,
                        help="GAL/department mapping CSV/XLSX. Searched like the risk workbook when omitted.")
    parser.add_argument("--archive-raw", action=argparse.BooleanOptionalAction, default=True,
                        help="Write archive_raw.parquet with raw nested payloads (default: on).")
    parser.add_argument("--allow-unenriched", action="store_true",
                        help="Proceed without enrichment inputs (logs a prominent warning).")
    parser.add_argument("--derive-target-domain", action=argparse.BooleanOptionalAction, default=True,
                        help="Derive target/originating domains when the export lacks them (default: on).")
    parser.add_argument("--sit-exclusions", default=None,
                        help="SIT exclusion JSON (default: ConfigFiles/AEStarSITExclusions.json).")
    parser.add_argument("--batch-size", type=int, default=50_000,
                        help="Rows buffered per parquet write.")
    args = parser.parse_args(argv)

    input_dir = Path(args.input_dir)
    if not input_dir.exists():
        print(f"ERROR: input directory does not exist: {input_dir}", file=sys.stderr)
        return 1

    try:
        manifest = convert(
            input_dir,
            Path(args.output_dir) if args.output_dir else None,
            risk_workbook=Path(args.risk_workbook) if args.risk_workbook else None,
            department_csv=Path(args.department_csv) if args.department_csv else None,
            archive_raw=args.archive_raw,
            allow_unenriched=args.allow_unenriched,
            derive_target_domain=args.derive_target_domain,
            sit_exclusions=Path(args.sit_exclusions) if args.sit_exclusions else None,
            batch_size=max(1, args.batch_size),
        )
    except (EnrichmentError, FileNotFoundError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1

    print()
    print("Star-schema v6 parquet export complete:")
    for table_name in sorted(manifest["row_counts"]):
        print(f"  {table_name}: {manifest['row_counts'][table_name]:,}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
