"""CLI entry point that orchestrates the pipeline stages."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .activities import process_activities
from .content import process_content
from .helpers import _run_stamp
from .policy import (
    process_dlp_policies,
    process_rbac,
    process_retention_labels,
    process_sensitivity_labels,
)
from .users import process_users_csv
from .writers import (
    PYARROW_IMPORT_ERROR,
    _records_to_table,
    write_c8_tuning_manifest,
    write_hive_partitioned,
    write_parquet,
)


def main() -> int:
    if PYARROW_IMPORT_ERROR is not None:
        missing_module = getattr(PYARROW_IMPORT_ERROR, "name", "pyarrow")
        print(
            f"ERROR: Missing Python dependency '{missing_module}'. "
            "Install runtime dependencies with `pip install -r requirements.txt`.",
            file=sys.stderr,
        )
        return 2

    parser = argparse.ArgumentParser(
        description="Convert Compl8CmdletExport JSON output to unified Parquet format"
    )
    parser.add_argument(
        "--input-dir", required=True,
        help="Path to Export-YYYYMMDD-HHMMSS directory"
    )
    parser.add_argument(
        "--output-dir",
        help="Target directory for Parquet output (default: <input-dir>/C8TuningInput)"
    )
    parser.add_argument(
        "--users-csv", action="append", default=[],
        help="Path to a GAL Scraper or Entra user export CSV (can be specified multiple times)"
    )
    args = parser.parse_args()

    input_dir = Path(args.input_dir).resolve()
    output_dir = Path(args.output_dir).resolve() if args.output_dir else input_dir / "C8TuningInput"
    stamp = _run_stamp(input_dir)

    if not input_dir.exists():
        print(f"ERROR: Input directory does not exist: {input_dir}", file=sys.stderr)
        return 1

    print(f"Input:  {input_dir}")
    print(f"Output: {output_dir}")
    print(f"Run:    {stamp}")
    print()

    wrote_any = False
    row_counts: dict[str, int] = {}

    # --- Activity Explorer ---
    print("Processing Activity Explorer...")
    activities, sit_matches, policy_matches, email_details = process_activities(input_dir)

    if activities:
        print("  Writing activities (Hive-partitioned)...")
        wrote_activities = write_hive_partitioned(
            activities, output_dir / "activities", stamp
        )
        wrote_any |= wrote_activities
        if wrote_activities:
            row_counts["activities"] = len(activities)

    if sit_matches:
        table = _records_to_table(sit_matches)
        wrote_activity_sit = write_parquet(
            table,
            output_dir / "activity_sit_matches" / "source=cmdletexport" / f"{stamp}.parquet",
        )
        wrote_any |= wrote_activity_sit
        if wrote_activity_sit:
            row_counts["activity_sit_matches"] = len(sit_matches)

    if policy_matches:
        table = _records_to_table(policy_matches)
        wrote_activity_policy = write_parquet(
            table,
            output_dir / "activity_policy_matches" / "source=cmdletexport" / f"{stamp}.parquet",
        )
        wrote_any |= wrote_activity_policy
        if wrote_activity_policy:
            row_counts["activity_policy_matches"] = len(policy_matches)

    if email_details:
        table = _records_to_table(email_details)
        wrote_activity_email = write_parquet(
            table,
            output_dir / "activity_email_details" / "source=cmdletexport" / f"{stamp}.parquet",
        )
        wrote_any |= wrote_activity_email
        if wrote_activity_email:
            row_counts["activity_email_details"] = len(email_details)

    # --- Content Explorer ---
    print("\nProcessing Content Explorer...")
    content_files, content_sit_detections = process_content(input_dir)
    if content_files:
        table = _records_to_table(content_files)
        wrote_content_files = write_parquet(
            table,
            output_dir / "content" / "content_files" / "source=cmdletexport" / f"{stamp}.parquet",
        )
        wrote_any |= wrote_content_files
        if wrote_content_files:
            row_counts["content_files"] = len(content_files)
    if content_sit_detections:
        table = _records_to_table(content_sit_detections)
        wrote_content_sits = write_parquet(
            table,
            output_dir / "content" / "sit_detections" / "source=cmdletexport" / f"{stamp}.parquet",
        )
        wrote_any |= wrote_content_sits
        if wrote_content_sits:
            row_counts["content_sit_detections"] = len(content_sit_detections)

    # --- Policy configs ---
    print("\nProcessing policy configs...")
    dlp_policies, dlp_rules = process_dlp_policies(input_dir)
    if dlp_policies:
        wrote_dlp_policies = write_parquet(
            _records_to_table(dlp_policies),
            output_dir / "policy" / "dlp_policies.parquet",
        )
        wrote_any |= wrote_dlp_policies
        if wrote_dlp_policies:
            row_counts["dlp_policies"] = len(dlp_policies)
    if dlp_rules:
        wrote_dlp_rules = write_parquet(
            _records_to_table(dlp_rules),
            output_dir / "policy" / "dlp_rules.parquet",
        )
        wrote_any |= wrote_dlp_rules
        if wrote_dlp_rules:
            row_counts["dlp_rules"] = len(dlp_rules)

    sens_labels = process_sensitivity_labels(input_dir)
    if sens_labels:
        wrote_sens_labels = write_parquet(
            _records_to_table(sens_labels),
            output_dir / "policy" / "sensitivity_labels.parquet",
        )
        wrote_any |= wrote_sens_labels
        if wrote_sens_labels:
            row_counts["sensitivity_labels"] = len(sens_labels)

    ret_labels = process_retention_labels(input_dir)
    if ret_labels:
        wrote_ret_labels = write_parquet(
            _records_to_table(ret_labels),
            output_dir / "policy" / "retention_labels.parquet",
        )
        wrote_any |= wrote_ret_labels
        if wrote_ret_labels:
            row_counts["retention_labels"] = len(ret_labels)

    rbac = process_rbac(input_dir)
    if rbac:
        wrote_rbac = write_parquet(
            _records_to_table(rbac),
            output_dir / "policy" / "rbac_role_groups.parquet",
        )
        wrote_any |= wrote_rbac
        if wrote_rbac:
            row_counts["rbac_role_groups"] = len(rbac)

    # --- Users ---
    if args.users_csv:
        print("\nProcessing user identity data...")
        all_users: list[dict] = []
        for csv_path_str in args.users_csv:
            all_users.extend(process_users_csv(Path(csv_path_str).resolve()))
        if all_users:
            wrote_users = write_parquet(
                _records_to_table(all_users),
                output_dir / "identity" / "users" / "source=cmdletexport" / "users.parquet",
            )
            wrote_any |= wrote_users
            if wrote_users:
                row_counts["users"] = len(all_users)

    # --- Summary ---
    print()
    if wrote_any:
        write_c8_tuning_manifest(output_dir, input_dir, stamp, row_counts, args.users_csv)
        print("Unified Parquet export complete.")
        print(f"C8 tuning input root: {output_dir}")
        print(f"  content_files:  {output_dir / 'content' / 'content_files'}")
        print(f"  sit_detections: {output_dir / 'content' / 'sit_detections'}")
    else:
        print("No data found to export. Check that the input directory contains JSON export files.")

    return 0
