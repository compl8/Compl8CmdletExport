"""
Unified Parquet Export - Compl8CmdletExport Post-Processor

Reads JSON output from Export-Compl8Configuration.ps1 and writes unified
Hive-partitioned Parquet files to a target directory.

Usage:
    python build_unified_parquet.py --input-dir Output/Export-20260307-123456

This module is a thin CLI shim; the implementation lives in the parquet_builder
package. Existing import paths are preserved via re-exports from
parquet_builder.__init__.
"""

from __future__ import annotations

import sys

from parquet_builder import (  # noqa: F401 - re-exports for backward compatibility
    ACTIVITY_NESTED_FIELDS,
    ACTIVITY_RENAMES,
    CE_METADATA_FIELDS,
    CONTENT_RENAMES,
    EGRESS_ACTIVITIES,
    PARQUET_WRITE_OPTS,
    POLICY_MATCH_RENAMES,
    SERVICE_ACCOUNT_PATTERNS,
    SIT_DETECTION_RENAMES,
    USER_RENAMES,
    find_ae_pages,
    find_ce_pages,
    load_json_config,
    load_page_records,
    main,
    process_activities,
    process_content,
    process_dlp_policies,
    process_rbac,
    process_retention_labels,
    process_sensitivity_labels,
    process_users_csv,
    write_c8_tuning_manifest,
    write_hive_partitioned,
    write_parquet,
)

if __name__ == "__main__":
    sys.exit(main())
