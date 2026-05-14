"""Unified Parquet export pipeline split into focused modules.

Backwards-compatible re-exports for callers that imported from the original
single-file build_unified_parquet.py.
"""

from .activities import process_activities
from .constants import (
    ACTIVITY_NESTED_FIELDS,
    ACTIVITY_RENAMES,
    CE_METADATA_FIELDS,
    CONTENT_RENAMES,
    EGRESS_ACTIVITIES,
    POLICY_MATCH_RENAMES,
    SERVICE_ACCOUNT_PATTERNS,
    SIT_DETECTION_RENAMES,
)
from .content import process_content
from .helpers import (
    _extract_file_name,
    _first_present,
    _now_iso,
    _parse_nested_json,
    _rename_record,
    _run_stamp,
    _safe_int,
    _safe_str,
    _sha1_text,
    _split_sit_ids,
)
from .loaders import (
    find_ae_pages,
    find_ce_pages,
    load_json_config,
    load_page_records,
)
from .main import main
from .policy import (
    process_dlp_policies,
    process_rbac,
    process_retention_labels,
    process_sensitivity_labels,
)
from .schema_drift import SchemaDriftTracker, write_schema_drift_report
from .users import USER_RENAMES, process_users_csv
from .writers import (
    PARQUET_WRITE_OPTS,
    _records_to_table,
    write_c8_tuning_manifest,
    write_hive_partitioned,
    write_parquet,
)

__all__ = [
    "ACTIVITY_NESTED_FIELDS",
    "ACTIVITY_RENAMES",
    "CE_METADATA_FIELDS",
    "CONTENT_RENAMES",
    "EGRESS_ACTIVITIES",
    "POLICY_MATCH_RENAMES",
    "PARQUET_WRITE_OPTS",
    "SERVICE_ACCOUNT_PATTERNS",
    "SIT_DETECTION_RENAMES",
    "SchemaDriftTracker",
    "USER_RENAMES",
    "find_ae_pages",
    "find_ce_pages",
    "load_json_config",
    "load_page_records",
    "main",
    "process_activities",
    "process_content",
    "process_dlp_policies",
    "process_rbac",
    "process_retention_labels",
    "process_sensitivity_labels",
    "process_users_csv",
    "write_c8_tuning_manifest",
    "write_hive_partitioned",
    "write_parquet",
    "write_schema_drift_report",
]
