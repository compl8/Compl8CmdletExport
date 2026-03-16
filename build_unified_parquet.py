"""
Unified Parquet Export — Compl8CmdletExport Post-Processor

Reads JSON output from Export-Compl8Configuration.ps1 and writes unified
Hive-partitioned Parquet files to a target directory.

Usage:
    python build_unified_parquet.py --input-dir Output/Export-20260307-123456 --output-dir C:/PurviewData
"""

from __future__ import annotations

import argparse
import csv
import json
import re
import sys
from datetime import datetime, timezone
from pathlib import Path

import pyarrow as pa
import pyarrow.parquet as pq

# ---------------------------------------------------------------------------
# Column rename maps (PascalCase cmdlet output -> snake_case unified schema)
# ---------------------------------------------------------------------------

ACTIVITY_RENAMES = {
    "RecordIdentity": "record_id",
    "Activity": "activity",
    "Timestamp": "happened_at",
    "Happened": "happened_at",  # API sometimes uses Happened
    "UserKey": "user_upn",
    "Workload": "workload",
    "DataPlatform": "data_platform",
    "UserType": "user_type",
    "EnforcementMode": "enforcement_mode",
    "FilePath": "file_path",
    "TargetFilePath": "target_file_path",
    "FileExtension": "file_extension",
    "FileType": "file_type",
    "FileSize": "file_size",
    "PreviousFileName": "previous_file_name",
    "Sha1": "sha1",
    "Sha256": "sha256",
    "SourceLocationType": "source_location_type",
    "DestinationLocationType": "destination_location_type",
    "ClientIP": "client_ip",
    "DeviceName": "device_name",
    "MDATPDeviceId": "mdatp_device_id",
    "Platform": "platform",
    "Application": "application",
    "TargetDomain": "target_domain",
    "TargetURL": "target_url",
    "PolicyId": "policy_id",
    "PolicyName": "policy_name",
    "PolicyVersion": "policy_version",
    "RuleName": "rule_name",
    "SensitivityLabelIds": "sensitivity_label_ids",
    "RmsEncrypted": "rms_encrypted",
    "JITTriggered": "jit_triggered",
    "ParentArchiveHash": "parent_archive_hash",
    "CopilotAppHost": "copilot_app_host",
    "CopilotThreadId": "copilot_thread_id",
    "AppIdentity": "app_identity",
    "AppIdentityCategory": "app_identity_category",
    "AppIdentityGroup": "app_identity_group",
    "PurviewAIAppName": "purview_ai_app_name",
    "AreFilesReferenced": "are_files_referenced",
    "AreSensitiveFilesReferenced": "are_sensitive_files_referenced",
    "HasWebSearchQuery": "has_web_search_query",
    "RmManufacturer": "rm_manufacturer",
    "RmModel": "rm_model",
    "RmSerialNumber": "rm_serial_number",
}

# Nested JSON fields that get exploded into separate tables (excluded from extra_fields)
ACTIVITY_NESTED_FIELDS = {"SensitiveInfoTypeData", "PolicyMatchInfo", "EmailInfo"}

SIT_DETECTION_RENAMES = {
    "SensitiveInfoTypeId": "sit_id",
    "Count": "match_count",
    "Confidence": "confidence_score",
    "ClassifierType": "classifier_type",
}

POLICY_MATCH_RENAMES = {
    "PolicyId": "policy_id",
    "PolicyName": "policy_name",
    "PolicyMode": "policy_mode",
    "RuleId": "rule_id",
    "RuleName": "rule_name",
    "RuleActions": "rule_actions",
    "Condition": "condition_json",
}

CONTENT_RENAMES = {
    "Name": "file_name",
    "FileName": "file_name",  # CE uses FileName
    "DocId": "doc_id",
    "SourceUrl": "source_url",
    "FileSourceUrl": "source_url",  # CE uses FileSourceUrl
    "FileUrl": "file_url",
    "Workload": "workload",
    "Location": "workload",  # CE uses Location (EXO/SPO/ODB/Teams)
    "FileType": "file_type",
    "DetectedLanguage": "detected_language",
    "SensitiveLabel": "sensitivity_label",
    "SensitivityLabel": "sensitivity_label",
    "RetentionLabel": "retention_label",
    "Title": "title",
    "UserCreated": "user_created",
    "UserModified": "user_modified",
    "LastModifiedTime": "last_modified_time",
    "SiteId": "site_id",
    "UniqueId": "unique_id",
    "SPFileId": "sp_file_id",
    "PreviewId": "preview_id",
    "SensitiveInfoTypeBucketsData": "matches_json",
    "SensitiveInfoTypesData": "matches_json",  # CE uses this name
    "SensitiveInfoTypes": "sensitive_info_type_ids",
    "TrainableClassifiers": "trainable_classifiers",
}

# CE metadata fields (added by export, not from API)
CE_METADATA_FIELDS = {"_ExportTagType", "_ExportTagName"}

# Activities that indicate egress
EGRESS_ACTIVITIES = {
    "FileUploaded", "FileCopiedToRemovableMedia", "FileCopiedToNetworkShare",
    "FileCopiedToCloud", "FileTransferredByAIP", "FilePrinted",
    "ContentExtractionAllowed", "AccessByUnallowedApp", "FileUploadedToCloud",
    "FileCopiedToClipboard", "BrowserUpload",
}

SERVICE_ACCOUNT_PATTERNS = re.compile(
    r"(?:^svc[_.-]|^service[_.-]|@.*\.onmicrosoft\.com$|^app@|^s-\d)",
    re.IGNORECASE,
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _now_iso() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _safe_str(val) -> str | None:
    if val is None:
        return None
    if isinstance(val, (list, dict)):
        return json.dumps(val, default=str)
    return str(val)


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


# ---------------------------------------------------------------------------
# JSON file discovery and loading
# ---------------------------------------------------------------------------

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
            return records
        return []
    if isinstance(data, list):
        return data
    return []


def load_json_config(path: Path) -> dict | None:
    """Load a JSON config export file."""
    if not path.exists():
        return None
    with open(path, "r", encoding="utf-8-sig") as f:
        return json.load(f)


# ---------------------------------------------------------------------------
# Activity Explorer processing
# ---------------------------------------------------------------------------

def process_activities(input_dir: Path) -> tuple[list[dict], list[dict], list[dict], list[dict]]:
    """Process AE pages -> (activities, sit_matches, policy_matches, email_details)."""
    pages = find_ae_pages(input_dir)
    if not pages:
        return [], [], [], []

    ingested_at = _now_iso()
    activities = []
    sit_matches = []
    policy_matches = []
    email_details = []

    for page_path in pages:
        records = load_page_records(page_path)
        for raw in records:
            record_id = raw.get("RecordIdentity")

            # Rename activity fields
            renamed, extra = _rename_record(raw, ACTIVITY_RENAMES, excluded_keys=ACTIVITY_NESTED_FIELDS)

            # Derived columns
            happened_at = renamed.get("happened_at")
            happened_date = None
            happened_hour = None
            if happened_at:
                try:
                    dt = datetime.fromisoformat(str(happened_at).replace("Z", "+00:00"))
                    happened_date = dt.strftime("%Y-%m-%d")
                    happened_hour = dt.hour
                except (ValueError, TypeError):
                    pass

            activity_val = renamed.get("activity", "")
            file_path_val = renamed.get("file_path")

            renamed["_source_tool"] = "cmdletexport"
            renamed["_ingested_at"] = ingested_at
            renamed["happened_date"] = happened_date
            renamed["happened_hour"] = happened_hour
            renamed["is_egress"] = activity_val in EGRESS_ACTIVITIES
            renamed["is_copilot"] = activity_val == "CopilotInteraction"
            renamed["is_dlp"] = raw.get("PolicyMatchInfo") is not None and raw.get("PolicyMatchInfo") != ""
            renamed["has_sensitive_data"] = raw.get("SensitiveInfoTypeData") is not None and raw.get("SensitiveInfoTypeData") != ""
            upn = renamed.get("user_upn", "") or ""
            renamed["is_service_account"] = bool(SERVICE_ACCOUNT_PATTERNS.search(upn))
            renamed["file_name"] = _extract_file_name(file_path_val)
            renamed["extra_fields"] = json.dumps(extra, default=str) if extra else None

            activities.append(renamed)

            # Explode SensitiveInfoTypeData
            sit_data = _parse_nested_json(raw.get("SensitiveInfoTypeData"))
            if sit_data:
                for sit in sit_data:
                    sit_row, _ = _rename_record(sit, SIT_DETECTION_RENAMES)
                    sit_row["record_id"] = record_id
                    sit_row["_source_tool"] = "cmdletexport"
                    sit_row["_ingested_at"] = ingested_at
                    sit_matches.append(sit_row)

            # Explode PolicyMatchInfo
            policy_data = _parse_nested_json(raw.get("PolicyMatchInfo"))
            if policy_data:
                for pm in policy_data:
                    pm_row, _ = _rename_record(pm, POLICY_MATCH_RENAMES)
                    pm_row["record_id"] = record_id
                    pm_row["_source_tool"] = "cmdletexport"
                    pm_row["_ingested_at"] = ingested_at
                    # Serialize complex fields
                    if "rule_actions" in pm_row and isinstance(pm_row["rule_actions"], (list, dict)):
                        pm_row["rule_actions"] = json.dumps(pm_row["rule_actions"], default=str)
                    if "condition_json" in pm_row and isinstance(pm_row["condition_json"], (list, dict)):
                        pm_row["condition_json"] = json.dumps(pm_row["condition_json"], default=str)
                    policy_matches.append(pm_row)

            # Explode EmailInfo
            email_data = _parse_nested_json(raw.get("EmailInfo"))
            if email_data:
                for em in email_data:
                    em_row = {k: _safe_str(v) for k, v in em.items()}
                    em_row["record_id"] = record_id
                    em_row["_source_tool"] = "cmdletexport"
                    em_row["_ingested_at"] = ingested_at
                    email_details.append(em_row)

    print(f"  Activities: {len(activities)} records, "
          f"{len(sit_matches)} SIT matches, "
          f"{len(policy_matches)} policy matches, "
          f"{len(email_details)} email details")

    return activities, sit_matches, policy_matches, email_details


# ---------------------------------------------------------------------------
# Content Explorer processing
# ---------------------------------------------------------------------------

def process_content(input_dir: Path) -> list[dict]:
    """Process CE pages -> content_files list."""
    pages = find_ce_pages(input_dir)
    if not pages:
        return []

    ingested_at = _now_iso()
    content_files = []

    for page_path in pages:
        records = load_page_records(page_path)

        # Try to get tag info from page wrapper
        page_tag_type = None
        page_tag_name = None
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

            # Add tag metadata from record or page wrapper
            renamed["tag_type"] = raw.get("_ExportTagType") or page_tag_type
            renamed["tag_name"] = raw.get("_ExportTagName") or page_tag_name
            renamed["_source_tool"] = "cmdletexport"
            renamed["_ingested_at"] = ingested_at

            # Serialize matches_json if it's still a complex type
            if "matches_json" in renamed and isinstance(renamed["matches_json"], (list, dict)):
                renamed["matches_json"] = json.dumps(renamed["matches_json"], default=str)

            renamed["extra_fields"] = json.dumps(extra, default=str) if extra else None

            content_files.append(renamed)

    print(f"  Content files: {len(content_files)} records")
    return content_files


# ---------------------------------------------------------------------------
# Policy config processing
# ---------------------------------------------------------------------------

def process_dlp_policies(input_dir: Path) -> tuple[list[dict], list[dict]]:
    """Process DLP-Config.json -> (policies, rules)."""
    data = load_json_config(input_dir / "DLP-Config.json")
    if not data:
        return [], []

    ingested_at = _now_iso()
    policies = []
    rules = []

    for p in data.get("Policies", []):
        row = {k: _safe_str(v) for k, v in p.items()}
        row["_source_tool"] = "cmdletexport"
        row["_ingested_at"] = ingested_at
        policies.append(row)

    for r in data.get("Rules", []):
        row = {k: _safe_str(v) for k, v in r.items()}
        row["_source_tool"] = "cmdletexport"
        row["_ingested_at"] = ingested_at
        rules.append(row)

    print(f"  DLP: {len(policies)} policies, {len(rules)} rules")
    return policies, rules


def process_sensitivity_labels(input_dir: Path) -> list[dict]:
    data = load_json_config(input_dir / "SensitivityLabels-Config.json")
    if not data:
        return []

    ingested_at = _now_iso()
    labels = []
    for lbl in data.get("Labels", []):
        row = {k: _safe_str(v) for k, v in lbl.items()}
        row["_source_tool"] = "cmdletexport"
        row["_ingested_at"] = ingested_at
        labels.append(row)

    print(f"  Sensitivity labels: {len(labels)} records")
    return labels


def process_retention_labels(input_dir: Path) -> list[dict]:
    data = load_json_config(input_dir / "RetentionLabels-Config.json")
    if not data:
        return []

    ingested_at = _now_iso()
    labels = []
    for lbl in data.get("Labels", []):
        row = {k: _safe_str(v) for k, v in lbl.items()}
        row["_source_tool"] = "cmdletexport"
        row["_ingested_at"] = ingested_at
        labels.append(row)

    print(f"  Retention labels: {len(labels)} records")
    return labels


def process_rbac(input_dir: Path) -> list[dict]:
    data = load_json_config(input_dir / "RBAC-Config.json")
    if not data:
        return []

    ingested_at = _now_iso()
    groups = []
    for rg in data.get("RoleGroups", []):
        row = {k: _safe_str(v) for k, v in rg.items()}
        row["_source_tool"] = "cmdletexport"
        row["_ingested_at"] = ingested_at
        groups.append(row)

    print(f"  RBAC role groups: {len(groups)} records")
    return groups


# ---------------------------------------------------------------------------
# User identity processing (GAL Scraper / Entra exports)
# ---------------------------------------------------------------------------

# Column renames for user CSVs — handles both GAL Scraper and Entra export field names
USER_RENAMES = {
    # Core identity
    "DisplayName": "display_name",
    "displayName": "display_name",
    "UserPrincipalName": "user_upn",
    "userPrincipalName": "user_upn",
    "Mail": "mail",
    "mail": "mail",
    "AccountEnabled": "account_enabled",
    "accountEnabled": "account_enabled",
    # Org structure
    "Department": "department",
    "department": "department",
    "JobTitle": "job_title",
    "jobTitle": "job_title",
    "CompanyName": "company_name",
    "companyName": "company_name",
    # Location
    "City": "city",
    "city": "city",
    "State": "state",
    "state": "state",
    "OfficeLocation": "office_location",
    "officeLocation": "office_location",
    "OfficeCity": "office_city",
    "UsageLocation": "usage_location",
    "usageLocation": "usage_location",
    # Manager
    "ManagerDisplayName": "manager_display_name",
    "ManagerUPN": "manager_upn",
    "ManagerMail": "manager_mail",
    # Metadata
    "CreatedDateTime": "created_at",
    "createdDateTime": "created_at",
    "AccountType": "account_type",
    "userType": "account_type",
    # On-premises
    "OnPremisesDN": "on_premises_dn",
    "onPremisesDistinguishedName": "on_premises_dn",
    "OnPremisesSamAccount": "on_premises_sam_account",
    "onPremisesSamAccountName": "on_premises_sam_account",
    "OnPremisesDomain": "on_premises_domain",
    "onPremisesDomainName": "on_premises_domain",
    "ADOrgUnit": "ad_org_unit",
    # GAL-specific enrichment
    "AgencyCode": "agency_code",
    "Branch": "branch",
    "RegionOrBU": "region_or_bu",
    "Division": "division",
    "SubBranch": "sub_branch",
    "HistoricalDept": "historical_dept",
    "LastLogonInfo": "last_logon_info",
    # Entra ID
    "id": "entra_id",
    "Id": "entra_id",
}

# Extension attributes (GAL Clean export)
for _i in range(1, 16):
    USER_RENAMES[f"ExtAttr{_i}"] = f"ext_attr_{_i}"
    USER_RENAMES[f"extension_{_i}"] = f"ext_attr_{_i}"


def process_users_csv(csv_path: Path) -> list[dict]:
    """Process a GAL Scraper or Entra user export CSV into unified user records."""
    if not csv_path.exists():
        print(f"  WARNING: Users CSV not found: {csv_path}")
        return []

    ingested_at = _now_iso()
    users = []

    with open(csv_path, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        for raw in reader:
            renamed, extra = _rename_record(raw, USER_RENAMES)
            renamed["_source_tool"] = "cmdletexport"
            renamed["_ingested_at"] = ingested_at

            upn = renamed.get("user_upn", "") or ""
            renamed["is_service_account"] = bool(SERVICE_ACCOUNT_PATTERNS.search(upn))
            renamed["extra_fields"] = json.dumps(extra, default=str) if extra else None

            users.append(renamed)

    print(f"  Users: {len(users)} records from {csv_path.name}")
    return users


# ---------------------------------------------------------------------------
# Parquet writing
# ---------------------------------------------------------------------------

PARQUET_WRITE_OPTS = {
    "compression": "ZSTD",
    "compression_level": 3,
    "use_dictionary": True,
    "write_statistics": True,
}


def _records_to_table(records: list[dict]) -> pa.Table | None:
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
    int_cols = {"happened_hour", "file_size", "match_count", "confidence_score"}

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


def write_parquet(table: pa.Table | None, output_path: Path) -> bool:
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


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> int:
    parser = argparse.ArgumentParser(
        description="Convert Compl8CmdletExport JSON output to unified Parquet format"
    )
    parser.add_argument(
        "--input-dir", required=True,
        help="Path to Export-YYYYMMDD-HHMMSS directory"
    )
    parser.add_argument(
        "--output-dir", default="C:/PurviewData",
        help="Target directory for Parquet output (default: C:/PurviewData)"
    )
    parser.add_argument(
        "--users-csv", action="append", default=[],
        help="Path to a GAL Scraper or Entra user export CSV (can be specified multiple times)"
    )
    args = parser.parse_args()

    input_dir = Path(args.input_dir).resolve()
    output_dir = Path(args.output_dir).resolve()
    stamp = _run_stamp(input_dir)

    if not input_dir.exists():
        print(f"ERROR: Input directory does not exist: {input_dir}", file=sys.stderr)
        return 1

    print(f"Input:  {input_dir}")
    print(f"Output: {output_dir}")
    print(f"Run:    {stamp}")
    print()

    wrote_any = False

    # --- Activity Explorer ---
    print("Processing Activity Explorer...")
    activities, sit_matches, policy_matches, email_details = process_activities(input_dir)

    if activities:
        print("  Writing activities (Hive-partitioned)...")
        wrote_any |= write_hive_partitioned(
            activities, output_dir / "activities", stamp
        )

    if sit_matches:
        table = _records_to_table(sit_matches)
        wrote_any |= write_parquet(
            table,
            output_dir / "activity_sit_matches" / "source=cmdletexport" / f"{stamp}.parquet",
        )

    if policy_matches:
        table = _records_to_table(policy_matches)
        wrote_any |= write_parquet(
            table,
            output_dir / "activity_policy_matches" / "source=cmdletexport" / f"{stamp}.parquet",
        )

    if email_details:
        table = _records_to_table(email_details)
        wrote_any |= write_parquet(
            table,
            output_dir / "activity_email_details" / "source=cmdletexport" / f"{stamp}.parquet",
        )

    # --- Content Explorer ---
    print("\nProcessing Content Explorer...")
    content_files = process_content(input_dir)
    if content_files:
        table = _records_to_table(content_files)
        wrote_any |= write_parquet(
            table,
            output_dir / "content" / "content_files" / "source=cmdletexport" / f"{stamp}.parquet",
        )

    # --- Policy configs ---
    print("\nProcessing policy configs...")
    dlp_policies, dlp_rules = process_dlp_policies(input_dir)
    if dlp_policies:
        wrote_any |= write_parquet(
            _records_to_table(dlp_policies),
            output_dir / "policy" / "dlp_policies.parquet",
        )
    if dlp_rules:
        wrote_any |= write_parquet(
            _records_to_table(dlp_rules),
            output_dir / "policy" / "dlp_rules.parquet",
        )

    sens_labels = process_sensitivity_labels(input_dir)
    if sens_labels:
        wrote_any |= write_parquet(
            _records_to_table(sens_labels),
            output_dir / "policy" / "sensitivity_labels.parquet",
        )

    ret_labels = process_retention_labels(input_dir)
    if ret_labels:
        wrote_any |= write_parquet(
            _records_to_table(ret_labels),
            output_dir / "policy" / "retention_labels.parquet",
        )

    rbac = process_rbac(input_dir)
    if rbac:
        wrote_any |= write_parquet(
            _records_to_table(rbac),
            output_dir / "policy" / "rbac_role_groups.parquet",
        )

    # --- Users ---
    if args.users_csv:
        print("\nProcessing user identity data...")
        all_users: list[dict] = []
        for csv_path_str in args.users_csv:
            all_users.extend(process_users_csv(Path(csv_path_str).resolve()))
        if all_users:
            wrote_any |= write_parquet(
                _records_to_table(all_users),
                output_dir / "identity" / "users" / "source=cmdletexport" / "users.parquet",
            )

    # --- Summary ---
    print()
    if wrote_any:
        print("Unified Parquet export complete.")
    else:
        print("No data found to export. Check that the input directory contains JSON export files.")

    return 0


if __name__ == "__main__":
    sys.exit(main())
