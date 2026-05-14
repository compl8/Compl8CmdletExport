"""Activity Explorer page processing."""

from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path

from .constants import (
    ACTIVITY_NESTED_FIELDS,
    ACTIVITY_RENAMES,
    EGRESS_ACTIVITIES,
    POLICY_MATCH_RENAMES,
    SERVICE_ACCOUNT_PATTERNS,
    SIT_DETECTION_RENAMES,
)
from .helpers import (
    _extract_file_name,
    _now_iso,
    _parse_nested_json,
    _rename_record,
    _safe_str,
)
from .loaders import find_ae_pages, load_page_records


def process_activities(input_dir: Path, drift_tracker=None) -> tuple[list[dict], list[dict], list[dict], list[dict]]:
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
            if drift_tracker is not None:
                drift_tracker.record("activities", extra)

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
