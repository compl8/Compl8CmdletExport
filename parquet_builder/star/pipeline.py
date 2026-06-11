"""Star-schema v6 conversion pipeline: AE export pages -> parquet tables.

Combines the proven v5 optimized-converter logic with the enhanced legacy
fork's robustness fixes:

- F2 fix: page Records arriving as a single dict (PowerShell ConvertTo-Json
  unwraps one-element arrays) are processed, not dropped (records.load_page).
- Nothing silently dropped: every raw key without a typed home lands in
  fact_activity_detail.extra_json and the schema-drift report.
- RecordIdentity dedup keeps the first occurrence.
- SIT exclusion list applies to fact_activity_sit / agg tables only;
  activities (and their risk scores) are unaffected.
"""

from __future__ import annotations

import json
from collections import Counter, defaultdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from ..helpers import _safe_str
from ..loaders import find_ae_pages
from .enrich import RiskLookup, _norm_text, resolve_detected_sit
from .finalize import write_aggregates, write_dimensions
from .keys import stable_int_id
from .records import (
    APP_IDENTITY_KEYS,
    CONSUMED_KEYS,
    COPILOT_KEYS,
    as_bool,
    as_list,
    date_key_of,
    derive_target_domain,
    extract_domain_from_email,
    json_str,
    load_page,
    parse_dt,
    split_receivers,
)
from .registry import IdRegistry, activity_group_of, extract_folder
from .sinks import ParquetSink

_SINK_TABLES = (
    "fact_activity", "fact_activity_sit", "fact_policy_activity",
    "fact_email_recipient", "fact_email_detail", "fact_copilot_interaction",
    "fact_activity_detail", "activity_record_index",
)

# fact_activity_detail typed contract columns -> raw key fallbacks.
_DETAIL_CONTRACT = {
    "target_file_path": ("TargetFilePath",),
    "target_url": ("TargetUrl", "TargetURL"),
    "device_name": ("DeviceName",),
    "client_ip": ("ClientIP",),
    "application": ("Application", "ApplicationName"),
    "platform": ("Platform",),
    "enforcement_mode": ("EnforcementMode",),
    "rms_encrypted": ("RMSEncrypted", "RmsEncrypted"),
    "previous_file_name": ("PreviousFileName",),
    "target_printer_name": ("TargetPrinterName",),
    "mdatp_device_id": ("MDATPDeviceId",),
    "jit_triggered": ("JitTriggered", "JITTriggered"),
    "evidence_file": ("EvidenceFile",),
    "removable_media_device_attributes": ("RemovableMediaDeviceAttributes",),
    "endpoint_operation": ("EndpointOperation",),
    "authorized_group": ("AuthorizedGroup",),
    "matched_policies": ("MatchedPolicies",),
    "dlp_audit_event_metadata": ("DlpAuditEventMetadata",),
    "session_metadata": ("SessionMetadata",),
    "item_metadata": ("ItemMetadata",),
    "parent_archive_hash": ("ParentArchiveHash",),
    "agent_id": ("AgentId",),
    "agent_name": ("AgentName",),
    "target_agent_id": ("TargetAgentId",),
    "target_agent_name": ("TargetAgentName",),
    "platform_target_agent_id": ("PlatformTargetAgentId",),
    "file_path_url": ("FilePathUrl",),
    "source_location_type": ("SourceLocationType",),
    "destination_location_type": ("DestinationLocationType",),
    "sha1": ("Sha1",),
    "sha256": ("Sha256",),
    "cold_scan_policy_id": ("ColdScanPolicyId",),
    "policy_version": ("PolicyVersion",),
    "associated_admin_units": ("AssociatedAdminUnits",),
    "sensitivity_label_ids_referenced": ("SensitivityLabelIdsReferenced",),
}


def _metric_row() -> dict[str, int]:
    return {"activity_count": 0, "match_count": 0, "risk_weighted_count": 0, "high_confidence_count": 0}


def _int_or_none(value: Any) -> int | None:
    if value is None or value == "":
        return None
    try:
        return int(float(str(value).strip()))
    except (TypeError, ValueError):
        return None


def _raw_first(raw: dict[str, Any], keys: tuple[str, ...]) -> Any:
    for key in keys:
        value = raw.get(key)
        if value is not None and value != "":
            return value
    return None


class StarPipeline:
    def __init__(self, *, input_dir: Path, output_dir: Path, risk: RiskLookup,
                 department_mappings: dict[str, dict[str, Any]],
                 excluded_sit_names: list[str] | None = None,
                 sit_names: dict[str, str] | None = None,
                 archive_raw: bool = True, derive_domains: bool = True,
                 batch_size: int = 50_000, drift_tracker=None) -> None:
        self.input_dir = input_dir
        self.output_dir = output_dir
        self.risk = risk
        self.department_mappings = department_mappings
        self.excluded_norm = {_norm_text(name) for name in (excluded_sit_names or [])}
        self.sit_names = sit_names or {}
        self.archive_raw = archive_raw
        self.derive_domains = derive_domains
        self.batch_size = batch_size
        self.drift_tracker = drift_tracker

        self.registry = IdRegistry()
        self.seen_record_ids: set[str] = set()
        self.seen_sit_keys: set[str] = set()
        # sit_key -> name source ("workbook"/"raw_payload"/"tenant_map"/"unresolved")
        self.sit_name_sources: dict[str, str] = {}
        # keys whose raw/tenant name bridged them onto a workbook row
        self.sit_bridged_keys: set[str] = set()
        # non-workbook keys: sit_key -> (display name, source) for dim_sit
        self.resolved_sit_names: dict[str, tuple[str, str]] = {}
        self.dates: dict[int, Any] = {}
        self.processed_at = datetime.now(timezone.utc).replace(tzinfo=None)
        self.stats: dict[str, int] = defaultdict(int)
        self.aggregates: dict[str, dict[tuple, dict[str, int]]] = {
            "agg_department_sit_day": defaultdict(_metric_row),
            "agg_user_sit_day": defaultdict(_metric_row),
            "agg_location_sit_day": defaultdict(_metric_row),
            "agg_activity_type_sit_day": defaultdict(_metric_row),
            "agg_domain_sit_day": defaultdict(_metric_row),
        }
        self.sinks: dict[str, ParquetSink] = {}

    # ------------------------------------------------------------------ run
    def run(self) -> dict[str, Any]:
        pages = find_ae_pages(self.input_dir)
        if not pages:
            raise FileNotFoundError(
                f"No Activity Explorer page files found under {self.input_dir}"
            )

        self.output_dir.mkdir(parents=True, exist_ok=True)
        sink_names = list(_SINK_TABLES) + (["archive_raw"] if self.archive_raw else [])
        self.sinks = {
            name: ParquetSink(self.output_dir, name, self.batch_size)
            for name in sink_names
        }

        for page_idx, page in enumerate(pages, start=1):
            if page_idx % 25 == 0 or page_idx == len(pages):
                print(f"  Processing page {page_idx:,}/{len(pages):,}: {page.name}")
            records, meta = load_page(page)
            source_file = page.relative_to(self.input_dir).as_posix()
            page_id = self.registry.get_page(source_file, meta, parse_dt)
            for raw in records:
                self._process_record(raw, page_id, source_file)

        counts = {name: sink.close() for name, sink in self.sinks.items()}
        counts.update(write_dimensions(
            self.output_dir, self.registry, self.department_mappings,
            self.dates, self.risk, self.seen_sit_keys,
            resolved_sit_names=self.resolved_sit_names,
        ))
        counts.update(write_aggregates(self.output_dir, self.aggregates))
        source_counts = Counter(self.sit_name_sources.values())
        return {
            "row_counts": counts,
            "raw_records_scanned": self.stats["raw_records"],
            "duplicates_skipped": self.stats["duplicates_skipped"],
            "missing_record_identity": self.stats["missing_record_identity"],
            "sit_rows_before_exclusions": self.stats["sit_rows_total"],
            "excluded_sit_rows": self.stats["sit_rows_excluded"],
            "sit_name_resolution": {
                "observed_sits": len(self.seen_sit_keys),
                "resolved_by": {
                    source: source_counts.get(source, 0)
                    for source in ("workbook", "raw_payload", "tenant_map")
                },
                "unresolved_guids": source_counts.get("unresolved", 0),
                "bridged_to_workbook": len(self.sit_bridged_keys),
            },
        }

    # --------------------------------------------------------------- record
    def _process_record(self, raw: dict[str, Any], page_id: int, source_file: str) -> None:
        self.stats["raw_records"] += 1
        record_identity = (_safe_str(raw.get("RecordIdentity")) or "").strip()
        if not record_identity:
            self.stats["missing_record_identity"] += 1
            return
        if record_identity in self.seen_record_ids:
            self.stats["duplicates_skipped"] += 1
            return
        self.seen_record_ids.add(record_identity)
        activity_id = stable_int_id("activity", record_identity)

        happened_at = parse_dt(raw.get("Happened") or raw.get("Timestamp"))
        date_key = date_key_of(happened_at)
        if date_key and happened_at:
            self.dates[date_key] = happened_at.date()

        policies = [
            row for row in as_list(raw.get("PolicyMatchInfo"))
            if row.get("PolicyId") or row.get("RuleId") or row.get("PolicyName") or row.get("RuleName")
        ]
        # Empty dicts ({}) are the API's "no email payload" — only truthy
        # payloads count as email (matches the legacy email_details row set).
        email_rows = [row for row in as_list(raw.get("EmailInfo")) if row]
        email = email_rows[0] if email_rows else None
        has_policy = bool(policies)
        has_email = email is not None
        activity = _safe_str(raw.get("Activity"))
        workload = _safe_str(raw.get("Workload") or raw.get("Location"))
        group = activity_group_of(activity, workload, has_policy, has_email)

        upn = _safe_str(raw.get("User") or raw.get("UserKey"))
        user_id, department_id = self.registry.get_user(upn, self.department_mappings)
        user_domain = extract_domain_from_email((upn or "").strip().lower())

        registry = self.registry
        workload_id = registry.get_workload(workload)
        activity_type_id = registry.get_activity_type(activity, group)
        file_path = _safe_str(raw.get("FilePath"))
        source_location_id = registry.get_location(extract_folder(file_path))
        target_location_id = registry.get_location(extract_folder(_safe_str(raw.get("TargetFilePath"))))
        item_name = _safe_str(_raw_first(raw, ("ItemName", "FileName")))
        file_id = registry.get_file(
            file_path, item_name, _safe_str(raw.get("FileType")), _safe_str(raw.get("FileExtension"))
        )
        policy_rule_id = registry.get_policy(policies[0]) if policies else None
        app_identity_id = registry.get_app_identity(
            *(raw.get(key) for key in APP_IDENTITY_KEYS)
        )

        if self.derive_domains:
            target_domain = derive_target_domain(raw, email, user_domain)
            originating_domain = _safe_str(raw.get("OriginatingDomain"))
            if not originating_domain or originating_domain == "None" or not originating_domain.strip():
                originating_domain = user_domain
        else:
            target_domain = _safe_str(raw.get("TargetDomain"))
            originating_domain = _safe_str(raw.get("OriginatingDomain"))
        target_domain_id = registry.get_domain(target_domain)
        originating_domain_id = registry.get_domain(originating_domain)

        sit_summary = self._process_sits(
            raw, activity_id=activity_id, date_key=date_key, user_id=user_id,
            department_id=department_id, activity_type_id=activity_type_id,
            workload_id=workload_id, source_location_id=source_location_id,
            target_location_id=target_location_id, file_id=file_id,
            target_domain_id=target_domain_id, policy_rule_id=policy_rule_id,
        )

        self.sinks["fact_activity"].write({
            "activity_id": activity_id,
            "date_key": date_key,
            "happened_at": happened_at,
            "user_id": user_id,
            "department_id": department_id,
            "activity_type_id": activity_type_id,
            "workload_id": workload_id,
            "source_location_id": source_location_id,
            "target_location_id": target_location_id,
            "file_id": file_id,
            "policy_rule_id": policy_rule_id,
            "target_domain_id": target_domain_id,
            "originating_domain_id": originating_domain_id,
            "app_identity_id": app_identity_id,
            "user_type": _safe_str(raw.get("UserType")),
            "data_platform": _safe_str(raw.get("DataPlatform")),
            "file_size_bytes": _int_or_none(raw.get("FileSize")),
            "has_sit": sit_summary["count"] > 0,
            "has_policy": has_policy,
            "has_email": has_email,
            "sit_type_count": sit_summary["count"],
            "activity_risk_score": sit_summary["risk"],
            "max_sit_risk_score": sit_summary["max_risk"],
        })

        for policy in policies:
            rule_id = registry.get_policy(policy)
            if rule_id is None:
                continue
            self.sinks["fact_policy_activity"].write({
                "activity_id": activity_id, "date_key": date_key,
                "user_id": user_id, "department_id": department_id,
                "policy_rule_id": rule_id, "activity_type_id": activity_type_id,
                "workload_id": workload_id,
            })

        if email is not None:
            self._process_email(raw, email, activity_id, date_key)

        if any(raw.get(key) is not None for key in COPILOT_KEYS):
            self.sinks["fact_copilot_interaction"].write({
                "activity_id": activity_id,
                "date_key": date_key,
                "user_id": user_id,
                "app_identity_id": app_identity_id,
                "purview_ai_app_location": _safe_str(raw.get("PurviewAIAppLocation")),
                "enriched_copilot_thread_or_correlation_id": _safe_str(
                    raw.get("EnrichedCopilotThreadOrCorrelationId")
                ),
                "enriched_llm_message_ids": json_str(raw.get("EnrichedLLMMessageIds")),
                "has_web_search_query": as_bool(_raw_first(raw, ("HasWebsearchQuery", "HasWebSearchQuery"))),
                "are_files_referenced": as_bool(raw.get("AreFilesReferenced")),
                "are_sensitive_files_referenced": as_bool(raw.get("AreSensitiveFilesReferenced")),
                "sensitivity_label_ids_referenced": json_str(raw.get("SensitivityLabelIdsReferenced")),
                "copilot_event_data_json": json_str(raw.get("CopilotEventData")),
                "accessed_resources_json": json_str(raw.get("AccessedResources")),
            })

        extras = {key: value for key, value in raw.items() if key not in CONSUMED_KEYS}
        if extras and self.drift_tracker is not None:
            self.drift_tracker.record("fact_activity_detail", extras)
        detail = {
            "activity_id": activity_id,
            "item_name": item_name,
            "source_file": source_file,
            "extra_json": json.dumps(extras, ensure_ascii=False, default=str) if extras else None,
        }
        for column, keys in _DETAIL_CONTRACT.items():
            detail[column] = json_str(_raw_first(raw, keys))
        self.sinks["fact_activity_detail"].write(detail)

        self.sinks["activity_record_index"].write({
            "record_identity": record_identity,
            "activity_id": activity_id,
            "page_id": page_id,
            "source_export": str(self.input_dir),
            "processed_at_utc": self.processed_at,
        })

        if self.archive_raw:
            self.sinks["archive_raw"].write({
                "activity_id": activity_id,
                "record_identity": record_identity,
                "original_activity_id": json_str(raw.get("ActivityId")),
                "sensitive_info_type_data": json_str(raw.get("SensitiveInfoTypeData")),
                "sensitive_info_type_buckets_data": json_str(raw.get("SensitiveInfoTypeBucketsData")),
                "policy_match_info": json_str(raw.get("PolicyMatchInfo")),
                "email_info": json_str(raw.get("EmailInfo")),
                "sha1": json_str(raw.get("Sha1")),
                "sha256": json_str(raw.get("Sha256")),
            })

    # ----------------------------------------------------------------- SITs
    def _process_sits(self, raw: dict[str, Any], **ids) -> dict[str, int]:
        buckets: dict[str, dict[str, Any]] = {}
        for row in as_list(raw.get("SensitiveInfoTypeBucketsData")):
            bucket_id = _safe_str(row.get("Id"))
            if bucket_id:
                buckets[bucket_id.lower()] = row

        sit_rows = as_list(raw.get("SensitiveInfoTypeData"))
        total_risk = 0
        max_risk = 0
        for sit in sit_rows:
            sit_id = _safe_str(sit.get("SensitiveInfoTypeId"))
            raw_name = _safe_str(_raw_first(
                sit, ("SensitiveInfoTypeName", "SitName", "DisplayName", "Name")))
            sit_key, risk_row, name_source, display_name, bridged = resolve_detected_sit(
                sit_id, raw_name, self.risk, self.sit_names)
            self.seen_sit_keys.add(sit_key)
            source = name_source or "unresolved"
            previous = self.sit_name_sources.get(sit_key)
            if previous is None or (previous == "unresolved" and source != "unresolved"):
                self.sit_name_sources[sit_key] = source
            if bridged:
                self.sit_bridged_keys.add(sit_key)
            elif display_name and risk_row is None:
                self.resolved_sit_names.setdefault(sit_key, (display_name, source))
            risk_score = (risk_row or {}).get("risk_score") or 0
            match_count = _int_or_none(sit.get("Count")) or 0
            confidence = _int_or_none(sit.get("Confidence"))
            bucket = buckets.get((sit_id or "").lower(), {})
            bucket_high = _int_or_none(bucket.get("High")) or 0
            high_confidence = bucket_high or (match_count if (confidence or 0) >= 85 else 0)
            risk_weighted = risk_score * match_count
            total_risk += risk_weighted
            max_risk = max(max_risk, risk_score)
            self.stats["sit_rows_total"] += 1

            # Exclusion keys on the RESOLVED display name: workbook rows match
            # exactly as before; raw-payload/tenant-map names are newly
            # excludable; unresolved GUIDs never match the name-based list.
            if display_name and _norm_text(display_name) in self.excluded_norm:
                self.stats["sit_rows_excluded"] += 1
                continue

            self.sinks["fact_activity_sit"].write({
                **ids,
                "sit_key": sit_key,
                "classifier_type": _safe_str(sit.get("ClassifierType") or bucket.get("ClassifierType")),
                "match_count": match_count,
                "unique_count": _int_or_none(sit.get("UniqueCount")),
                "confidence": confidence,
                "bucket_low": _int_or_none(bucket.get("Low")) or 0,
                "bucket_medium": _int_or_none(bucket.get("Medium")) or 0,
                "bucket_high": bucket_high,
                "risk_score": risk_score,
                "risk_weighted_count": risk_weighted,
                "high_confidence_count": high_confidence,
            })

            metrics = (match_count, risk_weighted, high_confidence)
            date_key = ids["date_key"]
            workload_id = ids["workload_id"]
            activity_type_id = ids["activity_type_id"]
            agg = self.aggregates
            self._add_metric(agg["agg_department_sit_day"],
                             (date_key, ids["department_id"], sit_key, workload_id, activity_type_id), metrics)
            if ids["user_id"] is not None:
                self._add_metric(agg["agg_user_sit_day"],
                                 (date_key, ids["user_id"], ids["department_id"], sit_key, workload_id, activity_type_id), metrics)
            if ids["source_location_id"] is not None:
                self._add_metric(agg["agg_location_sit_day"],
                                 (date_key, ids["source_location_id"], sit_key, workload_id, activity_type_id), metrics)
            self._add_metric(agg["agg_activity_type_sit_day"],
                             (date_key, activity_type_id, sit_key, workload_id), metrics)
            if ids["target_domain_id"] is not None:
                self._add_metric(agg["agg_domain_sit_day"],
                                 (date_key, ids["target_domain_id"], sit_key, workload_id, activity_type_id), metrics)

        return {"count": len(sit_rows), "risk": total_risk, "max_risk": max_risk}

    @staticmethod
    def _add_metric(agg: dict[tuple, dict[str, int]], key: tuple,
                    metrics: tuple[int, int, int]) -> None:
        row = agg[key]
        row["activity_count"] += 1
        row["match_count"] += metrics[0]
        row["risk_weighted_count"] += metrics[1]
        row["high_confidence_count"] += metrics[2]

    # ---------------------------------------------------------------- email
    def _process_email(self, raw: dict[str, Any], email: dict[str, Any],
                       activity_id: int, date_key: int | None) -> None:
        registry = self.registry
        sender_id = registry.get_email(_safe_str(email.get("Sender")))
        for receiver in split_receivers(email.get("Receivers")):
            recipient_id = registry.get_email(receiver)
            recipient_row = registry.email_rows.get(recipient_id or -1, {})
            self.sinks["fact_email_recipient"].write({
                "activity_id": activity_id,
                "date_key": date_key,
                "sender_email_address_id": sender_id,
                "recipient_email_address_id": recipient_id,
                "recipient_domain_id": recipient_row.get("domain_id"),
            })
        attachments = as_list(raw.get("AttachmentDetails"))
        self.sinks["fact_email_detail"].write({
            "activity_id": activity_id,
            "date_key": date_key,
            "sender_email_address_id": sender_id,
            "subject": _safe_str(email.get("Subject")),
            "message_id": _safe_str(_raw_first(email, ("MessageID", "MessageId"))),
            "attachment_count": len(attachments),
            "attachment_details_json": json_str(raw.get("AttachmentDetails")),
        })
