"""Fact, aggregate, index, and pipeline-only table declarations (v6)."""

from __future__ import annotations

from .spec_types import TableSpec, _c, _fk, _key, _metric

FACT_ACTIVITY = TableSpec(
    name="fact_activity",
    kind="fact",
    key="activity_id",
    description="One row per unique Activity Explorer record (RecordIdentity).",
    columns=(
        _key("activity_id", desc="Stable surrogate from RecordIdentity."),
        _fk("date_key", "FK to dim_date (yyyymmdd)."),
        _c("happened_at", "timestamp_us", "Activity timestamp (UTC).", fmt="yyyy-MM-dd HH:mm:ss"),
        _fk("user_id", "FK to dim_user."),
        _fk("department_id", "FK to dim_department."),
        _fk("activity_type_id", "FK to dim_activity_type."),
        _fk("workload_id", "FK to dim_workload."),
        _fk("source_location_id", "FK to dim_location (FilePath folder)."),
        _fk("target_location_id", "FK to dim_location (TargetFilePath folder; inactive relationship)."),
        _fk("file_id", "FK to dim_file."),
        _fk("policy_rule_id", "FK to dim_policy (first matched policy)."),
        _fk("target_domain_id", "FK to dim_domain (explicit or derived target)."),
        _fk("originating_domain_id", "FK to dim_domain (inactive relationship)."),
        _fk("app_identity_id", "FK to dim_app_identity."),
        _c("user_type", "string", "Raw UserType (Regular/Admin/System/...)."),
        _c("data_platform", "string", "Raw DataPlatform."),
        _metric("file_size_bytes", "FileSize in bytes."),
        _c("has_sit", "bool", "Record carried SIT detections."),
        _c("has_policy", "bool", "Record carried DLP policy matches."),
        _c("has_email", "bool", "Record carried EmailInfo."),
        _metric("sit_type_count", "Distinct SIT detections on the record."),
        _metric("activity_risk_score", "Sum of risk_score*match_count over all SIT detections."),
        _c("max_sit_risk_score", "int64", "Highest single SIT risk rating on the record.", fmt="#,0", agg="max"),
    ),
)

FACT_ACTIVITY_SIT = TableSpec(
    name="fact_activity_sit",
    kind="fact",
    description="One row per (activity, SIT detection). Excluded SIT names are filtered at ETL.",
    columns=(
        _fk("activity_id", "Activity surrogate key."),
        _fk("date_key", "FK to dim_date."),
        _fk("user_id", "FK to dim_user."),
        _fk("department_id", "FK to dim_department."),
        _fk("activity_type_id", "FK to dim_activity_type."),
        _fk("workload_id", "FK to dim_workload."),
        _fk("source_location_id", "FK to dim_location."),
        _fk("target_location_id", "FK to dim_location (inactive relationship)."),
        _fk("file_id", "FK to dim_file."),
        _fk("target_domain_id", "FK to dim_domain (enables domain x SIT without fact hops)."),
        _fk("policy_rule_id", "FK to dim_policy (enables risk-by-rule without fact hops)."),
        _c("sit_key", "string", "FK to dim_sit.", nullable=False),
        _c("classifier_type", "string", "ClassifierType from the detection."),
        _metric("match_count", "Detection count."),
        _metric("unique_count", "Unique match count."),
        _c("confidence", "int64", "Detection confidence."),
        _metric("bucket_low", "Low-confidence bucket count."),
        _metric("bucket_medium", "Medium-confidence bucket count."),
        _metric("bucket_high", "High-confidence bucket count."),
        _c("risk_score", "int64", "Risk rating of the SIT at conversion time."),
        _metric("risk_weighted_count", "risk_score * match_count."),
        _metric("high_confidence_count", "High bucket, else match_count when confidence >= 85."),
    ),
)

FACT_POLICY_ACTIVITY = TableSpec(
    name="fact_policy_activity",
    kind="fact",
    description="One row per (activity, matched policy rule).",
    columns=(
        _fk("activity_id", "Activity surrogate key."),
        _fk("date_key", "FK to dim_date."),
        _fk("user_id", "FK to dim_user."),
        _fk("department_id", "FK to dim_department."),
        _fk("policy_rule_id", "FK to dim_policy."),
        _fk("activity_type_id", "FK to dim_activity_type."),
        _fk("workload_id", "FK to dim_workload."),
    ),
)

FACT_EMAIL_RECIPIENT = TableSpec(
    name="fact_email_recipient",
    kind="fact",
    description="One row per (activity, email recipient).",
    columns=(
        _fk("activity_id", "Activity surrogate key."),
        _fk("date_key", "FK to dim_date."),
        _fk("sender_email_address_id", "FK to dim_email_address (inactive relationship)."),
        _fk("recipient_email_address_id", "FK to dim_email_address."),
        _fk("recipient_domain_id", "FK to dim_domain."),
    ),
)

FACT_EMAIL_DETAIL = TableSpec(
    name="fact_email_detail",
    kind="fact",
    key="activity_id",
    description="One row per email-bearing activity: subject/message id/attachments.",
    columns=(
        _key("activity_id", desc="Activity surrogate key."),
        _fk("date_key", "FK to dim_date (inactive relationship; the active inbound path is the fact_activity_sit rollup)."),
        _fk("sender_email_address_id", "FK to dim_email_address."),
        _c("subject", "string", "Email subject."),
        _c("message_id", "string", "Internet message id."),
        _metric("attachment_count", "Number of attachments."),
        _c("attachment_details_json", "string", "Raw AttachmentDetails JSON."),
    ),
)

FACT_COPILOT_INTERACTION = TableSpec(
    name="fact_copilot_interaction",
    kind="fact",
    key="activity_id",
    description="One row per Copilot/AI-enriched activity.",
    columns=(
        _key("activity_id", desc="Activity surrogate key."),
        _fk("date_key", "FK to dim_date."),
        _fk("user_id", "FK to dim_user."),
        _fk("app_identity_id", "FK to dim_app_identity."),
        _c("purview_ai_app_location", "string", "PurviewAIAppLocation (host app)."),
        _c("enriched_copilot_thread_or_correlation_id", "string", "Thread/correlation id."),
        _c("enriched_llm_message_ids", "string", "LLM message ids (JSON list)."),
        _c("has_web_search_query", "bool", "HasWebsearchQuery."),
        _c("are_files_referenced", "bool", "AreFilesReferenced."),
        _c("are_sensitive_files_referenced", "bool", "AreSensitiveFilesReferenced."),
        _c("sensitivity_label_ids_referenced", "string", "SensitivityLabelIdsReferenced (JSON list)."),
        _c("copilot_event_data_json", "string", "Raw CopilotEventData JSON."),
        _c("accessed_resources_json", "string", "Raw AccessedResources JSON."),
    ),
)

FACT_ACTIVITY_DETAIL = TableSpec(
    name="fact_activity_detail",
    kind="fact",
    key="activity_id",
    description=(
        "1:1 drillthrough detail per activity. record_identity lives in "
        "activity_record_index; file_path lives in dim_file. Unmapped raw keys "
        "land in extra_json (and the schema-drift report)."
    ),
    columns=(
        _key("activity_id", desc="Activity surrogate key."),
        _c("item_name", "string", "ItemName/FileName."),
        _c("target_file_path", "string", "TargetFilePath."),
        _c("target_url", "string", "TargetUrl."),
        _c("device_name", "string", "DeviceName."),
        _c("client_ip", "string", "ClientIP."),
        _c("application", "string", "Application/ApplicationName."),
        _c("platform", "string", "Platform."),
        _c("source_file", "string", "Source page file (relative path)."),
        # Typed endpoint/DLP raw-contract columns (from the fork's
        # _REPORT_ACTIVITY_COLUMNS). Raw values are preserved as strings since
        # upstream emits mixed primitives ("True", JSON blobs, GUIDs, ...).
        _c("enforcement_mode", "string", "EnforcementMode."),
        _c("rms_encrypted", "string", "RMSEncrypted."),
        _c("previous_file_name", "string", "PreviousFileName."),
        _c("target_printer_name", "string", "TargetPrinterName."),
        _c("mdatp_device_id", "string", "MDATPDeviceId."),
        _c("jit_triggered", "string", "JitTriggered."),
        _c("evidence_file", "string", "EvidenceFile (JSON)."),
        _c("removable_media_device_attributes", "string", "RemovableMediaDeviceAttributes (JSON)."),
        _c("endpoint_operation", "string", "EndpointOperation."),
        _c("authorized_group", "string", "AuthorizedGroup."),
        _c("matched_policies", "string", "MatchedPolicies (JSON)."),
        _c("dlp_audit_event_metadata", "string", "DlpAuditEventMetadata (JSON)."),
        _c("session_metadata", "string", "SessionMetadata (JSON)."),
        _c("item_metadata", "string", "ItemMetadata (JSON)."),
        _c("parent_archive_hash", "string", "ParentArchiveHash."),
        _c("agent_id", "string", "AgentId."),
        _c("agent_name", "string", "AgentName."),
        _c("target_agent_id", "string", "TargetAgentId."),
        _c("target_agent_name", "string", "TargetAgentName."),
        _c("platform_target_agent_id", "string", "PlatformTargetAgentId."),
        _c("file_path_url", "string", "FilePathUrl."),
        _c("source_location_type", "string", "SourceLocationType."),
        _c("destination_location_type", "string", "DestinationLocationType."),
        _c("sha1", "string", "File SHA1."),
        _c("sha256", "string", "File SHA256."),
        _c("cold_scan_policy_id", "string", "ColdScanPolicyId."),
        _c("policy_version", "string", "PolicyVersion."),
        _c("associated_admin_units", "string", "AssociatedAdminUnits (JSON)."),
        _c("sensitivity_label_ids_referenced", "string", "SensitivityLabelIdsReferenced (JSON)."),
        _c("extra_json", "string", "Catch-all: raw keys not consumed by any typed column."),
    ),
)

_AGG_METRICS = (
    _metric("activity_count", "Distinct activities contributing."),
    _metric("match_count", "Total SIT match count."),
    _metric("risk_weighted_count", "Total risk_score * match_count."),
    _metric("high_confidence_count", "Total high-confidence matches."),
)

AGG_DEPARTMENT_SIT_DAY = TableSpec(
    name="agg_department_sit_day",
    kind="agg",
    columns=(
        _fk("date_key"), _fk("department_id"),
        _c("sit_key", "string", "FK to dim_sit.", nullable=False),
        _fk("workload_id"), _fk("activity_type_id"),
        *_AGG_METRICS,
    ),
)

AGG_USER_SIT_DAY = TableSpec(
    name="agg_user_sit_day",
    kind="agg",
    columns=(
        _fk("date_key"), _fk("user_id"), _fk("department_id"),
        _c("sit_key", "string", "FK to dim_sit.", nullable=False),
        _fk("workload_id"), _fk("activity_type_id"),
        *_AGG_METRICS,
    ),
)

AGG_LOCATION_SIT_DAY = TableSpec(
    name="agg_location_sit_day",
    kind="agg",
    columns=(
        _fk("date_key"), _fk("source_location_id"),
        _c("sit_key", "string", "FK to dim_sit.", nullable=False),
        _fk("workload_id"), _fk("activity_type_id"),
        *_AGG_METRICS,
    ),
)

AGG_ACTIVITY_TYPE_SIT_DAY = TableSpec(
    name="agg_activity_type_sit_day",
    kind="agg",
    columns=(
        _fk("date_key"), _fk("activity_type_id"),
        _c("sit_key", "string", "FK to dim_sit.", nullable=False),
        _fk("workload_id"),
        *_AGG_METRICS,
    ),
)

AGG_DOMAIN_SIT_DAY = TableSpec(
    name="agg_domain_sit_day",
    kind="agg",
    description="Target-domain x SIT daily rollup (domain flow pages).",
    columns=(
        _fk("date_key"), _fk("target_domain_id"),
        _c("sit_key", "string", "FK to dim_sit.", nullable=False),
        _fk("workload_id"), _fk("activity_type_id"),
        *_AGG_METRICS,
    ),
)

ACTIVITY_RECORD_INDEX = TableSpec(
    name="activity_record_index",
    kind="index",
    key="record_identity",
    description="RecordIdentity -> activity_id map with page provenance.",
    columns=(
        _key("record_identity", "string", "Raw RecordIdentity GUID."),
        _c("activity_id", "int64", "Activity surrogate key.", nullable=False),
        _fk("page_id", "FK to dim_source_page."),
        _c("source_export", "string", "Export root the record came from."),
        _c("processed_at_utc", "timestamp_us", "Conversion timestamp.", fmt="yyyy-MM-dd HH:mm:ss"),
    ),
)

ARCHIVE_RAW = TableSpec(
    name="archive_raw",
    kind="pipeline_only",
    key="activity_id",
    description="Raw nested payloads per activity (future-proofing; not loaded into PBI).",
    columns=(
        _key("activity_id", desc="Activity surrogate key."),
        _c("record_identity", "string", "Raw RecordIdentity GUID."),
        _c("original_activity_id", "string", "Raw ActivityId field from the record."),
        _c("sensitive_info_type_data", "string", "Raw SensitiveInfoTypeData JSON."),
        _c("sensitive_info_type_buckets_data", "string", "Raw SensitiveInfoTypeBucketsData JSON."),
        _c("policy_match_info", "string", "Raw PolicyMatchInfo JSON."),
        _c("email_info", "string", "Raw EmailInfo JSON."),
        _c("sha1", "string", "Raw Sha1."),
        _c("sha256", "string", "Raw Sha256."),
    ),
)

FACT_TABLES = (
    FACT_ACTIVITY, FACT_ACTIVITY_SIT, FACT_POLICY_ACTIVITY,
    FACT_EMAIL_RECIPIENT, FACT_EMAIL_DETAIL, FACT_COPILOT_INTERACTION,
    FACT_ACTIVITY_DETAIL,
)

AGG_TABLES = (
    AGG_DEPARTMENT_SIT_DAY, AGG_USER_SIT_DAY, AGG_LOCATION_SIT_DAY,
    AGG_ACTIVITY_TYPE_SIT_DAY, AGG_DOMAIN_SIT_DAY,
)
