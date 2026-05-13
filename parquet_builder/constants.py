"""Column rename maps and detection patterns shared across the pipeline."""

from __future__ import annotations

import re

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
