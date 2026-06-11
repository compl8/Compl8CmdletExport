"""Shared field references and layout helpers for the Activity Explorer report.

Field constants bind report visuals to the star-schema v6 column names
(parquet_builder.star.dimensions / facts) and to the measure names declared
in ae_measures. Legacy-report bindings translate as:

    Activities[User/Department]    -> dim_user.user_upn / dim_department.department
    Activities[Activity/Workload]  -> dim_activity_type.activity / dim_workload.workload
    Activities[HappenedDate]       -> dim_date.date
    Activities[Happened]           -> fact_activity.happened_at
    Activities[FileName/FileType]  -> dim_file.file_name / dim_file.file_type
    Activities[TargetDomain]       -> dim_domain.domain (active target_domain rel)
    Activities[DeviceName/...]     -> fact_activity_detail.*
    Locations[FolderPath]          -> dim_location.folder_path
    Domains[Domain]                -> dim_domain.domain
    Policies[PolicyName/RuleName]  -> dim_policy.policy_name / rule_name
    Email_Details[Subject]         -> fact_email_detail.subject
    SIT_Reference[*]               -> dim_sit.* (sit_name, qgiscf_dlm, category, ...)
"""

from __future__ import annotations

from .expressions import Field, col, meas
from .report_layout import SLICER_HEIGHT, SLICER_ROW_Y, grid_row
from .visual_factories import VisualSpec, slicer, table

# --- dimension columns -------------------------------------------------------

DATE = col("dim_date", "date", "Date")
MONTH_SHORT = col("dim_date", "month_short", "Month")

DEPARTMENT = col("dim_department", "department", "Department")
USER = col("dim_user", "user_upn", "User")

SIT_NAME = col("dim_sit", "sit_name", "SIT Name")
QGISCF_DLM = col("dim_sit", "qgiscf_dlm", "QGISCF DLM")
SIT_CATEGORY = col("dim_sit", "category", "SIT Category")
SIT_SOURCE = col("dim_sit", "source", "SIT Source")
RISK_BAND = col("dim_sit", "risk_band", "Risk Band")
PSPF_CLASSIFICATION = col("dim_sit", "pspf_classification", "PSPF Classification")

ACTIVITY = col("dim_activity_type", "activity", "Activity")
ACTIVITY_GROUP = col("dim_activity_type", "activity_group", "Activity Group")
WORKLOAD = col("dim_workload", "workload", "Workload")

FOLDER_PATH = col("dim_location", "folder_path", "Folder Path")
FOLDER_NAME = col("dim_location", "folder_name", "Folder")
PATH_DEPTH = col("dim_location", "path_depth", "Path Depth")

DOMAIN = col("dim_domain", "domain", "Domain")
PARENT_DOMAIN = col("dim_domain", "parent_domain", "Parent Domain")

POLICY_NAME = col("dim_policy", "policy_name", "Policy")
POLICY_MODE = col("dim_policy", "policy_mode", "Policy Mode")
RULE_NAME = col("dim_policy", "rule_name", "Rule")

FILE_NAME = col("dim_file", "file_name", "File Name")
FILE_TYPE = col("dim_file", "file_type", "File Type")

APP_IDENTITY = col("dim_app_identity", "app_identity", "App Identity")
APP_IDENTITY_CATEGORY = col("dim_app_identity", "app_identity_category", "App Category")
PURVIEW_AI_APP_NAME = col("dim_app_identity", "purview_ai_app_name", "AI App")

# --- fact columns ------------------------------------------------------------

HAPPENED_AT = col("fact_activity", "happened_at", "Happened")
FILE_SIZE_BYTES = col("fact_activity", "file_size_bytes", "File Size (Bytes)")
USER_TYPE = col("fact_activity", "user_type", "User Type")

ITEM_NAME = col("fact_activity_detail", "item_name", "Item Name")
TARGET_URL = col("fact_activity_detail", "target_url", "Target URL")
DEVICE_NAME = col("fact_activity_detail", "device_name", "Device")
APPLICATION = col("fact_activity_detail", "application", "Application")
PLATFORM = col("fact_activity_detail", "platform", "Platform")
SOURCE_LOCATION_TYPE = col("fact_activity_detail", "source_location_type", "Source Location Type")
DESTINATION_LOCATION_TYPE = col(
    "fact_activity_detail", "destination_location_type", "Destination Location Type")
AGENT_NAME = col("fact_activity_detail", "agent_name", "Agent")
TARGET_AGENT_NAME = col("fact_activity_detail", "target_agent_name", "Target Agent")
SOURCE_FILE = col("fact_activity_detail", "source_file", "Source Page File")

EMAIL_SUBJECT = col("fact_email_detail", "subject", "Subject")
ATTACHMENT_COUNT = col("fact_email_detail", "attachment_count", "Attachments")

AI_APP_LOCATION = col("fact_copilot_interaction", "purview_ai_app_location", "AI App Host")

# --- measure references (homes match ae_measures declarations) ---------------

RAW_ACTIVITIES = meas("fact_activity", "Raw Activities")
TOTAL_ACTIVITIES = meas("fact_activity", "Total Activities")
UNIQUE_USERS = meas("fact_activity", "Unique Users")
UNIQUE_FILES = meas("fact_activity", "Unique Files")
DLP_RULE_MATCHES = meas("fact_activity", "DLP Rule Matches")
ACTIVITIES_WITH_SIT_DATA = meas("fact_activity", "Activities with SIT Data")
TOTAL_FILE_SIZE_GB = meas("fact_activity", "Total File Size (GB)")
EMAIL_ACTIVITIES = meas("fact_activity", "Email Activities")
ACTIVITY_RISK_SCORE = meas("fact_activity", "Activity Risk Score")
HIGH_RISK_ACTIVITIES = meas("fact_activity", "High Risk Activities")
TOTAL_RISK = meas("fact_activity", "TotalRisk")
POLICY_MATCH_COUNT = meas("fact_activity", "Policy Match Count")
UNIQUE_POLICIES_TRIGGERED = meas("fact_activity", "Unique Policies Triggered")
UNIQUE_RULES_TRIGGERED = meas("fact_activity", "Unique Rules Triggered")
TARGET_LOCATION_ACTIVITIES = meas("fact_activity", "Target Location Activities")
EXTERNAL_DOMAIN_ACTIVITIES = meas("fact_activity", "External Domain Activities")
UNIQUE_TARGET_DOMAINS = meas("fact_activity", "Unique Target Domains")
DAILY_ACTIVITY_AVERAGE = meas("fact_activity", "Daily Activity Average")

ACTIVITIES_BY_SIT = meas("fact_activity_sit", "Activities by SIT")
TOTAL_SIT_DETECTIONS = meas("fact_activity_sit", "Total SIT Detections")
TOTAL_SIT_INSTANCE_COUNT = meas("fact_activity_sit", "Total SIT Instance Count")
AVG_CONFIDENCE = meas("fact_activity_sit", "Avg Confidence")
HIGH_CONFIDENCE_DETECTIONS = meas("fact_activity_sit", "High Confidence Detections")
UNIQUE_SIT_TYPES = meas("fact_activity_sit", "Unique SIT Types Detected")
TOTAL_SIT_RISK = meas("fact_activity_sit", "Total SIT Risk")
WEIGHTED_RISK_SCORE = meas("fact_activity_sit", "Weighted Risk Score")
AVG_WEIGHTED_RISK = meas("fact_activity_sit", "Avg Weighted Risk")
HIGH_RISK_DETECTIONS = meas("fact_activity_sit", "High Risk Detections")
PROTECTED_CLASSIFICATION_COUNT = meas("fact_activity_sit", "Protected Classification Count")
HIGH_CONFIDENCE_PCT = meas("fact_activity_sit", "High Confidence %")
CRITICAL_SIT_EVENTS = meas("fact_activity_sit", "Critical SIT Events")

AVG_RISK_RATING = meas("dim_sit", "Avg Risk Rating")
MAX_RISK_RATING = meas("dim_sit", "Max Risk Rating")

TOTAL_EMAIL_RECIPIENTS = meas("fact_email_recipient", "Total Email Recipients")
EXTERNAL_EMAIL_RECIPIENTS = meas("fact_email_recipient", "External Email Recipients")
UNIQUE_RECEIVER_DOMAINS = meas("fact_email_recipient", "Unique Receiver Domains")

COPILOT_INTERACTIONS = meas("fact_copilot_interaction", "Copilot Interactions")
COPILOT_FILE_REFERENCES = meas("fact_copilot_interaction", "Copilot File References")
COPILOT_SENSITIVE_FILE_REFERENCES = meas(
    "fact_copilot_interaction", "Copilot Sensitive File References")

DEPT_SIT_EVENTS = meas("agg_department_sit_day", "Department SIT Activity Events")
DEPT_SIT_MATCHES = meas("agg_department_sit_day", "Department SIT Matches")
DEPT_RISK_PRESSURE = meas("agg_department_sit_day", "Department Risk Pressure")
DEPT_HIGH_CONFIDENCE_PCT = meas("agg_department_sit_day", "Department High Confidence %")
DEPT_HIGH_CRITICAL_EVENTS = meas("agg_department_sit_day", "Department High Critical Events")
DEPT_AVG_RISK_PER_MATCH = meas("agg_department_sit_day", "Department Avg Risk Per Match")

ACTIVITY_TYPE_RISK_PRESSURE = meas("agg_activity_type_sit_day", "Activity Type Risk Pressure")
LOCATION_SIT_MATCHES = meas("agg_location_sit_day", "Location SIT Matches")
LOCATION_RISK_PRESSURE = meas("agg_location_sit_day", "Location Risk Pressure")
LOCATION_HIGH_CONFIDENCE = meas("agg_location_sit_day", "Location High Confidence Matches")
USER_RISK_PRESSURE = meas("agg_user_sit_day", "User Risk Pressure")
USER_SIT_MATCHES = meas("agg_user_sit_day", "User SIT Matches")

# --- layout helpers ----------------------------------------------------------

STANDARD_SLICERS: tuple[Field, ...] = (DATE, WORKLOAD, DEPARTMENT, QGISCF_DLM)


def slicer_band(prefix: str, fields: tuple[Field, ...] = STANDARD_SLICERS, *,
                y: float = SLICER_ROW_Y, height: float = SLICER_HEIGHT) -> list[VisualSpec]:
    """A row of titled slicers (analysis-page standard: date/workload/dept/DLM)."""
    cells = grid_row(len(fields), y, height)
    return [
        slicer(f"{prefix}-slicer-{field.name}", field, cell, title=field.shown_as())
        for field, cell in zip(fields, cells)
    ]


def by_sit_table(seed: str, field: Field, rect, *, title: str) -> VisualSpec:
    """Legacy staple: a two-column table of <category> x [Activities by SIT]."""
    return table(seed, [field, ACTIVITIES_BY_SIT], rect, title=title,
                 order_by=ACTIVITIES_BY_SIT)
