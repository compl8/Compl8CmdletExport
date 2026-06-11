"""Dimension table declarations for the star schema SSOT (v6)."""

from __future__ import annotations

from .spec_types import TableSpec, _c, _fk, _key

DIM_DATE = TableSpec(
    name="dim_date",
    kind="dim",
    key="date_key",
    description="Continuous calendar covering min..max activity dates (no holes).",
    columns=(
        _key("date_key", desc="Date as yyyymmdd integer."),
        _c("date", "date32", "Calendar date (mark as PBI date table key).", nullable=False, fmt="yyyy-MM-dd"),
        _c("year", "int64", "Calendar year."),
        _c("month", "int64", "Month number 1-12."),
        _c("month_name", "string", "Full month name."),
        _c("month_short", "string", "Abbreviated month name (Jan..Dec)."),
        _c("quarter", "int64", "Calendar quarter 1-4."),
        _c("week_of_year", "int64", "ISO week of year."),
        _c("day_of_week", "string", "Full weekday name."),
        _c("day_of_week_num", "int64", "Weekday number, Monday=0."),
        _c("is_weekend", "bool", "True for Saturday/Sunday."),
    ),
)

DIM_DEPARTMENT = TableSpec(
    name="dim_department",
    kind="dim",
    key="department_id",
    description="Departments from the GAL/department mapping; 'Unmapped' row for unknowns.",
    columns=(
        _key("department_id"),
        _c("department", "string", "Department name ('Unmapped' when no mapping)."),
        _c("division", "string", "Division/directorate when the mapping carries one."),
        _c("business_unit", "string", "Business unit/section when the mapping carries one."),
        _c("mapping_source", "string", "File the mapping came from."),
        _c("is_mapped", "bool", "False for the synthesized Unmapped row."),
    ),
)

DIM_USER = TableSpec(
    name="dim_user",
    kind="dim",
    key="user_id",
    description="Observed activity users UNION the full GAL population.",
    columns=(
        _key("user_id"),
        _c("user_upn", "string", "User principal name (uppercased)."),
        _c("user_domain", "string", "Domain part of the UPN."),
        _fk("department_id", "FK to dim_department."),
        _c("division", "string",
           "GAL CompanyName falling back to Department; 'Unknown' when unmapped. "
           "The primary org lens (Department is ~one value tenant-wide)."),
        _c("region", "string",
           "OU directly under the GAL OnPremisesDN 'Regions' OU "
           "(Central/Kedron/South East/...); 'Unknown' for non-Regions DNs."),
        _c("job_title", "string", "GAL JobTitle when the mapping carries one."),
        _c("is_leaver", "bool",
           "GAL OnPremisesDN sits in a Leavers OU (account departed the org)."),
        _c("is_generic_account", "bool",
           "GAL OnPremisesDN sits in a shared-account pool OU (Generic "
           "Accounts/SharedUsers). Kept separate from is_service_account, "
           "which is a UPN naming-pattern heuristic that also covers "
           "activity-only users with no GAL row."),
        _c("is_service_account", "bool", "Matches service-account naming patterns."),
        _c("has_activity", "bool", "True when the user appears in the activity data; False for GAL-only rows."),
    ),
)

DIM_SIT = TableSpec(
    name="dim_sit",
    kind="dim",
    key="sit_key",
    description=(
        "Sensitive information types: every risk-workbook row (observed or not) "
        "plus generated rows for observed SITs missing from the workbook."
    ),
    columns=(
        _key("sit_key", "string", "GUID (lowercase), slug, or name:<normalized name>."),
        _c("sit_name", "string", "Display name."),
        _c("sit_id", "string", "SIT GUID when known."),
        _c("sit_slug", "string", "Custom SIT slug when the identifier is not a GUID."),
        _c("category", "string", "Workbook category."),
        _c("risk_score", "int64", "Risk rating 1-10 from the workbook."),
        _c("risk_band", "string", "Low/Medium/High/Critical/Unrated."),
        _c("risk_description", "string", "Workbook risk description."),
        _c("reference_url", "string", "Reference URL."),
        _c("pspf_classification", "string", "Australian PSPF classification."),
        _c("qgiscf", "string", "QGISCF classification."),
        _c("qgiscf_dlm", "string", "QGISCF DLM."),
        _c("label_code", "string", "Label code."),
        _c("sit_classifier_type", "string", "Workbook classifier type."),
        _c("source", "string", "SIT source (Microsoft built-in / custom)."),
        _c("jurisdictions", "string", "Jurisdictions."),
        _c("scope", "string", "Scope."),
        _c("reference_confidence", "string", "Workbook confidence."),
        _c("classification_tier", "string", "Classification tier."),
        _c("generic_classification", "string", "Generic classification."),
        _c("generic_dlm", "string", "Generic DLM."),
        _c("data_categories", "string", "Data categories (cross-reference sheet)."),
        _c("regulations", "string", "Regulations (cross-reference sheet)."),
        _c("small_tenant", "bool", "Recommended for small tenants."),
        _c("medium_tenant", "bool", "Recommended for medium tenants."),
        _c("large_tenant", "bool", "Recommended for large tenants."),
        _c("source_sheet", "string", "Worksheet (or generator) the row came from."),
        _c("is_unrated", "bool", "True when the workbook has no risk rating for this SIT."),
        _c("observed", "bool", "True when this SIT appeared in the activity detections."),
    ),
)

DIM_ACTIVITY_TYPE = TableSpec(
    name="dim_activity_type",
    kind="dim",
    key="activity_type_id",
    columns=(
        _key("activity_type_id"),
        _c("activity", "string", "Raw activity name."),
        _c("activity_group", "string", "Copilot/DLP/Email/Egress/File/Other."),
        _c("is_egress", "bool", "Activity is an egress action."),
        _c("is_copilot", "bool", "Activity is Copilot related."),
    ),
)

DIM_WORKLOAD = TableSpec(
    name="dim_workload",
    kind="dim",
    key="workload_id",
    columns=(
        _key("workload_id"),
        _c("workload", "string", "Workload (Exchange/SharePoint/OneDrive/Teams/Endpoint/Copilot/...)."),
    ),
)

DIM_LOCATION = TableSpec(
    name="dim_location",
    kind="dim",
    key="location_id",
    description="Folder-level locations extracted from file paths.",
    columns=(
        _key("location_id"),
        _c("folder_path", "string", "Folder portion of the file path."),
        _c("folder_name", "string", "Last path segment."),
        _c("path_depth", "int64", "Number of path segments."),
    ),
)

DIM_FILE = TableSpec(
    name="dim_file",
    kind="dim",
    key="file_id",
    columns=(
        _key("file_id"),
        _c("file_path", "string", "Full file path when present."),
        _c("file_name", "string", "File or item name."),
        _c("file_extension", "string", "Raw FileExtension field, else derived from the name."),
        _c("file_type", "string", "Raw FileType field."),
    ),
)

DIM_POLICY = TableSpec(
    name="dim_policy",
    kind="dim",
    key="policy_rule_id",
    description="DLP policy/rule combinations seen in PolicyMatchInfo.",
    columns=(
        _key("policy_rule_id"),
        _c("policy_id", "string", "Policy GUID."),
        _c("policy_name", "string", "Policy display name."),
        _c("policy_mode", "string", "Enable/TestWithNotifications/..."),
        _c("rule_id", "string", "Rule GUID."),
        _c("rule_name", "string", "Rule display name."),
        _c("rule_actions", "string", "Rule actions (JSON when complex)."),
        _c("condition", "string", "Match condition (JSON when complex)."),
    ),
)

DIM_DOMAIN = TableSpec(
    name="dim_domain",
    kind="dim",
    key="domain_id",
    columns=(
        _key("domain_id"),
        _c("domain", "string", "Full domain (lowercase)."),
        _c("parent_domain", "string", "Last three labels (or the domain itself)."),
        _c("tld", "string", "Top-level domain."),
    ),
)

DIM_EMAIL_ADDRESS = TableSpec(
    name="dim_email_address",
    kind="dim",
    key="email_address_id",
    columns=(
        _key("email_address_id"),
        _c("email_address", "string", "Email address (lowercase)."),
        _c("local_part", "string", "Address before the @."),
        _fk("domain_id", "FK to dim_domain."),
    ),
)

DIM_APP_IDENTITY = TableSpec(
    name="dim_app_identity",
    kind="dim",
    key="app_identity_id",
    description="AI/app identities observed on Copilot-enriched activities.",
    columns=(
        _key("app_identity_id"),
        _c("app_identity", "string", "AppIdentity value."),
        _c("app_identity_category", "string", "AppIdentityCategory value."),
        _c("app_identity_group", "string", "AppIdentityGroup value."),
        _c("purview_ai_app_name", "string", "PurviewAIAppName value."),
    ),
)

DIM_SOURCE_PAGE = TableSpec(
    name="dim_source_page",
    kind="dim",
    key="page_id",
    description="Provenance: one row per source export page file.",
    columns=(
        _key("page_id"),
        _c("source_file", "string", "Page path relative to the export root."),
        _c("page_number", "int64", "PageNumber from the page wrapper."),
        _c("export_timestamp", "timestamp_us", "ExportTimestamp from the page wrapper.", fmt="yyyy-MM-dd HH:mm:ss"),
        _c("watermark", "string", "WaterMark/PageCookie from the page wrapper."),
        _c("record_count", "int64", "RecordCount declared by the page wrapper."),
    ),
)

DIM_TABLES = (
    DIM_DATE, DIM_DEPARTMENT, DIM_USER, DIM_SIT, DIM_ACTIVITY_TYPE,
    DIM_WORKLOAD, DIM_LOCATION, DIM_FILE, DIM_POLICY, DIM_DOMAIN,
    DIM_EMAIL_ADDRESS, DIM_APP_IDENTITY, DIM_SOURCE_PAGE,
)
