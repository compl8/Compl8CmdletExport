"""Activity Explorer measures: all 45 legacy measures ported to star-schema v6
DAX, plus the useful aggregate-table measures from the interim 9-page report.

Porting rules applied (see docs/ANALYTICS-V6-PLAN.md T3):
- [Activities by SIT]: legacy bidirectional CROSSFILTER hack ->
  DISTINCTCOUNT(fact_activity_sit[activity_id]).
- Risk iterators (SUMX + LOOKUPVALUE / SUMX + RELATED) -> SUM over the
  precomputed fact_activity_sit[risk_score] / [risk_weighted_count] and
  fact_activity[activity_risk_score] columns.
- Unique Users / Files -> DISTINCTCOUNT over int surrogate FK keys.
- Workload counts -> CALCULATE([Raw Activities], dim_workload[workload] = ...).
- Time intelligence anchored to the latest DATA date
  (CALCULATE(MAX(dim_date[date]), REMOVEFILTERS())), never TODAY().
- USERELATIONSHIP measures use the SSOT's inactive target_location /
  originating_domain relationships.

Intentionally NOT re-declared from the interim report (exact duplicates of a
legacy measure): Activities With SIT (= Activities with SIT Data), Activities
With Policy (= DLP Rule Matches), Activities With Email (= Email Activities),
Raw Activity Risk Pressure (= Activity Risk Score / TotalRisk).
"""

from __future__ import annotations

from .tmdl_model import MeasureSpec

_LATEST_DATE_VAR = "VAR LastDataDate = CALCULATE ( MAX ( dim_date[date] ), REMOVEFILTERS () )"

_FACT_ACTIVITY_MEASURES = [
    # --- Key Metrics (legacy Activities home) --------------------------------
    MeasureSpec("fact_activity", "Raw Activities",
                "COUNTROWS ( fact_activity )",
                "#,##0", "Key Metrics", "One per unique Activity Explorer record."),
    MeasureSpec("fact_activity", "Total Activities",
                "COUNTROWS ( fact_activity )",
                "#,##0", "Key Metrics", "Legacy alias of [Raw Activities]."),
    MeasureSpec("fact_activity", "Unique Users",
                "DISTINCTCOUNT ( fact_activity[user_id] )",
                "#,##0", "Key Metrics"),
    MeasureSpec("fact_activity", "Unique Files",
                "DISTINCTCOUNT ( fact_activity[file_id] )",
                "#,##0", "Key Metrics"),
    MeasureSpec("fact_activity", "DLP Rule Matches",
                "CALCULATE ( [Raw Activities], fact_activity[has_policy] = TRUE () )",
                "#,##0", "Key Metrics"),
    MeasureSpec("fact_activity", "Activities with SIT Data",
                "CALCULATE ( [Raw Activities], fact_activity[has_sit] = TRUE () )",
                "#,##0", "Key Metrics"),
    MeasureSpec("fact_activity", "Total File Size (GB)",
                "DIVIDE ( SUM ( fact_activity[file_size_bytes] ), 1073741824, 0 )",
                "#,##0.00", "Key Metrics"),
    MeasureSpec("fact_activity", "Avg File Size (KB)",
                "DIVIDE ( AVERAGE ( fact_activity[file_size_bytes] ), 1024, 0 )",
                "#,##0.0", "Key Metrics"),
    MeasureSpec("fact_activity", "Email Activities",
                "CALCULATE ( [Raw Activities], fact_activity[has_email] = TRUE () )",
                "#,##0", "Key Metrics"),
    MeasureSpec("fact_activity", "HappCount",
                "COUNT ( fact_activity[happened_at] )",
                "0", "Key Metrics", "Legacy COUNT of the activity timestamp."),
    # --- Workload ------------------------------------------------------------
    MeasureSpec("fact_activity", "Endpoint Activities",
                'CALCULATE ( [Raw Activities], dim_workload[workload] = "Endpoint devices" )',
                "#,##0", "Workload"),
    MeasureSpec("fact_activity", "Exchange Activities",
                'CALCULATE ( [Raw Activities], dim_workload[workload] = "Exchange" )',
                "#,##0", "Workload"),
    MeasureSpec("fact_activity", "SharePoint Activities",
                'CALCULATE ( [Raw Activities], dim_workload[workload] = "SharePoint" )',
                "#,##0", "Workload"),
    MeasureSpec("fact_activity", "Teams Activities",
                'CALCULATE ( [Raw Activities], dim_workload[workload] = "MicrosoftTeams" )',
                "#,##0", "Workload"),
    # --- Risk ----------------------------------------------------------------
    MeasureSpec("fact_activity", "Activity Risk Score",
                "SUM ( fact_activity[activity_risk_score] )",
                "#,##0", "Risk",
                "Precomputed at ETL: sum of risk_score * match_count per record "
                "(replaces the legacy SUMX + LOOKUPVALUE iterator)."),
    MeasureSpec("fact_activity", "TotalRisk",
                "SUM ( fact_activity[activity_risk_score] )",
                "0", "Risk", "Legacy alias used by the Sankey/graph gates."),
    MeasureSpec("fact_activity", "High Risk Activities",
                "CALCULATE ( [Raw Activities], fact_activity[activity_risk_score] >= 50 )",
                "#,##0", "Risk"),
    MeasureSpec("fact_activity", "Avg Activity Risk Score",
                "AVERAGE ( fact_activity[activity_risk_score] )",
                "#,##0.0", "Risk"),
    MeasureSpec("fact_activity", "Average Activity Risk",
                "AVERAGE ( fact_activity[max_sit_risk_score] )",
                "0.0", "Risk", "Average of the highest single SIT risk per record."),
    # --- Time Intelligence (anchored to MAX data date, never TODAY()) --------
    MeasureSpec("fact_activity", "Activities Today",
                f"{_LATEST_DATE_VAR}\n"
                "RETURN CALCULATE ( [Raw Activities], dim_date[date] = LastDataDate )",
                "#,##0", "Time Intelligence",
                "Activities on the most recent data date (anchored to the data, "
                "not the wall clock)."),
    MeasureSpec("fact_activity", "Activities Last 7 Days",
                f"{_LATEST_DATE_VAR}\n"
                "RETURN CALCULATE (\n"
                "    [Raw Activities],\n"
                "    DATESINPERIOD ( dim_date[date], LastDataDate, -7, DAY )\n"
                ")",
                "#,##0", "Time Intelligence"),
    MeasureSpec("fact_activity", "Activities Last 30 Days",
                f"{_LATEST_DATE_VAR}\n"
                "RETURN CALCULATE (\n"
                "    [Raw Activities],\n"
                "    DATESINPERIOD ( dim_date[date], LastDataDate, -30, DAY )\n"
                ")",
                "#,##0", "Time Intelligence"),
    MeasureSpec("fact_activity", "Daily Activity Average",
                "AVERAGEX ( VALUES ( dim_date[date] ), [Raw Activities] )",
                "#,##0.0", "Time Intelligence"),
    # --- Domain Analysis / data flow -----------------------------------------
    MeasureSpec("fact_activity", "Target Location Activities",
                "CALCULATE (\n"
                "    [Raw Activities],\n"
                "    USERELATIONSHIP ( fact_activity[target_location_id], dim_location[location_id] )\n"
                ")",
                "#,##0", "Domain Analysis",
                "Counts activities by their TARGET folder (inactive relationship)."),
    MeasureSpec("fact_activity", "External Domain Activities",
                "CALCULATE ( [Raw Activities], NOT ( ISBLANK ( fact_activity[target_domain_id] ) ) )",
                "#,##0", "Domain Analysis"),
    MeasureSpec("fact_activity", "Unique Target Domains",
                "DISTINCTCOUNT ( fact_activity[target_domain_id] )",
                "#,##0", "Domain Analysis"),
    MeasureSpec("fact_activity", "Originating Domain Activities",
                "CALCULATE (\n"
                "    [Raw Activities],\n"
                "    USERELATIONSHIP ( fact_activity[originating_domain_id], dim_domain[domain_id] ),\n"
                "    NOT ( ISBLANK ( fact_activity[originating_domain_id] ) )\n"
                ")",
                "#,##0", "Domain Analysis"),
    # --- Policy --------------------------------------------------------------
    MeasureSpec("fact_activity", "Policy Match Count",
                "CALCULATE ( [Raw Activities], NOT ( ISBLANK ( fact_activity[policy_rule_id] ) ) )",
                "#,##0", "Policy"),
    MeasureSpec("fact_activity", "Unique Policies Triggered",
                "CALCULATE ( DISTINCTCOUNT ( dim_policy[policy_name] ), fact_policy_activity )",
                "#,##0", "Policy"),
    MeasureSpec("fact_activity", "Unique Rules Triggered",
                "CALCULATE ( DISTINCTCOUNT ( dim_policy[rule_name] ), fact_policy_activity )",
                "#,##0", "Policy"),
]

_FACT_ACTIVITY_SIT_MEASURES = [
    MeasureSpec("fact_activity_sit", "Activities by SIT",
                "DISTINCTCOUNT ( fact_activity_sit[activity_id] )",
                "0", "SIT",
                "Workhorse: distinct activities carrying a SIT detection "
                "(replaces the legacy bidirectional CROSSFILTER)."),
    MeasureSpec("fact_activity_sit", "Total SIT Detections",
                "COUNTROWS ( fact_activity_sit )",
                "#,##0", "SIT"),
    MeasureSpec("fact_activity_sit", "Total SIT Instance Count",
                "SUM ( fact_activity_sit[match_count] )",
                "#,##0", "SIT"),
    MeasureSpec("fact_activity_sit", "Avg Confidence",
                "AVERAGE ( fact_activity_sit[confidence] )",
                "#,##0", "SIT"),
    MeasureSpec("fact_activity_sit", "High Confidence Detections",
                "CALCULATE ( COUNTROWS ( fact_activity_sit ), fact_activity_sit[confidence] >= 85 )",
                "#,##0", "SIT"),
    MeasureSpec("fact_activity_sit", "Unique SIT Types Detected",
                "DISTINCTCOUNT ( fact_activity_sit[sit_key] )",
                "#,##0", "SIT"),
    MeasureSpec("fact_activity_sit", "High Confidence %",
                "DIVIDE (\n"
                "    SUM ( fact_activity_sit[high_confidence_count] ),\n"
                "    SUM ( fact_activity_sit[match_count] )\n"
                ")",
                "0.0%", "SIT"),
    # --- Risk (legacy SIT_Reference/SIT_Detections homes) --------------------
    MeasureSpec("fact_activity_sit", "Total SIT Risk",
                "SUM ( fact_activity_sit[risk_score] )",
                "#,##0", "Risk",
                "Precomputed risk per detection row (replaces SUMX + RELATED)."),
    MeasureSpec("fact_activity_sit", "Weighted Risk Score",
                "SUM ( fact_activity_sit[risk_weighted_count] )",
                "#,##0", "Risk",
                "Precomputed risk_score * match_count (replaces SUMX + RELATED)."),
    MeasureSpec("fact_activity_sit", "Avg Weighted Risk",
                "DIVIDE ( [Weighted Risk Score], [Total SIT Instance Count], 0 )",
                "#,##0.0", "Risk"),
    MeasureSpec("fact_activity_sit", "High Risk Detections",
                "CALCULATE ( COUNTROWS ( fact_activity_sit ), dim_sit[risk_score] >= 8 )",
                "#,##0", "Risk"),
    MeasureSpec("fact_activity_sit", "Protected Classification Count",
                'CALCULATE ( COUNTROWS ( fact_activity_sit ), dim_sit[pspf_classification] = "PROTECTED" )',
                "#,##0", "Risk"),
    MeasureSpec("fact_activity_sit", "Critical SIT Events",
                'CALCULATE ( [Activities by SIT], dim_sit[risk_band] = "Critical" )',
                "#,##0", "Risk"),
    MeasureSpec("fact_activity_sit", "High Critical SIT Events",
                'CALCULATE ( [Activities by SIT], dim_sit[risk_band] IN { "High", "Critical" } )',
                "#,##0", "Risk"),
    MeasureSpec("fact_activity_sit", "Unrated SIT Events",
                "CALCULATE ( [Activities by SIT], dim_sit[is_unrated] = TRUE () )",
                "#,##0", "Risk"),
]

_DIM_SIT_MEASURES = [
    MeasureSpec("dim_sit", "Avg Risk Rating",
                "AVERAGE ( dim_sit[risk_score] )", "#,##0.0", "Risk"),
    MeasureSpec("dim_sit", "Max Risk Rating",
                "MAX ( dim_sit[risk_score] )", "#,##0", "Risk"),
    MeasureSpec("dim_sit", "High Risk SITs (8+)",
                "CALCULATE ( COUNTROWS ( dim_sit ), dim_sit[risk_score] >= 8 )",
                "#,##0", "Risk"),
]

_EMAIL_MEASURES = [
    MeasureSpec("fact_email_recipient", "Total Email Recipients",
                "COUNTROWS ( fact_email_recipient )", "#,##0", "Email"),
    MeasureSpec("fact_email_recipient", "External Email Recipients",
                "CALCULATE (\n"
                "    COUNTROWS ( fact_email_recipient ),\n"
                "    NOT ( ISBLANK ( fact_email_recipient[recipient_domain_id] ) )\n"
                ")",
                "#,##0", "Email"),
    MeasureSpec("fact_email_recipient", "Unique Receiver Domains",
                "DISTINCTCOUNT ( fact_email_recipient[recipient_domain_id] )",
                "#,##0", "Email"),
]

_COPILOT_MEASURES = [
    MeasureSpec("fact_copilot_interaction", "Copilot Interactions",
                "COUNTROWS ( fact_copilot_interaction )", "#,##0", "Copilot"),
    MeasureSpec("fact_copilot_interaction", "Copilot File References",
                "CALCULATE (\n"
                "    [Copilot Interactions],\n"
                "    fact_copilot_interaction[are_files_referenced] = TRUE ()\n"
                ")",
                "#,##0", "Copilot"),
    MeasureSpec("fact_copilot_interaction", "Copilot Sensitive File References",
                "CALCULATE (\n"
                "    [Copilot Interactions],\n"
                "    fact_copilot_interaction[are_sensitive_files_referenced] = TRUE ()\n"
                ")",
                "#,##0", "Copilot"),
]

# Aggregate-table rollups (from the interim 9-page report): fast paths for
# overview pages where day x SIT grain is enough.
_AGG_MEASURES = [
    MeasureSpec("agg_department_sit_day", "Department SIT Activity Events",
                "SUM ( agg_department_sit_day[activity_count] )", "#,##0", "Aggregates"),
    MeasureSpec("agg_department_sit_day", "Department SIT Matches",
                "SUM ( agg_department_sit_day[match_count] )", "#,##0", "Aggregates"),
    MeasureSpec("agg_department_sit_day", "Department Risk Pressure",
                "SUM ( agg_department_sit_day[risk_weighted_count] )", "#,##0", "Aggregates"),
    MeasureSpec("agg_department_sit_day", "Department High Confidence Matches",
                "SUM ( agg_department_sit_day[high_confidence_count] )", "#,##0", "Aggregates"),
    MeasureSpec("agg_department_sit_day", "Department Avg Risk Per Match",
                "DIVIDE ( [Department Risk Pressure], [Department SIT Matches] )",
                "0.0", "Aggregates"),
    MeasureSpec("agg_department_sit_day", "Department High Confidence %",
                "DIVIDE ( [Department High Confidence Matches], [Department SIT Matches] )",
                "0.0%", "Aggregates"),
    MeasureSpec("agg_department_sit_day", "Department Critical Events",
                'CALCULATE ( [Department SIT Activity Events], dim_sit[risk_band] = "Critical" )',
                "#,##0", "Aggregates"),
    MeasureSpec("agg_department_sit_day", "Department High Critical Events",
                'CALCULATE ( [Department SIT Activity Events], dim_sit[risk_band] IN { "High", "Critical" } )',
                "#,##0", "Aggregates"),
    MeasureSpec("agg_department_sit_day", "Department Unrated Events",
                "CALCULATE ( [Department SIT Activity Events], dim_sit[is_unrated] = TRUE () )",
                "#,##0", "Aggregates"),
    MeasureSpec("agg_activity_type_sit_day", "Activity Type SIT Events",
                "SUM ( agg_activity_type_sit_day[activity_count] )", "#,##0", "Aggregates"),
    MeasureSpec("agg_activity_type_sit_day", "Activity Type SIT Matches",
                "SUM ( agg_activity_type_sit_day[match_count] )", "#,##0", "Aggregates"),
    MeasureSpec("agg_activity_type_sit_day", "Activity Type Risk Pressure",
                "SUM ( agg_activity_type_sit_day[risk_weighted_count] )", "#,##0", "Aggregates"),
    MeasureSpec("agg_activity_type_sit_day", "Activity Type High Confidence Matches",
                "SUM ( agg_activity_type_sit_day[high_confidence_count] )", "#,##0", "Aggregates"),
    MeasureSpec("agg_location_sit_day", "Location SIT Activity Events",
                "SUM ( agg_location_sit_day[activity_count] )", "#,##0", "Aggregates"),
    MeasureSpec("agg_location_sit_day", "Location SIT Matches",
                "SUM ( agg_location_sit_day[match_count] )", "#,##0", "Aggregates"),
    MeasureSpec("agg_location_sit_day", "Location Risk Pressure",
                "SUM ( agg_location_sit_day[risk_weighted_count] )", "#,##0", "Aggregates"),
    MeasureSpec("agg_location_sit_day", "Location High Confidence Matches",
                "SUM ( agg_location_sit_day[high_confidence_count] )", "#,##0", "Aggregates"),
    MeasureSpec("agg_user_sit_day", "User SIT Matches",
                "SUM ( agg_user_sit_day[match_count] )", "#,##0", "Aggregates"),
    MeasureSpec("agg_user_sit_day", "User Risk Pressure",
                "SUM ( agg_user_sit_day[risk_weighted_count] )", "#,##0", "Aggregates"),
]

MEASURES: list[MeasureSpec] = [
    *_FACT_ACTIVITY_MEASURES,
    *_FACT_ACTIVITY_SIT_MEASURES,
    *_DIM_SIT_MEASURES,
    *_EMAIL_MEASURES,
    *_COPILOT_MEASURES,
    *_AGG_MEASURES,
]

# The 45 measure names from the legacy 29-page report; all are re-declared
# above with identical names (tests assert this superset).
LEGACY_MEASURE_NAMES: tuple[str, ...] = (
    "Total Activities", "Unique Users", "Unique Files", "DLP Rule Matches",
    "Activities with SIT Data", "Total File Size (GB)", "Avg File Size (KB)",
    "Email Activities", "Endpoint Activities", "Exchange Activities",
    "SharePoint Activities", "Teams Activities", "Activity Risk Score",
    "High Risk Activities", "Avg Activity Risk Score", "Activities Today",
    "Activities Last 7 Days", "Activities Last 30 Days", "Daily Activity Average",
    "Target Location Activities", "External Domain Activities",
    "Unique Target Domains", "Originating Domain Activities", "Policy Match Count",
    "Unique Policies Triggered", "Unique Rules Triggered", "TotalRisk", "HappCount",
    "Activities by SIT", "External Email Recipients", "Unique Receiver Domains",
    "Total Email Recipients", "Total SIT Detections", "Total SIT Instance Count",
    "Avg Confidence", "High Confidence Detections", "Unique SIT Types Detected",
    "Avg Risk Rating", "Max Risk Rating", "High Risk SITs (8+)", "Total SIT Risk",
    "Weighted Risk Score", "Avg Weighted Risk", "High Risk Detections",
    "Protected Classification Count",
)
