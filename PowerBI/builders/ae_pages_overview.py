"""Activity Explorer report pages: Overview, Risk/SIT analysis, and People
groups (000-2xx). Legacy-page provenance is noted per function; the full
legacy -> new mapping lives in build_activity_explorer.LEGACY_PAGE_MAPPING.
"""

from __future__ import annotations

from . import ae_fields as f
from .filters import COMPARISON_GE, COMPARISON_GT, measure_threshold_filter
from .report_layout import (
    CHART_HEIGHT,
    CHART_ROW_Y,
    PageSpec,
    TABLE_HEIGHT,
    TABLE_ROW_Y,
    full_width,
    grid_row,
    row_of_cards,
    title_rect,
)
from .visual_factories import (
    bar_chart,
    card,
    column_chart,
    line_chart,
    pie_chart,
    pivot_table,
    table,
    textbox,
    treemap,
)

# Tall main-content band for pages whose hero visual replaces the chart+table rows.
TALL_HEIGHT = 720 - CHART_ROW_Y - 40  # 496


def executive_overview_page() -> PageSpec:
    """000: legacy 'Executive Overview' merged with the interim report's
    Executive page.

    T6 polish 3: the department visuals became division (dim_user) — on this
    tenant Department is a single QFES value, while the GAL CompanyName-derived
    division carries the real org split. Division groupings resolve via
    fact_activity_sit.user_id -> dim_user, so the risk-pressure visuals use the
    fact-grain [Weighted Risk Score] (identical semantics to the agg-table
    [Department Risk Pressure]: SUM of risk_weighted_count). A Leaver
    Activities KPI surfaces the GAL is_leaver flag."""
    kpis = row_of_cards(6)
    charts = grid_row(3, CHART_ROW_Y, CHART_HEIGHT)
    tables = grid_row(2, TABLE_ROW_Y, TABLE_HEIGHT)
    return PageSpec(
        folder="000_Executive_Overview",
        display_name="Executive Overview",
        visuals=[
            textbox("exec-title", "Activity Explorer Risk — Executive Overview", title_rect()),
            card("exec-card-activities", f.RAW_ACTIVITIES, kpis[0], title="Activities"),
            card("exec-card-sit", f.ACTIVITIES_WITH_SIT_DATA, kpis[1], title="With SIT Data"),
            card("exec-card-matches", f.TOTAL_SIT_INSTANCE_COUNT, kpis[2], title="SIT Matches"),
            card("exec-card-risk", f.WEIGHTED_RISK_SCORE, kpis[3], title="Risk Pressure"),
            card("exec-card-highconf", f.HIGH_CONFIDENCE_DETECTIONS, kpis[4],
                 title="High Confidence Detections"),
            card("exec-card-leavers", f.LEAVER_ACTIVITIES, kpis[5],
                 title="Leaver Activities"),
            bar_chart(
                "exec-bar-dlm", f.QGISCF_DLM, [f.ACTIVITIES_BY_SIT, f.TOTAL_SIT_RISK],
                charts[0], title="Activities and Risk by Classification (DLM)"),
            treemap(
                "exec-treemap-division", f.DIVISION, f.WEIGHTED_RISK_SCORE,
                charts[1], title="Division Risk Pressure"),
            column_chart(
                "exec-col-group", f.ACTIVITY_GROUP, [f.ACTIVITY_TYPE_RISK_PRESSURE],
                charts[2], title="Risk Pressure by Activity Group",
                order_by=f.ACTIVITY_TYPE_RISK_PRESSURE),
            f.by_sit_table("exec-table-rules", f.RULE_NAME, tables[0],
                           title="Top DLP Rules by SIT Activities"),
            table(
                "exec-table-division",
                [f.DIVISION, f.RISK_BAND, f.TOTAL_SIT_INSTANCE_COUNT,
                 f.WEIGHTED_RISK_SCORE, f.HIGH_CONFIDENCE_PCT],
                tables[1], title="Division Risk Summary",
                order_by=f.WEIGHTED_RISK_SCORE),
        ],
    )


def activity_summary_page() -> PageSpec:
    """010: legacy 'Activity Summary Table' + 'Summary Activity Detail' merged
    (the DLM x Workload pivot appeared on both)."""
    charts = grid_row(2, CHART_ROW_Y, CHART_HEIGHT)
    tables = grid_row(3, TABLE_ROW_Y, TABLE_HEIGHT)
    return PageSpec(
        folder="010_Activity_Summary",
        display_name="Activity Summary",
        visuals=[
            textbox("summary-title", "Activity Summary", title_rect()),
            *f.slicer_band("summary"),
            pivot_table(
                "summary-pivot-day", rows=[f.DATE], columns=[f.WORKLOAD],
                values=[f.ACTIVITIES_BY_SIT], rect=charts[0],
                title="SIT Activities by Day and Workload"),
            pivot_table(
                "summary-pivot-dlm", rows=[f.QGISCF_DLM], columns=[f.WORKLOAD],
                values=[f.ACTIVITIES_BY_SIT], rect=charts[1],
                title="SIT Activities by Classification and Workload"),
            f.by_sit_table("summary-table-domain", f.DOMAIN, tables[0],
                           title="Target Domains by SIT Activities"),
            f.by_sit_table("summary-table-dept", f.DEPARTMENT, tables[1],
                           title="Departments by SIT Activities"),
            f.by_sit_table("summary-table-activity", f.ACTIVITY, tables[2],
                           title="Activities by SIT Activities"),
        ],
    )


def timeline_page() -> PageSpec:
    """020: legacy 'Timeline' (detection trends by SIT name and by DLM)."""
    charts = grid_row(2, CHART_ROW_Y, CHART_HEIGHT)
    return PageSpec(
        folder="020_Timeline",
        display_name="Timeline",
        visuals=[
            textbox("timeline-title", "Detection Timeline", title_rect()),
            *f.slicer_band("timeline", (f.DATE, f.USER, f.DIVISION, f.QGISCF_DLM)),
            line_chart(
                "timeline-line-sit", f.DATE, [f.TOTAL_SIT_DETECTIONS], charts[0],
                title="SIT Detections over Time by SIT Name", series=f.SIT_NAME),
            line_chart(
                "timeline-line-dlm", f.DATE, [f.TOTAL_SIT_DETECTIONS], charts[1],
                title="SIT Detections over Time by Classification (DLM)",
                series=f.QGISCF_DLM),
            line_chart(
                "timeline-line-daily", f.DATE,
                [f.RAW_ACTIVITIES, f.ACTIVITIES_BY_SIT],
                full_width(TABLE_ROW_Y, TABLE_HEIGHT),
                title="Daily Activity Volume (All vs SIT-Bearing)"),
        ],
    )


def risk_assessment_page() -> PageSpec:
    """100: legacy 'Risk Assessment' (detections + risk by DLP rule)."""
    return PageSpec(
        folder="100_Risk_Assessment",
        display_name="Risk Assessment",
        visuals=[
            textbox("riskassess-title", "Risk Assessment by DLP Rule", title_rect()),
            *f.slicer_band("riskassess"),
            column_chart(
                "riskassess-col-rule", f.RULE_NAME,
                [f.TOTAL_SIT_DETECTIONS, f.TOTAL_SIT_RISK],
                full_width(CHART_ROW_Y, CHART_HEIGHT),
                title="SIT Detections and Risk by Rule",
                order_by=f.TOTAL_SIT_DETECTIONS),
            table(
                "riskassess-table-rule",
                [f.RULE_NAME, f.POLICY_NAME, f.TOTAL_SIT_DETECTIONS,
                 f.WEIGHTED_RISK_SCORE, f.ACTIVITIES_BY_SIT],
                full_width(TABLE_ROW_Y, TABLE_HEIGHT),
                title="Rule Risk Detail", order_by=f.WEIGHTED_RISK_SCORE,
                column_widths={f.RULE_NAME: 320.0, f.POLICY_NAME: 240.0}),
        ],
    )


def classifier_analysis_page() -> PageSpec:
    """110: legacy 'Classifier Analysis' (SIT inventory with category/source)."""
    main = table(
        "clsanalysis-table-sit",
        [f.SIT_NAME, f.TOTAL_SIT_DETECTIONS, f.SIT_CATEGORY, f.SIT_SOURCE,
         f.RISK_BAND, f.WEIGHTED_RISK_SCORE],
        full_width(CHART_ROW_Y, TALL_HEIGHT - TABLE_HEIGHT - 12),
        title="Classifier Inventory (Detected SITs)",
        order_by=f.TOTAL_SIT_DETECTIONS,
        column_widths={f.SIT_NAME: 320.0})
    # Legacy gate: only SITs that actually fired (Total SIT Detections >= 1).
    main.filters.append(
        measure_threshold_filter(f.TOTAL_SIT_DETECTIONS, 1, COMPARISON_GE))
    tables = grid_row(2, TABLE_ROW_Y, TABLE_HEIGHT)
    return PageSpec(
        folder="110_Classifier_Analysis",
        display_name="Classifier Analysis",
        visuals=[
            textbox("clsanalysis-title", "Classifier Analysis", title_rect()),
            *f.slicer_band("clsanalysis", (f.DATE, f.WORKLOAD, f.SIT_CATEGORY, f.QGISCF_DLM)),
            main,
            f.by_sit_table("clsanalysis-table-dlm", f.QGISCF_DLM, tables[0],
                           title="Classification (DLM) Summary"),
            table("clsanalysis-table-band",
                  [f.RISK_BAND, f.TOTAL_SIT_DETECTIONS, f.HIGH_CONFIDENCE_PCT],
                  tables[1], title="Risk Band Summary",
                  order_by=f.TOTAL_SIT_DETECTIONS),
        ],
    )


def classifier_focus_page() -> PageSpec:
    """120: legacy 'Classifier Focus' (SIT mix, folder hotspots, workload mix)."""
    charts = grid_row(3, CHART_ROW_Y, CHART_HEIGHT)
    tables = grid_row(2, TABLE_ROW_Y, TABLE_HEIGHT)
    sit_pie = pie_chart(
        "clsfocus-pie-sit", f.SIT_NAME, f.TOTAL_SIT_DETECTIONS, charts[0],
        title="Detections by SIT Name", series=f.QGISCF_DLM)
    sit_pie.filters.append(
        measure_threshold_filter(f.TOTAL_SIT_DETECTIONS, 1, COMPARISON_GT))
    folder_treemap = treemap(
        "clsfocus-treemap-folder", f.SIT_NAME, f.TOTAL_SIT_DETECTIONS, charts[1],
        title="SIT Hotspots by Folder", details=f.FOLDER_PATH)
    folder_treemap.filters.append(
        measure_threshold_filter(f.TOTAL_SIT_DETECTIONS, 1, COMPARISON_GE))
    return PageSpec(
        folder="120_Classifier_Focus",
        display_name="Classifier Focus",
        visuals=[
            textbox("clsfocus-title", "Classifier Focus", title_rect()),
            *f.slicer_band("clsfocus"),
            sit_pie,
            folder_treemap,
            pie_chart(
                "clsfocus-pie-workload", f.WORKLOAD, f.TOTAL_SIT_DETECTIONS,
                charts[2], title="Detections by Workload"),
            f.by_sit_table("clsfocus-table-dlm", f.QGISCF_DLM, tables[0],
                           title="Classification (DLM) Summary"),
            table("clsfocus-table-sit",
                  [f.SIT_NAME, f.TOTAL_SIT_DETECTIONS, f.AVG_CONFIDENCE],
                  tables[1], title="SIT Confidence Summary",
                  order_by=f.TOTAL_SIT_DETECTIONS),
        ],
    )


def file_analysis_page() -> PageSpec:
    """130: legacy 'File Analysis' (file-type mix + DLM/domain/department)."""
    charts = grid_row(2, CHART_ROW_Y, CHART_HEIGHT)
    tables = grid_row(3, TABLE_ROW_Y, TABLE_HEIGHT)
    return PageSpec(
        folder="130_File_Analysis",
        display_name="File Analysis",
        visuals=[
            textbox("fileanalysis-title", "File Analysis", title_rect()),
            *f.slicer_band("fileanalysis"),
            pie_chart(
                "fileanalysis-pie-type", f.FILE_TYPE, f.ACTIVITIES_BY_SIT,
                charts[0], title="SIT Activities by File Type"),
            bar_chart(
                "fileanalysis-bar-type", f.FILE_TYPE,
                [f.TOTAL_RISK, f.ACTIVITIES_BY_SIT], charts[1],
                title="Risk and SIT Activities by File Type"),
            f.by_sit_table("fileanalysis-table-dlm", f.QGISCF_DLM, tables[0],
                           title="Classification (DLM) Summary"),
            f.by_sit_table("fileanalysis-table-domain", f.DOMAIN, tables[1],
                           title="Target Domains"),
            f.by_sit_table("fileanalysis-table-division", f.DIVISION, tables[2],
                           title="Divisions"),
        ],
    )


def department_analysis_page() -> PageSpec:
    """200: legacy 'Department Analysis', rebound to division (T6 polish 3).

    Department on this tenant is one wall-to-wall QFES value; the GAL
    CompanyName-derived division is the real org axis. The rollup table moved
    from the agg_department measures to fact-grain fact_activity_sit measures
    (same semantics) so the division/region slicers actually filter it —
    agg_department_sit_day carries no user_id and cannot respond to dim_user
    filters. Department itself stays available on 010/370 tables, the 5xx
    drillthrough pages, and as a drill field."""
    charts = grid_row(3, CHART_ROW_Y, CHART_HEIGHT)
    return PageSpec(
        folder="200_Department_Analysis",
        display_name="Division Analysis",
        visuals=[
            textbox("deptanalysis-title", "Division Analysis", title_rect()),
            *f.slicer_band("deptanalysis", f.ORG_SLICERS),
            pie_chart(
                "deptanalysis-pie-dlm", f.DIVISION, f.TOTAL_SIT_DETECTIONS,
                charts[0], title="Detections by Division and Classification",
                series=f.QGISCF_DLM),
            pivot_table(
                "deptanalysis-pivot-category", rows=[f.DIVISION],
                columns=[f.SIT_CATEGORY], values=[f.TOTAL_SIT_DETECTIONS],
                rect=charts[1], title="Detections by Division and SIT Category"),
            treemap(
                "deptanalysis-treemap-domain", f.DOMAIN, f.TOTAL_SIT_DETECTIONS,
                charts[2], title="Detections by Target Domain"),
            table(
                "deptanalysis-table-division",
                [f.DIVISION, f.TOTAL_SIT_INSTANCE_COUNT, f.WEIGHTED_RISK_SCORE,
                 f.AVG_WEIGHTED_RISK, f.HIGH_CONFIDENCE_PCT],
                full_width(TABLE_ROW_Y, TABLE_HEIGHT),
                title="Division Rollup",
                order_by=f.WEIGHTED_RISK_SCORE),
        ],
    )


def department_treemap_page() -> PageSpec:
    """210: legacy 'TreeDept' (treemap, detections > 50 gate), rebound from
    department to division (T6 polish 3 — see department_analysis_page)."""
    division_treemap = treemap(
        "depttree-treemap", f.DIVISION, f.TOTAL_SIT_DETECTIONS,
        full_width(CHART_ROW_Y, TALL_HEIGHT),
        title="SIT Detections by Division (>50)")
    division_treemap.filters.append(
        measure_threshold_filter(f.TOTAL_SIT_DETECTIONS, 50, COMPARISON_GT))
    return PageSpec(
        folder="210_Department_Treemap",
        display_name="Division Treemap",
        visuals=[
            textbox("depttree-title", "Division Treemap", title_rect()),
            *f.slicer_band("depttree", f.ORG_SLICERS),
            division_treemap,
        ],
    )


def user_investigation_page() -> PageSpec:
    """220: legacy 'User' (single-user activity evidence table).

    Evidence is SIT-grain: every bare column binds a dim that
    fact_activity_sit reaches via an active M:1 relationship, and the
    [Activities by SIT] measure mediates the multi-dim join (bare dim columns
    alone cannot be joined across a fact's many-side). The legacy SOURCE_FILE
    (fact_activity_detail) column is dropped: fact_activity_detail has no
    unambiguous relationship path to dim_sit / fact_activity_sit.

    T6 polish 3: evidence gains the GAL job_title / is_leaver columns, the
    band gains division + region, and a flagged-accounts table lists leaver /
    generic accounts that actually generated activity (gated on the
    [Flagged Account Activities] OR-measure — a visual filter cannot OR two
    columns)."""
    flagged = table(
        "userpage-table-flagged",
        [f.USER, f.DIVISION, f.JOB_TITLE, f.IS_LEAVER, f.IS_GENERIC_ACCOUNT,
         f.FLAGGED_ACCOUNT_ACTIVITIES],
        full_width(CHART_ROW_Y + 312, TALL_HEIGHT - 312),
        title="Flagged Accounts with Activity (Leavers / Generic Accounts)",
        order_by=f.FLAGGED_ACCOUNT_ACTIVITIES,
        column_widths={f.USER: 240.0, f.JOB_TITLE: 200.0})
    flagged.filters.append(
        measure_threshold_filter(f.FLAGGED_ACCOUNT_ACTIVITIES, 0, COMPARISON_GT))
    return PageSpec(
        folder="220_User_Investigation",
        display_name="User Investigation",
        visuals=[
            textbox("userpage-title", "User Investigation", title_rect()),
            *f.slicer_band("userpage", (f.USER, f.DATE, f.DIVISION, f.REGION, f.ACTIVITY)),
            table(
                "userpage-table-evidence",
                [f.USER, f.JOB_TITLE, f.IS_LEAVER, f.ACTIVITY, f.DATE, f.DOMAIN,
                 f.RULE_NAME, f.SIT_NAME, f.FILE_NAME, f.ACTIVITIES_BY_SIT],
                full_width(CHART_ROW_Y, 300),
                title="User Activity Evidence",
                column_widths={f.USER: 200.0, f.SIT_NAME: 220.0, f.FILE_NAME: 200.0,
                               f.RULE_NAME: 200.0, f.JOB_TITLE: 160.0}),
            flagged,
        ],
    )


def overview_pages() -> list[PageSpec]:
    return [
        executive_overview_page(),
        activity_summary_page(),
        timeline_page(),
        risk_assessment_page(),
        classifier_analysis_page(),
        classifier_focus_page(),
        file_analysis_page(),
        department_analysis_page(),
        department_treemap_page(),
        user_investigation_page(),
    ]
