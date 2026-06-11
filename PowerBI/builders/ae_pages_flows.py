"""Activity Explorer report pages: Locations/Movement (3xx) and
Email/Policy/AI (4xx) groups. The legacy Sankey pages keep their
[TotalRisk] > 100 measure-threshold gates; the legacy Graph page's fallback
matrix is kept AND upgraded with a ForceGraph department -> domain view.
"""

from __future__ import annotations

from . import ae_fields as f
from .filters import (
    COMPARISON_GT,
    categorical_in_filter,
    measure_threshold_filter,
)
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
    column_chart,
    card,
    force_graph,
    line_chart,
    pie_chart,
    pivot_table,
    sankey,
    scatter_chart,
    table,
    textbox,
    treemap,
    word_cloud,
)

from .ae_pages_overview import TALL_HEIGHT

# Legacy gate on the Sankey/graph/word-cloud pages: hide sub-threshold flows.
SANKEY_RISK_GATE = 100

# Removable-media / remote-session activities (union of the legacy USB
# Breakdown page filter and the legacy Activity Detail 'Sensitive USB' values).
# Both naming styles the Activity Explorer API emits are listed: humanized
# display strings (what cloud-DLP exports such as the QFD tenant contain —
# e.g. 'DLP rule matched') and the raw enum forms used by endpoint exports
# (constants.EGRESS_ACTIVITIES style). Cloud-only exports legitimately match
# zero rows here: the page is an endpoint/removable-media surface.
USB_ACTIVITIES = (
    "File copied to removable media",
    "File created on removable media",
    "File copied to remote desktop session",
    "FileCopiedToRemovableMedia",
    "FileCreatedOnRemovableMedia",
    "FileCopiedToRemoteDesktopSession",
)


def _risk_gated(visual):
    visual.filters.append(
        measure_threshold_filter(f.TOTAL_RISK, SANKEY_RISK_GATE, COMPARISON_GT))
    return visual


def _sit_risk_gated(visual):
    """Risk gate at SIT grain — for visuals grouped by columns that only
    filter fact_activity_sit (e.g. fact_email_detail via the rollup
    relationship), where [TotalRisk] on fact_activity would not respond."""
    visual.filters.append(
        measure_threshold_filter(f.TOTAL_SIT_RISK, SANKEY_RISK_GATE, COMPARISON_GT))
    return visual


def location_hotspots_page() -> PageSpec:
    """300: legacy 'Location' (folder treemap) + aggregate location rollup."""
    return PageSpec(
        folder="300_Location_Hotspots",
        display_name="Location Hotspots",
        visuals=[
            textbox("lochot-title", "Location Hotspots", title_rect()),
            *f.slicer_band("lochot"),
            treemap(
                "lochot-treemap-folder", f.FOLDER_PATH, f.ACTIVITIES_BY_SIT,
                full_width(CHART_ROW_Y, CHART_HEIGHT),
                title="SIT Activities by Folder"),
            table(
                "lochot-table-folder",
                [f.FOLDER_PATH, f.PATH_DEPTH, f.LOCATION_SIT_MATCHES,
                 f.LOCATION_RISK_PRESSURE, f.LOCATION_HIGH_CONFIDENCE],
                full_width(TABLE_ROW_Y, TABLE_HEIGHT),
                title="Folder Risk Rollup (Aggregate)",
                order_by=f.LOCATION_RISK_PRESSURE,
                column_widths={f.FOLDER_PATH: 420.0}),
        ],
    )


def location_risk_page() -> PageSpec:
    """310: legacy 'Location Risk' (risk vs detections scatter per folder).
    Bubble size uses [Avg Weighted Risk] (fact_activity_sit) rather than the
    legacy [Avg Risk Rating] (dim_sit): a dim_sit-home measure cannot respond
    to dim_location groupings across single-direction relationships, so every
    folder bubble would get the same size."""
    return PageSpec(
        folder="310_Location_Risk",
        display_name="Location Risk",
        visuals=[
            textbox("locrisk-title", "Location Risk", title_rect()),
            *f.slicer_band("locrisk"),
            scatter_chart(
                "locrisk-scatter", f.TOTAL_SIT_RISK, f.TOTAL_SIT_DETECTIONS,
                f.FOLDER_PATH, full_width(CHART_ROW_Y, TALL_HEIGHT),
                title="Folder Risk vs Detections (size = avg weighted risk)",
                size=f.AVG_WEIGHTED_RISK, series=f.QGISCF_DLM),
        ],
    )


def domain_data_flows_page() -> PageSpec:
    """320: legacy 'Domain Data Flows' (org -> target domain Sankey). The org
    side is division since T6 polish 3 (Department is one QFES value on this
    tenant; dim_user.division resolves to fact_activity_sit via user_id)."""
    tables = grid_row(3, TABLE_ROW_Y, TABLE_HEIGHT)
    return PageSpec(
        folder="320_Domain_Data_Flows",
        display_name="Domain Data Flows",
        visuals=[
            textbox("domflow-title", "Domain Data Flows", title_rect()),
            *f.slicer_band("domflow"),
            sankey(
                "domflow-sankey", f.DIVISION, f.DOMAIN, f.ACTIVITIES_BY_SIT,
                full_width(CHART_ROW_Y, CHART_HEIGHT),
                title="External Flows by Division and Domain"),
            f.by_sit_table("domflow-table-dlm", f.QGISCF_DLM, tables[0],
                           title="Classification (DLM)"),
            f.by_sit_table("domflow-table-domain", f.DOMAIN, tables[1],
                           title="Target Domains"),
            f.by_sit_table("domflow-table-division", f.DIVISION, tables[2],
                           title="Divisions"),
        ],
    )


def location_domain_flows_page() -> PageSpec:
    """330: legacy 'Location Domain Data Flows' (folder -> domain Sankey,
    TotalRisk > 100 gate)."""
    tables = grid_row(2, TABLE_ROW_Y, TABLE_HEIGHT)
    return PageSpec(
        folder="330_Location_Domain_Flows",
        display_name="Location Domain Data Flows",
        visuals=[
            textbox("locdomflow-title", "Location Domain Data Flows", title_rect()),
            *f.slicer_band("locdomflow"),
            _risk_gated(sankey(
                "locdomflow-sankey", f.FOLDER_PATH, f.DOMAIN, f.ACTIVITIES_BY_SIT,
                full_width(CHART_ROW_Y, CHART_HEIGHT),
                title="External Flows by Folder and Domain (risk > 100)")),
            f.by_sit_table("locdomflow-table-dlm", f.QGISCF_DLM, tables[0],
                           title="Classification (DLM)"),
            f.by_sit_table("locdomflow-table-folder", f.FOLDER_PATH, tables[1],
                           title="Folders"),
        ],
    )


def folder_data_flows_page() -> PageSpec:
    """340: legacy 'Folder Data Flows' (org -> folder Sankey, TotalRisk > 100
    gate). Division replaced department in T6 polish 3."""
    tables = grid_row(2, TABLE_ROW_Y, TABLE_HEIGHT)
    return PageSpec(
        folder="340_Folder_Data_Flows",
        display_name="Folder Data Flows",
        visuals=[
            textbox("folderflow-title", "Folder Data Flows", title_rect()),
            *f.slicer_band("folderflow"),
            _risk_gated(sankey(
                "folderflow-sankey", f.DIVISION, f.FOLDER_PATH,
                f.ACTIVITIES_BY_SIT, full_width(CHART_ROW_Y, CHART_HEIGHT),
                title="Flows by Division and Folder (risk > 100)")),
            f.by_sit_table("folderflow-table-dlm", f.QGISCF_DLM, tables[0],
                           title="Classification (DLM)"),
            f.by_sit_table("folderflow-table-division", f.DIVISION, tables[1],
                           title="Divisions"),
        ],
    )


def domain_graph_page() -> PageSpec:
    """350: legacy 'Graph Domain Data Flows' — keeps the fallback matrix
    (org x domain pivot) AND adds the ForceGraph upgrade. Division replaced
    department in T6 polish 3."""
    charts = grid_row(2, CHART_ROW_Y, CHART_HEIGHT + 64)
    return PageSpec(
        folder="350_Domain_Graph",
        display_name="Graph Domain Data Flows",
        visuals=[
            textbox("domgraph-title", "Graph Domain Data Flows", title_rect()),
            *f.slicer_band("domgraph"),
            _risk_gated(force_graph(
                "domgraph-force", f.DIVISION, f.DOMAIN, f.ACTIVITIES_BY_SIT,
                f.QGISCF_DLM, charts[0],
                title="Division to Domain Network (risk > 100)")),
            _risk_gated(pivot_table(
                "domgraph-pivot", rows=[f.DIVISION], columns=[f.DOMAIN],
                values=[f.ACTIVITIES_BY_SIT], rect=charts[1],
                title="External Flows by Division and Domain (risk > 100)")),
            f.by_sit_table(
                "domgraph-table-dlm", f.QGISCF_DLM,
                full_width(TABLE_ROW_Y + 64, TABLE_HEIGHT - 64),
                title="Classification (DLM)"),
        ],
    )


def device_page() -> PageSpec:
    """360: legacy 'Device' (endpoint evidence table)."""
    return PageSpec(
        folder="360_Device_Activity",
        display_name="Device Activity",
        visuals=[
            textbox("device-title", "Device Activity", title_rect()),
            *f.slicer_band("device", (f.DATE, f.ACTIVITY, f.WORKLOAD, f.APPLICATION)),
            table(
                "device-table-evidence",
                [f.ACTIVITY, f.APPLICATION, f.DESTINATION_LOCATION_TYPE,
                 f.DEVICE_NAME, f.FILE_NAME, f.FILE_SIZE_BYTES, f.FILE_TYPE,
                 f.DATE, f.TARGET_URL],
                full_width(CHART_ROW_Y, TALL_HEIGHT),
                title="Endpoint Device Evidence",
                column_widths={f.FILE_NAME: 240.0, f.TARGET_URL: 240.0,
                               f.ACTIVITY: 180.0, f.DEVICE_NAME: 140.0}),
        ],
    )


def usb_breakdown_page() -> PageSpec:
    """370: legacy 'USB Breakdown' (removable-media/remote-session focus).
    DEVICE_NAME (fact_activity_detail) is dropped from the SIT evidence table:
    fact_activity_detail has no relationship path to fact_activity_sit, so the
    [Activities by SIT] values would cross-join over every device name."""
    tables = grid_row(2, TABLE_ROW_Y, TABLE_HEIGHT)
    return PageSpec(
        folder="370_USB_Breakdown",
        display_name="USB Breakdown",
        page_filters=[
            categorical_in_filter("dim_activity_type", "activity", list(USB_ACTIVITIES)),
        ],
        visuals=[
            textbox("usb-title", "USB Breakdown", title_rect()),
            *f.slicer_band("usb", (f.DATE, f.USER, f.DEPARTMENT, f.SIT_NAME)),
            table(
                "usb-table-evidence",
                [f.USER, f.ACTIVITY, f.SIT_NAME, f.ACTIVITIES_BY_SIT,
                 f.FOLDER_PATH],
                full_width(CHART_ROW_Y, CHART_HEIGHT),
                title="Removable Media Activity Evidence",
                order_by=f.ACTIVITIES_BY_SIT,
                column_widths={f.FOLDER_PATH: 320.0, f.SIT_NAME: 240.0}),
            table(
                "usb-table-sit", [f.QGISCF_DLM, f.SIT_NAME, f.ACTIVITIES_BY_SIT],
                tables[0], title="Classifications on Removable Media",
                order_by=f.ACTIVITIES_BY_SIT),
            f.by_sit_table("usb-table-dept", f.DEPARTMENT, tables[1],
                           title="Departments"),
        ],
    )


def dlp_policy_analysis_page() -> PageSpec:
    """400: legacy 'DLP Policy Analysis' (detections by rule / by SIT)."""
    charts = grid_row(2, CHART_ROW_Y, CHART_HEIGHT)
    return PageSpec(
        folder="400_DLP_Policy_Analysis",
        display_name="DLP Policy Analysis",
        visuals=[
            textbox("dlp-title", "DLP Policy Analysis", title_rect()),
            *f.slicer_band("dlp"),
            column_chart(
                "dlp-col-rule", f.RULE_NAME, [f.TOTAL_SIT_DETECTIONS], charts[0],
                title="SIT Detections by DLP Rule", order_by=f.TOTAL_SIT_DETECTIONS),
            column_chart(
                "dlp-col-sit", f.SIT_NAME, [f.TOTAL_SIT_DETECTIONS], charts[1],
                title="SIT Detections by SIT Name", order_by=f.TOTAL_SIT_DETECTIONS),
            table(
                "dlp-table-policy",
                [f.POLICY_NAME, f.RULE_NAME, f.POLICY_MODE, f.POLICY_MATCH_COUNT,
                 f.ACTIVITIES_BY_SIT],
                full_width(TABLE_ROW_Y, TABLE_HEIGHT),
                title="Policy and Rule Detail", order_by=f.POLICY_MATCH_COUNT,
                column_widths={f.POLICY_NAME: 280.0, f.RULE_NAME: 280.0}),
        ],
    )


def email_subject_cloud_page() -> PageSpec:
    """410: legacy 'Subject Heading Word Cloud' (subjects + email KPIs).
    The word cloud groups by fact_email_detail[subject]; SIT measures respond
    to that grouping via the fact_activity_sit -> fact_email_detail rollup
    relationship. The legacy [TotalRisk] (fact_activity) gate becomes
    [Total SIT Risk] (fact_activity_sit) — subject groupings cannot filter
    fact_activity, so the old gate would have been all-or-nothing."""
    kpis = row_of_cards(4)
    tables = grid_row(2, TABLE_ROW_Y, TABLE_HEIGHT)
    dlm_table = f.by_sit_table("subject-table-dlm", f.QGISCF_DLM, tables[0],
                               title="Classification (DLM)")
    dlm_table.filters.append(
        measure_threshold_filter(f.ACTIVITIES_BY_SIT, 0, COMPARISON_GT))
    dept_table = f.by_sit_table("subject-table-division", f.DIVISION, tables[1],
                                title="Divisions")
    dept_table.filters.append(
        measure_threshold_filter(f.ACTIVITIES_BY_SIT, 1, COMPARISON_GT))
    return PageSpec(
        folder="410_Email_Subject_Cloud",
        display_name="Subject Heading Word Cloud",
        visuals=[
            textbox("subject-title", "Email Subject Analysis", title_rect()),
            card("subject-card-email", f.EMAIL_ACTIVITIES, kpis[0],
                 title="Email Activities"),
            card("subject-card-recipients", f.TOTAL_EMAIL_RECIPIENTS, kpis[1],
                 title="Recipients"),
            card("subject-card-external", f.EXTERNAL_EMAIL_RECIPIENTS, kpis[2],
                 title="External Recipients"),
            card("subject-card-domains", f.UNIQUE_RECEIVER_DOMAINS, kpis[3],
                 title="Receiver Domains"),
            _sit_risk_gated(word_cloud(
                "subject-wordcloud", f.EMAIL_SUBJECT, f.ACTIVITIES_BY_SIT,
                full_width(CHART_ROW_Y, CHART_HEIGHT),
                title="Subject Keyword Distribution (risk > 100)")),
            dlm_table,
            dept_table,
        ],
    )


def ai_view_page() -> PageSpec:
    """420: legacy 'AI View' (trends and mixes for AI-bound flows; the legacy
    manual domain page-filter becomes a Domain slicer)."""
    charts = grid_row(2, CHART_ROW_Y, CHART_HEIGHT)
    tables = grid_row(2, TABLE_ROW_Y, TABLE_HEIGHT)
    return PageSpec(
        folder="420_AI_View",
        display_name="AI View",
        visuals=[
            textbox("aiview-title", "AI View", title_rect()),
            *f.slicer_band("aiview", (f.DATE, f.DOMAIN, f.DIVISION, f.QGISCF_DLM)),
            line_chart(
                "aiview-line-dlm", f.DATE, [f.TOTAL_SIT_DETECTIONS], charts[0],
                title="Detections over Time by Classification (DLM)",
                series=f.QGISCF_DLM),
            line_chart(
                "aiview-line-domain", f.DATE, [f.TOTAL_SIT_DETECTIONS], charts[1],
                title="Detections over Time by Target Domain", series=f.DOMAIN),
            pie_chart(
                "aiview-pie-activity", f.ACTIVITY, f.TOTAL_SIT_DETECTIONS,
                tables[0], title="Detections by Activity and Domain",
                series=f.DOMAIN),
            pie_chart(
                "aiview-pie-division", f.DIVISION, f.TOTAL_SIT_DETECTIONS,
                tables[1], title="Detections by Division and Domain",
                series=f.DOMAIN),
        ],
    )


def agent_activity_page() -> PageSpec:
    """430: legacy 'Agent Activity' upgraded onto fact_copilot_interaction +
    dim_app_identity (v6 AI enrichment). The agent-names table counts with
    [Detail Activities] (fact_activity_detail home): [Raw Activities] on
    fact_activity cannot respond to agent_name groupings across the
    single-direction fact_activity_detail -> fact_activity relationship."""
    kpis = row_of_cards(4)
    tables = grid_row(2, TABLE_ROW_Y, TABLE_HEIGHT)
    return PageSpec(
        folder="430_Agent_Activity",
        display_name="Agent Activity",
        visuals=[
            textbox("agent-title", "Agent and Copilot Activity", title_rect()),
            card("agent-card-interactions", f.COPILOT_INTERACTIONS, kpis[0],
                 title="Copilot Interactions"),
            card("agent-card-files", f.COPILOT_FILE_REFERENCES, kpis[1],
                 title="File References"),
            card("agent-card-sensitive", f.COPILOT_SENSITIVE_FILE_REFERENCES,
                 kpis[2], title="Sensitive File References"),
            card("agent-card-sit", f.ACTIVITIES_BY_SIT, kpis[3],
                 title="SIT Activities"),
            table(
                "agent-table-identity",
                [f.APP_IDENTITY, f.APP_IDENTITY_CATEGORY, f.PURVIEW_AI_APP_NAME,
                 f.AI_APP_LOCATION, f.COPILOT_INTERACTIONS],
                full_width(CHART_ROW_Y, CHART_HEIGHT),
                title="AI App Identities", order_by=f.COPILOT_INTERACTIONS),
            table(
                "agent-table-agents",
                [f.AGENT_NAME, f.TARGET_AGENT_NAME, f.ACTIVITY, f.DETAIL_ACTIVITIES],
                tables[0], title="Agent Names by Activity",
                order_by=f.DETAIL_ACTIVITIES),
            f.by_sit_table("agent-table-group", f.ACTIVITY_GROUP, tables[1],
                           title="Activity Groups"),
        ],
    )


def flows_pages() -> list[PageSpec]:
    return [
        location_hotspots_page(),
        location_risk_page(),
        domain_data_flows_page(),
        location_domain_flows_page(),
        folder_data_flows_page(),
        domain_graph_page(),
        device_page(),
        usb_breakdown_page(),
        dlp_policy_analysis_page(),
        email_subject_cloud_page(),
        ai_view_page(),
        agent_activity_page(),
    ]
