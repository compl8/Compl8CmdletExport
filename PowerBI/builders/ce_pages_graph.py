"""Content Explorer SIT Risk pages 052-070: graph/network visualisations,
Sankey flows, executive summary, data quality, and terminology.

Folder numbering deviates from the legacy project (which reused 054/055
prefixes out of nav order); displayName and page ORDER are identical, the
folders are simply renumbered 052/053 so folder sort == nav order (an engine
invariant). Every visual mirrors a legacy visual and gains a curated title.
"""

from __future__ import annotations

from .ce_pages_core import (
    KPI_BAND_NARROW,
    QGISCF_DLM,
    RISK_BAND,
    TALL_HEIGHT,
    kpi_cells,
    m,
    slicer_band,
    split_row,
)
from .expressions import col
from .report_layout import (
    CHART_HEIGHT,
    CHART_ROW_Y,
    GUTTER,
    MARGIN,
    PageSpec,
    TABLE_HEIGHT,
    TABLE_ROW_Y,
    full_width,
    grid_row,
    title_rect,
)
from .visual_factories import (
    Rect,
    bar_chart,
    card,
    force_graph,
    sankey,
    scatter_chart,
    slicer,
    table,
    textbox,
    treemap,
)

# Hero-graph layout: wide graph on the left, a stacked side column on the right.
HERO_WIDTH = 784.0
SIDE_X = MARGIN + HERO_WIDTH + GUTTER
SIDE_WIDTH = 1280 - 2 * MARGIN - HERO_WIDTH - GUTTER  # 420


def _side_stack(count: int, y: float = CHART_ROW_Y,
                total_height: float = TALL_HEIGHT) -> list[Rect]:
    height = (total_height - GUTTER * (count - 1)) / count
    return [Rect(SIDE_X, y + index * (height + GUTTER), SIDE_WIDTH, height)
            for index in range(count)]


def graph_visualizer_page() -> PageSpec:
    """052 (legacy 054_Graph_Visualizer): force-directed graph of the focused
    edge subset."""
    kpis = kpi_cells(4, KPI_BAND_NARROW)
    side = _side_stack(3)
    edge_type = col("FactGraphEdgeFocus", "edge_type", "Link Type")
    target_label = col("FactGraphEdgeFocus", "target_visual_label", "Target")
    return PageSpec(
        folder="052_Graph_Visualizer", display_name="Graph Visualizer",
        visuals=[
            textbox("graphviz-title", "Microsoft Force-Directed Graph", title_rect()),
            card("graphviz-card-edges", m("Force Graph Edges", "Focused Edges"),
                 kpis[0], title="Focused Edges"),
            card("graphviz-card-risk", m("Force Graph Risk Pressure", "Risk Pressure"),
                 kpis[1], title="Risk Pressure"),
            card("graphviz-card-files", m("Force Graph Files", "Files"), kpis[2],
                 title="Files"),
            card("graphviz-card-highconf",
                 m("Force Graph High Confidence Matches", "High Conf Matches"),
                 kpis[3], title="High Confidence Matches"),
            *slicer_band("graphviz",
                         [edge_type,
                          col("FactGraphEdgeFocus", "source_node_type", "Source Type"),
                          col("FactGraphEdgeFocus", "focus_bucket", "Focus")]),
            force_graph("graphviz-graph",
                        col("FactGraphEdgeFocus", "source_visual_label", "Source"),
                        target_label,
                        m("Force Graph Edge Weight", "Weight"), edge_type,
                        Rect(MARGIN, CHART_ROW_Y, HERO_WIDTH, TALL_HEIGHT),
                        source_type=col("FactGraphEdgeFocus", "source_node_type", "Source Type"),
                        target_type=col("FactGraphEdgeFocus", "target_node_type", "Target Type"),
                        charge=-80, title="Sensitive Data Movement Graph"),
            treemap("graphviz-treemap-type", edge_type,
                    m("Force Graph Risk Pressure", "Risk Pressure"), side[0],
                    title="Risk Pressure by Link Type"),
            bar_chart("graphviz-bar-target", target_label,
                      [m("Force Graph Risk Pressure", "Risk Pressure")], side[1],
                      title="Riskiest Targets"),
            table("graphviz-table-edges",
                  [col("FactGraphEdgeFocus", "focus_rank", "Rank"),
                   col("FactGraphEdgeFocus", "source_visual_label", "Source"),
                   target_label, edge_type, m("Force Graph Files", "Files"),
                   m("Force Graph Risk Pressure", "Risk Pressure"),
                   m("Force Graph Edge Weight", "Edge Weight")],
                  side[2], title="Focused Edge Detail",
                  order_by=m("Force Graph Risk Pressure")),
        ],
    )


def node_neighbourhood_page() -> PageSpec:
    """053 (legacy 055_Node_Neighbourhood): adjacency explorer for a selected
    node."""
    kpis = kpi_cells(4)
    filters_row = grid_row(2, 160, 74, x0=MARGIN, total_width=408)
    hero_y = 240.0
    hero_height = 720 - hero_y - 40
    side = _side_stack(3, y=CHART_ROW_Y, total_height=TALL_HEIGHT)
    rel_type = col("FactGraphAdjacency", "relationship_type", "Link Type")
    connected_label = col("FactGraphAdjacency", "connected_visual_label", "Connected")
    connected_type = col("FactGraphAdjacency", "connected_node_type", "Connected Type")
    return PageSpec(
        folder="053_Node_Neighbourhood", display_name="Node Neighbourhood",
        visuals=[
            textbox("nodehood-title", "Node Neighbourhood", title_rect()),
            card("nodehood-card-nodes", m("Graph Nodes"), kpis[0], title="Graph Nodes"),
            card("nodehood-card-connected", m("Connected Nodes"), kpis[1],
                 title="Connected Nodes"),
            card("nodehood-card-risk", m("Graph Edge Risk Pressure", "Edge Risk"),
                 kpis[2], title="Edge Risk"),
            card("nodehood-card-files", m("Graph Edge Files", "Edge Files"), kpis[3],
                 title="Edge Files"),
            *slicer_band("nodehood",
                         [col("DimGraphNode", "node_type", "Node Type"),
                          col("DimGraphNode", "node_label", "Node")]),
            slicer("nodehood-slicer-reltype", rel_type, filters_row[0],
                   title="Link Type"),
            slicer("nodehood-slicer-conntype", connected_type, filters_row[1],
                   title="Connected Type"),
            force_graph("nodehood-graph",
                        col("FactGraphAdjacency", "selected_visual_label", "Selected"),
                        connected_label, m("Graph Edge Weight", "Weight"), rel_type,
                        Rect(MARGIN, hero_y, HERO_WIDTH, hero_height),
                        source_type=col("FactGraphAdjacency", "selected_node_type", "Selected Type"),
                        target_type=connected_type,
                        charge=-110, name_max_length=56,
                        title="Selected Node Neighbourhood"),
            treemap("nodehood-treemap-type", connected_type,
                    m("Graph Edge Risk Pressure", "Edge Risk"), side[0],
                    title="Edge Risk by Connected Type"),
            bar_chart("nodehood-bar-connected", connected_label,
                      [m("Graph Edge Risk Pressure", "Edge Risk")], side[1],
                      title="Riskiest Connections"),
            table("nodehood-table-edges",
                  [col("FactGraphAdjacency", "selected_visual_label", "Selected"),
                   connected_label,
                   col("FactGraphAdjacency", "relationship_type", "Relationship"),
                   m("Graph Edge Files", "Edge Files"),
                   m("Graph Edge SIT Pairs", "File/SIT Pairs"),
                   m("Graph Edge Risk Pressure", "Edge Risk"),
                   m("Graph Edge High Confidence Matches", "High Conf Matches"),
                   m("Graph Edge Weight", "Edge Weight")],
                  side[2], title="Neighbourhood Edge Detail",
                  order_by=m("Graph Edge Risk Pressure")),
        ],
    )


def sankey_flows_page() -> PageSpec:
    """054: user->location and location->sensitivity Sankey flows."""
    kpis = kpi_cells(4, KPI_BAND_NARROW)
    flows = grid_row(2, CHART_ROW_Y, 300)
    tables = grid_row(2, 512, 720 - 512 - 40)
    return PageSpec(
        folder="054_Sankey_Flows", display_name="Sankey Flows",
        visuals=[
            textbox("sankey-title", "Sankey Flows", title_rect()),
            card("sankey-card-ulfiles",
                 m("Sankey User Location Files", "User/Location Files"), kpis[0],
                 title="User/Location Files"),
            card("sankey-card-ulrisk",
                 m("Sankey User Location Risk", "User/Location Risk"), kpis[1],
                 title="User/Location Risk"),
            card("sankey-card-sensfiles",
                 m("Sankey Sensitivity Files", "Sensitivity Files"), kpis[2],
                 title="Sensitivity Files"),
            card("sankey-card-sensrisk",
                 m("Sankey Sensitivity Risk", "Sensitivity Risk"), kpis[3],
                 title="Sensitivity Risk"),
            *slicer_band("sankey",
                         [RISK_BAND, QGISCF_DLM,
                          col("DimUserDepartment", "department", "Department")]),
            sankey("sankey-flow-userloc",
                   col("FactSankeyUserLocationFlow", "source_node", "Source"),
                   col("FactSankeyUserLocationFlow", "target_node", "Destination"),
                   m("Sankey User Location Weight", "Flow Weight"), flows[0],
                   title="User to Location Flows"),
            sankey("sankey-flow-sensitivity",
                   col("FactSankeySensitivityFlow", "source_node", "Source"),
                   col("FactSankeySensitivityFlow", "target_node", "Destination"),
                   m("Sankey Sensitivity Weight", "Flow Weight"), flows[1],
                   title="Location to Sensitivity Flows"),
            table("sankey-table-userloc",
                  [col("FactSankeyUserLocationFlow", "flow_rank", "Rank"),
                   col("FactSankeyUserLocationFlow", "department", "Department"),
                   col("FactSankeyUserLocationFlow", "user_display_name", "User"),
                   col("FactSankeyUserLocationFlow", "location_name", "Location"),
                   col("FactSankeyUserLocationFlow", "qgiscf_dlm", "QGISCF DLM"),
                   col("FactSankeyUserLocationFlow", "risk_band", "Risk Band"),
                   m("Sankey User Location Files", "Files"),
                   m("Sankey User Location Risk", "Risk")],
                  tables[0], title="Top User/Location Flows",
                  order_by=m("Sankey User Location Risk")),
            table("sankey-table-sensitivity",
                  [col("FactSankeySensitivityFlow", "flow_rank", "Rank"),
                   col("FactSankeySensitivityFlow", "location_name", "Location"),
                   col("FactSankeySensitivityFlow", "qgiscf_dlm", "QGISCF DLM"),
                   col("FactSankeySensitivityFlow", "sit_name", "SIT"),
                   col("FactSankeySensitivityFlow", "risk_band", "Risk Band"),
                   m("Sankey Sensitivity Files", "Files"),
                   m("Sankey Sensitivity Risk", "Risk")],
                  tables[1], title="Top Sensitivity Flows",
                  order_by=m("Sankey Sensitivity Risk")),
        ],
    )


def network_navigator_page() -> PageSpec:
    """055: tabular/chart navigator over the adjacency graph."""
    kpis = kpi_cells(4)
    charts = grid_row(3, CHART_ROW_Y, CHART_HEIGHT)
    bottoms = grid_row(3, TABLE_ROW_Y, TABLE_HEIGHT)
    node_table_rect = bottoms[0]
    edge_table_rect = Rect(bottoms[1].x, TABLE_ROW_Y,
                           bottoms[1].width * 2 + GUTTER, TABLE_HEIGHT)
    connected_label = col("FactGraphAdjacency", "connected_node_label", "Connected Node")
    connected_type = col("FactGraphAdjacency", "connected_node_type", "Connected Type")
    return PageSpec(
        folder="055_Network_Navigator", display_name="Network Navigator",
        visuals=[
            textbox("netnav-title", "Network Navigator", title_rect()),
            card("netnav-card-nodes", m("Graph Nodes"), kpis[0], title="Graph Nodes"),
            card("netnav-card-connected", m("Connected Nodes"), kpis[1],
                 title="Connected Nodes"),
            card("netnav-card-risk", m("Graph Edge Risk Pressure", "Edge Risk"),
                 kpis[2], title="Edge Risk"),
            card("netnav-card-files", m("Graph Edge Files", "Edge Files"), kpis[3],
                 title="Edge Files"),
            *slicer_band("netnav",
                         [col("DimGraphNode", "node_type", "Node Type"),
                          col("DimGraphNode", "node_label", "Node")]),
            treemap("netnav-treemap-type", connected_type,
                    m("Graph Edge Risk Pressure", "Edge Risk"), charts[0],
                    title="Edge Risk by Connected Type"),
            scatter_chart("netnav-scatter",
                          m("Graph Edge Files", "Edge Files"),
                          m("Graph Edge Risk Pressure", "Edge Risk"),
                          connected_label, charts[1],
                          size=m("Graph Edge Weight", "Edge Weight"),
                          title="Edge Files vs Edge Risk"),
            bar_chart("netnav-bar-connected", connected_label,
                      [m("Graph Edge Risk Pressure", "Edge Risk")], charts[2],
                      title="Riskiest Connected Nodes"),
            table("netnav-table-nodes",
                  [col("DimGraphNode", "node_type", "Type"),
                   col("DimGraphNode", "node_label", "Selected Node"),
                   col("DimGraphNode", "node_subtype", "Subtype"),
                   col("DimGraphNode", "files_with_sit", "Files"),
                   col("DimGraphNode", "file_sit_pair_count", "File/SIT Pairs"),
                   col("DimGraphNode", "risk_pressure", "Risk Pressure"),
                   col("DimGraphNode", "node_note", "Note")],
                  node_table_rect, title="Node Inventory"),
            table("netnav-table-edges",
                  [connected_type, connected_label,
                   col("FactGraphAdjacency", "relationship_type", "Relationship"),
                   m("Graph Edge Files", "Edge Files"),
                   m("Graph Edge SIT Pairs", "File/SIT Pairs"),
                   m("Graph Edge Risk Pressure", "Edge Risk"),
                   m("Graph Edge High Confidence Matches", "High Conf Matches"),
                   m("Graph Edge Weight", "Edge Weight")],
                  edge_table_rect, title="Adjacency Detail",
                  order_by=m("Graph Edge Risk Pressure")),
        ],
    )


def cluster_graph_page() -> PageSpec:
    """056: cluster-level rollup of the graph."""
    kpis = kpi_cells(4)
    charts = grid_row(3, CHART_ROW_Y, CHART_HEIGHT)
    cluster_label = col("DimGraphCluster", "cluster_label", "Cluster")
    return PageSpec(
        folder="056_Cluster_Graph", display_name="Cluster Graph",
        visuals=[
            textbox("cluster-title", "Cluster Graph", title_rect()),
            card("cluster-card-clusters", m("Graph Clusters", "Clusters"), kpis[0],
                 title="Clusters"),
            card("cluster-card-files", m("Cluster Files With SIT", "Files With SIT"),
                 kpis[1], title="Files With SIT"),
            card("cluster-card-pressure", m("Cluster Risk Pressure", "Risk Pressure"),
                 kpis[2], title="Risk Pressure"),
            card("cluster-card-density", m("Cluster Risk Density", "Risk / File"),
                 kpis[3], title="Risk per File"),
            *slicer_band("cluster",
                         [col("DimGraphCluster", "workload", "Workload"),
                          col("DimGraphCluster", "location_name", "Location")]),
            treemap("cluster-treemap", cluster_label,
                    m("Cluster Risk Pressure", "Risk Pressure"), charts[0],
                    title="Risk Pressure by Cluster"),
            scatter_chart("cluster-scatter",
                          m("Cluster Files With SIT", "Files With SIT"),
                          m("Cluster Risk Density", "Risk / File"), cluster_label,
                          charts[1], size=m("Cluster Node Size", "Cluster Size"),
                          title="Cluster Volume vs Risk Density"),
            bar_chart("cluster-bar-highconf", cluster_label,
                      [m("Cluster High Confidence Matches", "High Conf Matches")],
                      charts[2], title="High Confidence Matches by Cluster"),
            table("cluster-table",
                  [cluster_label, col("DimGraphCluster", "area_count", "Areas"),
                   m("Cluster Files With SIT", "Files With SIT"),
                   col("DimGraphCluster", "file_sit_pair_count", "File/SIT Pairs"),
                   col("DimGraphCluster", "distinct_sit_count", "Distinct SITs"),
                   m("Cluster Risk Pressure", "Risk Pressure"),
                   m("Cluster Risk Density", "Risk / File"),
                   m("Cluster High Confidence Matches", "High Conf Matches"),
                   col("DimGraphCluster", "label_coverage_pct", "Label Coverage")],
                  full_width(TABLE_ROW_Y, TABLE_HEIGHT), title="Cluster Detail",
                  order_by=m("Cluster Risk Pressure"),
                  column_widths={cluster_label: 240.0}),
        ],
    )


def executive_summary_page() -> PageSpec:
    """057: precomputed executive findings."""
    kpis = kpi_cells(4)
    charts = split_row((1, 2), CHART_ROW_Y, CHART_HEIGHT)
    finding_type = col("FactExecFinding", "finding_type", "Finding Type")
    return PageSpec(
        folder="057_Executive_Summary", display_name="Executive Summary",
        visuals=[
            textbox("exec-title", "Executive Summary", title_rect()),
            card("exec-card-findings", m("Executive Findings"), kpis[0],
                 title="Findings"),
            card("exec-card-critical",
                 m("Critical Executive Findings", "Critical Findings"), kpis[1],
                 title="Critical Findings"),
            card("exec-card-highcrit",
                 m("High Or Critical Executive Findings", "High/Critical Findings"),
                 kpis[2], title="High/Critical Findings"),
            card("exec-card-metric", m("Executive Finding Metric", "Finding Metric"),
                 kpis[3], title="Finding Metric"),
            *slicer_band("exec",
                         [col("FactExecFinding", "severity", "Severity"), finding_type]),
            treemap("exec-treemap-type", finding_type, m("Executive Findings"),
                    charts[0], title="Findings by Type"),
            bar_chart("exec-bar-entity",
                      col("FactExecFinding", "entity_label", "Entity"),
                      [m("Executive Finding Metric", "Metric")], charts[1],
                      title="Finding Metric by Entity"),
            table("exec-table-findings",
                  [col("FactExecFinding", "severity", "Severity"), finding_type,
                   col("FactExecFinding", "finding_title", "Finding"),
                   col("FactExecFinding", "finding_evidence", "Evidence"),
                   col("FactExecFinding", "recommended_action", "Recommended Action"),
                   col("FactExecFinding", "metric_name", "Metric"),
                   col("FactExecFinding", "metric_value", "Value"),
                   col("FactExecFinding", "secondary_metric_name", "Secondary Metric"),
                   col("FactExecFinding", "secondary_metric_value", "Secondary Value")],
                  full_width(TABLE_ROW_Y, TABLE_HEIGHT), title="Executive Findings",
                  order_by=col("FactExecFinding", "finding_rank"),
                  column_widths={
                      col("FactExecFinding", "finding_title", "Finding"): 220.0,
                      col("FactExecFinding", "finding_evidence", "Evidence"): 260.0,
                      col("FactExecFinding", "recommended_action",
                          "Recommended Action"): 240.0}),
        ],
    )


def data_quality_page() -> PageSpec:
    """060: unrated SITs and labelling coverage."""
    kpis = kpi_cells(4)
    charts = split_row((1, 2), CHART_ROW_Y, CHART_HEIGHT)
    tables = grid_row(2, TABLE_ROW_Y, TABLE_HEIGHT)
    sit_name = col("DimSIT", "sit_name", "SIT")
    file_ext = col("DimFile", "file_extension", "Extension")
    return PageSpec(
        folder="060_Data_Quality", display_name="Data Quality",
        visuals=[
            textbox("quality-title", "Data Quality And Coverage", title_rect()),
            card("quality-card-unrated", m("Unrated SITs"), kpis[0],
                 title="Unrated SITs"),
            card("quality-card-items", m("Unrated Aggregate Items", "Unrated Items"),
                 kpis[1], title="Unrated Items"),
            card("quality-card-unlabelled",
                 m("Unlabelled Files With SIT", "Unlabelled Files"), kpis[2],
                 title="Unlabelled Files"),
            card("quality-card-coverage",
                 m("Sensitivity Label Coverage %", "Sensitivity Coverage"), kpis[3],
                 title="Sensitivity Coverage"),
            *slicer_band("quality",
                         [col("DimLocation", "workload", "Workload"), QGISCF_DLM]),
            bar_chart("quality-bar-band", RISK_BAND,
                      [m("Unlabelled Files With SIT")], charts[0],
                      title="Unlabelled Files by Risk Band"),
            bar_chart("quality-bar-ext", file_ext,
                      [m("Sensitivity Label Coverage %")], charts[1],
                      title="Label Coverage by Extension"),
            table("quality-table-unrated",
                  [sit_name, col("DimSIT", "category", "Category"),
                   col("DimSIT", "is_unrated", "Is Unrated"),
                   col("DimSIT", "source_sheet", "Risk Source"),
                   m("Aggregate Items"), m("Unrated Aggregate Items")],
                  tables[0], title="SIT Risk Mapping Coverage",
                  order_by=m("Unrated Aggregate Items"),
                  column_widths={sit_name: 220.0}),
            table("quality-table-labels",
                  [file_ext, m("Unique Files"), m("Files With Sensitivity Label"),
                   m("Files With Retention Label"), m("Sensitivity Label Coverage %")],
                  tables[1], title="Label Coverage by Extension",
                  order_by=m("Unique Files")),
        ],
    )


_TERMINOLOGY: tuple[str, ...] = (
    "Sensitive Information Type (SIT): a Microsoft Purview classifier for data "
    "such as identifiers, financial records, credentials, health information, "
    "or other regulated content.\n\nExported SIT: the SIT/classifier folder "
    "that produced the Content Explorer export page.\n\nDetected SIT: a SIT ID "
    "found inside the detailed SensitiveInfoTypesData payload for a file.",
    "Location: a Content Explorer source such as a SharePoint site, "
    "Teams-backed SharePoint site, or OneDrive.\n\nArea: a leaf folder within "
    "a location, derived from each file relative path. Area pages use location "
    "plus folder path, so the same folder name in two locations remains "
    "separate.\n\nFolder depth: the number of folders in the relative path "
    "before the file name.",
    "Files With SIT: distinct files that have at least one exported SIT "
    "match.\n\nFile/SIT pair: one file matched to one SIT. A file with five "
    "SITs contributes five file/SIT pairs.\n\nSITs / File: file/SIT pairs "
    "divided by Files With SIT. Higher values indicate denser mixed-sensitive "
    "content.\n\nMatch count: the confidence-count total reported by Content "
    "Explorer for matched SITs.",
    "Risk score: the numeric score from the SIT risk workbook.\n\nRisk band: "
    "Low, Medium, High, or Critical, derived from the workbook risk "
    "score.\n\nRisk pressure: sum of risk score across file/SIT pairs. This "
    "intentionally rewards both higher risk SITs and higher volume.\n\nRisk / "
    "File: risk pressure divided by distinct files with SIT in the selected "
    "area.",
    "Unrated SIT: a SIT found in the export but not mapped to a risk score in "
    "the supplied workbook. It is tracked separately because treating it as "
    "low risk would hide mapping gaps.\n\nUnlabelled file: a matched file "
    "where the sensitivity label field is blank.\n\nArea label coverage: "
    "sensitivity-labelled files divided by all matched files in the area.",
    "QGISCF DLM and PSPF: risk workbook fields used to connect Microsoft SITs "
    "to Queensland Government and PSPF classification concepts.\n\nDrilldown: "
    "select a location, folder, risk band, SIT, or treemap segment to "
    "cross-filter the tables on the page. Area Drilldown then shows both SIT "
    "composition and file-level evidence.",
)


def terminology_page() -> PageSpec:
    """070: report terminology reference."""
    rows = [grid_row(3, 72, 176), grid_row(3, 286, 176)]
    boxes = [
        textbox(f"terms-box-{index}", text, cell, font_size=11, bold=False)
        for index, (text, cell) in enumerate(zip(_TERMINOLOGY,
                                                 [*rows[0], *rows[1]]))
    ]
    return PageSpec(
        folder="070_Terminology", display_name="Terminology",
        visuals=[
            textbox("terms-title", "Terminology", title_rect()),
            *boxes,
        ],
    )


def graph_pages() -> list[PageSpec]:
    return [
        graph_visualizer_page(),
        node_neighbourhood_page(),
        sankey_flows_page(),
        network_navigator_page(),
        cluster_graph_page(),
        executive_summary_page(),
        data_quality_page(),
        terminology_page(),
    ]
