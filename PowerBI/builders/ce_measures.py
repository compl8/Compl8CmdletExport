"""Content Explorer SIT Risk measures: all 71 measures from the legacy
generated report, declared on their legacy home table (FactLocationSIT).

Porting rules (names and semantics identical to the legacy report):
- DAX kept verbatim except three measures whose `CALCULATE ( ..., FILTER (
  DimFile, ... ) )` whole-table iterators are rewritten as equivalent
  column-predicate filters (far cheaper, identical results — in DAX,
  `x = ""` is TRUE for blank, so ISBLANK(x) || x = "" === x = "" and
  NOT ISBLANK(x) && x <> "" === x <> ""):
    * Unlabelled Files With SIT
    * Files With Retention Label
    * Files With Sensitivity Label
- displayFolder grouping added (model-only polish; legacy had none).
"""

from __future__ import annotations

from .tmdl_model import MeasureSpec

_T = "FactLocationSIT"


def _m(name: str, dax: str, fmt: str, folder: str) -> MeasureSpec:
    return MeasureSpec(table=_T, name=name, dax=dax, format_string=fmt,
                       display_folder=folder)


MEASURES: list[MeasureSpec] = [
    # --- Aggregates (Export-ContentExplorerData aggregate counts) -----------
    _m("Aggregate Items", "SUM ( FactLocationSIT[item_count] )", "#,0", "Aggregates"),
    _m("Locations", "DISTINCTCOUNT ( FactLocationSIT[location_id] )", "#,0", "Aggregates"),
    _m("SITs With Aggregate Hits", "DISTINCTCOUNT ( FactLocationSIT[sit_key] )", "#,0", "Aggregates"),
    # --- Files & SIT tags ----------------------------------------------------
    _m("Files With Exported SIT", "DISTINCTCOUNT ( FactFileSITTag[file_id] )", "#,0", "Files"),
    _m("File SIT Tag Rows", "COUNTROWS ( FactFileSITTag )", "#,0", "Files"),
    _m("Unique Files", "DISTINCTCOUNT ( DimFile[file_id] )", "#,0", "Files"),
    # --- Areas ----------------------------------------------------------------
    _m("Areas", "DISTINCTCOUNT ( DimArea[area_key] )", "#,0", "Areas"),
    _m("Area File SIT Pairs", "SUM ( FactAreaSIT[file_sit_pair_count] )", "#,0", "Areas"),
    _m("Area Distinct SITs", "DISTINCTCOUNT ( FactAreaSIT[sit_key] )", "#,0", "Areas"),
    _m("Area Total Matches", "SUM ( FactAreaSIT[total_match_count] )", "#,0", "Areas"),
    _m("Area High Confidence Matches", "SUM ( FactAreaSIT[high_confidence_match_count] )", "#,0", "Areas"),
    _m("Area Risk Pressure", "SUM ( FactAreaSIT[risk_weighted_file_count] )", "#,0", "Areas"),
    _m("Area Risk Density", "DIVIDE ( [Area Risk Pressure], [Files With Exported SIT] )", "0.0", "Areas"),
    _m("Area SITs Per File", "DIVIDE ( [Area File SIT Pairs], [Files With Exported SIT] )", "0.0", "Areas"),
    _m("High Or Critical Area File SIT Pairs",
       'CALCULATE ( [Area File SIT Pairs], DimSIT[risk_band] IN { "High", "Critical" } )', "#,0", "Areas"),
    _m("Critical Area File SIT Pairs",
       'CALCULATE ( [Area File SIT Pairs], DimSIT[risk_band] = "Critical" )', "#,0", "Areas"),
    _m("Unrated Area File SIT Pairs",
       "CALCULATE ( [Area File SIT Pairs], DimSIT[is_unrated] = TRUE () )", "#,0", "Areas"),
    _m("Area Sensitivity Label Coverage %",
       "DIVIDE ( SUM ( FactArea[sensitivity_labelled_files] ), SUM ( FactArea[files_with_sit] ) )",
       "0.0%", "Areas"),
    # --- Graph (node neighbourhood / adjacency) ------------------------------
    _m("Graph Nodes", "DISTINCTCOUNT ( DimGraphNode[graph_node_key] )", "#,0", "Graph"),
    _m("Connected Nodes", "DISTINCTCOUNT ( FactGraphAdjacency[connected_node_key] )", "#,0", "Graph"),
    _m("Graph Edge Weight", "SUM ( FactGraphAdjacency[edge_weight] )", "#,0.0", "Graph"),
    _m("Graph Edge Files", "SUM ( FactGraphAdjacency[file_count] )", "#,0", "Graph"),
    _m("Graph Edge SIT Pairs", "SUM ( FactGraphAdjacency[file_sit_pair_count] )", "#,0", "Graph"),
    _m("Graph Edge Risk Pressure", "SUM ( FactGraphAdjacency[risk_pressure] )", "#,0", "Graph"),
    _m("Graph Edge High Confidence Matches",
       "SUM ( FactGraphAdjacency[high_confidence_match_count] )", "#,0", "Graph"),
    # --- Force graph (focused edge subset) ------------------------------------
    _m("Force Graph Edges", "COUNTROWS ( FactGraphEdgeFocus )", "#,0", "Force Graph"),
    _m("Force Graph Edge Weight", "SUM ( FactGraphEdgeFocus[edge_weight] )", "#,0.0", "Force Graph"),
    _m("Force Graph Risk Pressure", "SUM ( FactGraphEdgeFocus[risk_pressure] )", "#,0", "Force Graph"),
    _m("Force Graph Files", "SUM ( FactGraphEdgeFocus[file_count] )", "#,0", "Force Graph"),
    _m("Force Graph High Confidence Matches",
       "SUM ( FactGraphEdgeFocus[high_confidence_match_count] )", "#,0", "Force Graph"),
    # --- Sankey flows ----------------------------------------------------------
    _m("Sankey User Location Weight", "SUM ( FactSankeyUserLocationFlow[flow_weight] )", "#,0.0", "Sankey"),
    _m("Sankey User Location Files", "SUM ( FactSankeyUserLocationFlow[file_count] )", "#,0", "Sankey"),
    _m("Sankey User Location Risk", "SUM ( FactSankeyUserLocationFlow[risk_pressure] )", "#,0", "Sankey"),
    _m("Sankey Sensitivity Weight", "SUM ( FactSankeySensitivityFlow[flow_weight] )", "#,0.0", "Sankey"),
    _m("Sankey Sensitivity Files", "SUM ( FactSankeySensitivityFlow[file_count] )", "#,0", "Sankey"),
    _m("Sankey Sensitivity Risk", "SUM ( FactSankeySensitivityFlow[risk_pressure] )", "#,0", "Sankey"),
    # --- Clusters ---------------------------------------------------------------
    _m("Graph Clusters", "DISTINCTCOUNT ( DimGraphCluster[graph_cluster_key] )", "#,0", "Clusters"),
    _m("Cluster Files With SIT", "SUM ( DimGraphCluster[files_with_sit] )", "#,0", "Clusters"),
    _m("Cluster Risk Pressure", "SUM ( DimGraphCluster[risk_pressure] )", "#,0", "Clusters"),
    _m("Cluster Risk Density", "DIVIDE ( [Cluster Risk Pressure], [Cluster Files With SIT] )", "0.0", "Clusters"),
    _m("Cluster High Confidence Matches",
       "SUM ( DimGraphCluster[high_confidence_match_count] )", "#,0", "Clusters"),
    _m("Cluster Node Size", "SUM ( DimGraphCluster[node_size] )", "#,0.0", "Clusters"),
    # --- Executive findings -------------------------------------------------------
    _m("Executive Findings", "COUNTROWS ( FactExecFinding )", "#,0", "Executive"),
    _m("Critical Executive Findings",
       'CALCULATE ( [Executive Findings], FactExecFinding[severity] = "Critical" )', "#,0", "Executive"),
    _m("High Or Critical Executive Findings",
       'CALCULATE ( [Executive Findings], FactExecFinding[severity] IN { "High", "Critical" } )',
       "#,0", "Executive"),
    _m("Executive Finding Metric", "SUM ( FactExecFinding[metric_value] )", "#,0.0", "Executive"),
    # --- Detected SIT confidence --------------------------------------------------
    _m("Detected Files By Location", "SUM ( FactDetectedSITByLocation[file_count] )", "#,0", "Confidence"),
    _m("Detected Confidence Matches",
       "SUM ( FactDetectedSITByLocation[total_confidence_count] )", "#,0", "Confidence"),
    _m("Low Confidence Matches", "SUM ( FactDetectedSITByLocation[low_confidence_count] )", "#,0", "Confidence"),
    _m("Medium Confidence Matches",
       "SUM ( FactDetectedSITByLocation[medium_confidence_count] )", "#,0", "Confidence"),
    _m("High Confidence Matches",
       "SUM ( FactDetectedSITByLocation[high_confidence_count] )", "#,0", "Confidence"),
    _m("High Confidence %", "DIVIDE ( [High Confidence Matches], [Detected Confidence Matches] )",
       "0.0%", "Confidence"),
    _m("Medium Confidence %", "DIVIDE ( [Medium Confidence Matches], [Detected Confidence Matches] )",
       "0.0%", "Confidence"),
    _m("Low Confidence %", "DIVIDE ( [Low Confidence Matches], [Detected Confidence Matches] )",
       "0.0%", "Confidence"),
    # --- Risk weighting --------------------------------------------------------------
    _m("Risk Weighted Aggregate Items",
       "SUMX ( FactLocationSIT, FactLocationSIT[item_count] * COALESCE ( RELATED ( DimSIT[risk_score] ), 0 ) )",
       "#,0", "Risk"),
    _m("Risk Weighted Detected Files",
       "SUMX ( FactDetectedSITByLocation, FactDetectedSITByLocation[file_count] * COALESCE ( RELATED ( DimSIT[risk_score] ), 0 ) )",
       "#,0", "Risk"),
    _m("Weighted Average Risk", "DIVIDE ( [Risk Weighted Aggregate Items], [Aggregate Items] )", "0.0", "Risk"),
    _m("Critical Aggregate Items",
       'CALCULATE ( [Aggregate Items], DimSIT[risk_band] = "Critical" )', "#,0", "Risk"),
    _m("High Or Critical Aggregate Items",
       'CALCULATE ( [Aggregate Items], DimSIT[risk_band] IN { "High", "Critical" } )', "#,0", "Risk"),
    _m("Critical Detected Files",
       'CALCULATE ( [Detected Files By Location], DimSIT[risk_band] = "Critical" )', "#,0", "Risk"),
    _m("High Or Critical Detected Files",
       'CALCULATE ( [Detected Files By Location], DimSIT[risk_band] IN { "High", "Critical" } )',
       "#,0", "Risk"),
    _m("Unrated Aggregate Items",
       "CALCULATE ( [Aggregate Items], DimSIT[is_unrated] = TRUE () )", "#,0", "Risk"),
    _m("Unrated SITs",
       "CALCULATE ( DISTINCTCOUNT ( DimSIT[sit_key] ), DimSIT[is_unrated] = TRUE () )", "#,0", "Risk"),
    # --- Labels (column-predicate rewrites of the legacy FILTER iterators) ------------
    _m("Unlabelled Files With SIT",
       'CALCULATE ( [Files With Exported SIT], DimFile[sensitivity_label] = "" )', "#,0", "Labels"),
    _m("Critical Unlabelled Files",
       'CALCULATE ( [Unlabelled Files With SIT], DimSIT[risk_band] = "Critical" )', "#,0", "Labels"),
    _m("Files With Retention Label",
       'CALCULATE ( [Unique Files], DimFile[retention_label] <> "" )', "#,0", "Labels"),
    _m("Files With Sensitivity Label",
       'CALCULATE ( [Unique Files], DimFile[sensitivity_label] <> "" )', "#,0", "Labels"),
    _m("Sensitivity Label Coverage %",
       "DIVIDE ( [Files With Sensitivity Label], [Unique Files] )", "0.0%", "Labels"),
    # --- Users -------------------------------------------------------------------------
    _m("Users On Files", "DISTINCTCOUNT ( BridgeFileUser[user_id] )", "#,0", "Users"),
    _m("User Files", "DISTINCTCOUNT ( BridgeFileUser[file_id] )", "#,0", "Users"),
    _m("User File SIT Tags",
       "VAR SelectedFiles = VALUES ( BridgeFileUser[file_id] ) "
       "RETURN CALCULATE ( [Files With Exported SIT], TREATAS ( SelectedFiles, FactFileSITTag[file_id] ) )",
       "#,0", "Users"),
]

# The 71 legacy measure names — the porting contract asserted by tests.
LEGACY_MEASURE_NAMES: tuple[str, ...] = tuple(measure.name for measure in MEASURES)
