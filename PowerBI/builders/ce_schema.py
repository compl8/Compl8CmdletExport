"""Content Explorer SIT Risk — declarative model spec for the EXISTING CE
parquet layout (produced by the legacy repo's convert_content_explorer_to_parquet.py
+ derive_content_explorer_area_tables.py).

The CE parquet schema is NOT changing in this phase, so unlike AE (whose model
is generated from parquet_builder.star.schema) this module mirrors the legacy
project's TMDL exactly: 22 tables, 21 single-direction relationships (all
active), the same hidden columns, and the same parquet file names. It reuses
the engine's TableSpec/ColumnSpec/RelationshipSpec types from
parquet_builder.star.spec_types (they fit as-is; dtype tokens map 1:1 onto the
legacy M conversion types, incl. the tz-aware `timestamp_tz` for
DimClassifier.export_date) and feeds tmdl_model via a ModelSource — so a
future CE star schema can swap in by replacing this module alone.
"""

from __future__ import annotations

from parquet_builder.star.spec_types import ColumnSpec, RelationshipSpec, TableSpec

from .tmdl_model import ModelSource

# Default ParquetRoot parameter value — placeholder, never a machine path.
# Point it at the CE converter's --output-dir.
CE_DEFAULT_PARQUET_ROOT = "C:\\CHANGEME\\PowerBI-CE-Parquet"


def _c(name: str, dtype: str = "string", fmt: str | None = None) -> ColumnSpec:
    return ColumnSpec(name=name, dtype=dtype, format_string=fmt)


def _t(name: str, kind: str, key: str | None, *columns: ColumnSpec) -> TableSpec:
    return TableSpec(name=name, kind=kind, key=key, columns=tuple(columns))


_SHORT_DATE = "Short Date"

# Tables in the legacy model's PBI_QueryOrder (load order preserved).
CE_TABLES: tuple[TableSpec, ...] = (
    _t("DimSIT", "dim", "sit_key",
       _c("sit_key"), _c("sit_name"), _c("sit_name_norm"), _c("sit_id"),
       _c("sit_slug"), _c("category"), _c("risk_score", "int64"),
       _c("risk_band"), _c("risk_description"), _c("reference_url"),
       _c("pspf_classification"), _c("qgiscf"), _c("qgiscf_dlm"),
       _c("small_tenant", "bool"), _c("medium_tenant", "bool"),
       _c("large_tenant", "bool"), _c("scope"), _c("confidence"),
       _c("data_categories"), _c("regulations"), _c("source_sheet"),
       _c("is_unrated", "bool")),
    _t("DimQGISCFDLM", "dim", None,
       _c("label_name"), _c("label_order", "int64"), _c("visual_marking"),
       _c("encrypt"), _c("description")),
    _t("DimLocation", "dim", "location_id",
       _c("location_id"), _c("location_url"), _c("location_type"),
       _c("workload"), _c("workload_code"), _c("host"), _c("site_path"),
       _c("location_name"), _c("owner_slug"), _c("owner_candidate")),
    _t("DimArea", "dim", "area_key",
       _c("area_key"), _c("location_id"), _c("workload"), _c("workload_code"),
       _c("location_type"), _c("location_name"), _c("folder_path"),
       _c("folder_name"), _c("folder_depth", "int64"),
       _c("area_display_name"), _c("area_display_path"),
       _c("is_deep_folder", "bool"), _c("area_level_1"), _c("area_level_2"),
       _c("area_level_3"), _c("area_level_4"), _c("area_level_5"),
       _c("area_level_6")),
    _t("DimFile", "dim", "file_id",
       _c("file_id"), _c("file_url"), _c("file_name"), _c("file_extension"),
       _c("file_source_url"), _c("location_id"), _c("workload"),
       _c("workload_code"), _c("relative_path"), _c("path_depth", "int64"),
       _c("leaf_area_key"), _c("folder_path"), _c("folder_name"),
       _c("folder_depth", "int64"), _c("area_level_1"), _c("area_level_2"),
       _c("area_level_3"), _c("area_level_4"), _c("area_level_5"),
       _c("area_level_6"), _c("sensitivity_label"), _c("retention_label"),
       _c("trainable_classifiers"), _c("created_by"), _c("modified_by"),
       _c("last_modified_at", "timestamp_us", _SHORT_DATE),
       _c("last_modified_date", "date32", _SHORT_DATE)),
    _t("DimUser", "dim", "user_id",
       _c("user_id"), _c("user_display_name"), _c("normalized_display_name")),
    _t("BridgeFileUser", "fact", None,
       _c("file_id"), _c("user_id"), _c("user_display_name"), _c("user_role")),
    _t("DimClassifier", "dim", "classifier_key",
       _c("classifier_key"), _c("tag_type"), _c("sit_key"), _c("sit_name"),
       _c("total_records", "int64"), _c("total_pages", "int64"),
       _c("export_date", "timestamp_tz", _SHORT_DATE),
       _c("onedrive_records", "int64"), _c("onedrive_pages", "int64"),
       _c("onedrive_status"), _c("sharepoint_records", "int64"),
       _c("sharepoint_pages", "int64"), _c("sharepoint_status")),
    _t("FactLocationSIT", "fact", "fact_id",
       _c("fact_id"), _c("exported_at", "timestamp_us", _SHORT_DATE),
       _c("tag_type"), _c("sit_key"), _c("sit_name"), _c("workload"),
       _c("workload_code"), _c("location_id"), _c("location_url"),
       _c("item_count", "int64"), _c("error")),
    _t("FactFileSITTag", "fact", "hit_id",
       _c("hit_id"), _c("file_id"), _c("leaf_area_key"), _c("location_id"),
       _c("tag_type"), _c("sit_key"), _c("sit_name"), _c("workload"),
       _c("workload_code"), _c("last_modified_date", "date32", _SHORT_DATE),
       _c("tag_low_confidence_count", "int64"),
       _c("tag_medium_confidence_count", "int64"),
       _c("tag_high_confidence_count", "int64"),
       _c("tag_total_count", "int64"), _c("has_tag_confidence", "bool"),
       _c("observed_count", "int64"), _c("sensitivity_label"),
       _c("retention_label")),
    _t("FactArea", "fact", None,
       _c("area_key"), _c("files_with_sit", "int64"),
       _c("unlabelled_files_with_sit", "int64"),
       _c("sensitivity_labelled_files", "int64"),
       _c("file_sit_pair_count", "int64"), _c("distinct_sit_count", "int64"),
       _c("critical_sit_count", "int64"),
       _c("high_or_critical_sit_count", "int64"),
       _c("unrated_sit_count", "int64"),
       _c("critical_file_sit_pairs", "int64"),
       _c("high_or_critical_file_sit_pairs", "int64"),
       _c("total_match_count", "int64"),
       _c("high_confidence_match_count", "int64"),
       _c("risk_weighted_file_sit_count", "int64"),
       _c("max_risk_score", "int64")),
    _t("FactAreaSIT", "fact", None,
       _c("area_key"), _c("sit_key"), _c("sit_name"), _c("file_count", "int64"),
       _c("file_sit_pair_count", "int64"), _c("total_match_count", "int64"),
       _c("low_confidence_match_count", "int64"),
       _c("medium_confidence_match_count", "int64"),
       _c("high_confidence_match_count", "int64"),
       _c("observed_count", "int64"), _c("risk_score", "int64"),
       _c("risk_band"), _c("risk_weighted_file_count", "int64"),
       _c("is_unrated", "bool")),
    _t("FactDetectedSITByLocation", "fact", None,
       _c("location_id"), _c("sit_key"), _c("detected_sit_id"), _c("workload"),
       _c("workload_code"), _c("file_count", "int64"),
       _c("low_confidence_count", "int64"),
       _c("medium_confidence_count", "int64"),
       _c("high_confidence_count", "int64"),
       _c("total_confidence_count", "int64")),
    _t("DimGraphCluster", "dim", "graph_cluster_key",
       _c("graph_cluster_key"), _c("cluster_label"), _c("workload"),
       _c("location_id"), _c("location_name"), _c("area_level_1"),
       _c("area_count", "int64"), _c("files_with_sit", "int64"),
       _c("file_sit_pair_count", "int64"), _c("distinct_sit_count", "int64"),
       _c("risk_pressure", "int64"), _c("risk_density", "double"),
       _c("high_confidence_match_count", "int64"),
       _c("label_coverage_pct", "double"), _c("cluster_rank", "int64"),
       _c("layout_x", "double"), _c("layout_y", "double"),
       _c("node_size", "double")),
    _t("DimGraphNode", "dim", "graph_node_key",
       _c("graph_node_key"), _c("node_type"), _c("node_label"),
       _c("node_subtype"), _c("location_id"), _c("area_key"), _c("sit_key"),
       _c("user_id"), _c("graph_cluster_key"), _c("workload"),
       _c("location_name"), _c("folder_path"), _c("risk_pressure", "int64"),
       _c("files_with_sit", "int64"), _c("file_sit_pair_count", "int64"),
       _c("distinct_sit_count", "int64"),
       _c("high_confidence_match_count", "int64"), _c("node_size", "double"),
       _c("layout_x", "double"), _c("layout_y", "double"),
       _c("is_high_value", "bool"), _c("node_note")),
    _t("DimUserDepartment", "dim", "user_id",
       _c("user_id"), _c("user_display_name"), _c("normalized_display_name"),
       _c("department"), _c("division"), _c("business_unit"),
       _c("mapping_source"), _c("is_mapped", "bool")),
    _t("FactGraphEdge", "fact", "edge_key",
       _c("edge_key"), _c("source_node_key"), _c("target_node_key"),
       _c("source_node_type"), _c("target_node_type"), _c("source_node_label"),
       _c("target_node_label"), _c("edge_type"), _c("edge_label"),
       _c("graph_cluster_key"), _c("file_count", "int64"),
       _c("file_sit_pair_count", "int64"), _c("distinct_sit_count", "int64"),
       _c("risk_pressure", "int64"),
       _c("high_confidence_match_count", "int64"),
       _c("total_match_count", "int64"), _c("edge_weight", "double")),
    _t("FactGraphEdgeFocus", "fact", "edge_key",
       _c("edge_key"), _c("focus_rank", "int64"), _c("focus_bucket"),
       _c("source_node_key"), _c("target_node_key"), _c("source_node_type"),
       _c("target_node_type"), _c("source_node_label"),
       _c("target_node_label"), _c("source_visual_label"),
       _c("target_visual_label"), _c("edge_type"), _c("edge_label"),
       _c("graph_cluster_key"), _c("file_count", "int64"),
       _c("file_sit_pair_count", "int64"), _c("distinct_sit_count", "int64"),
       _c("risk_pressure", "int64"),
       _c("high_confidence_match_count", "int64"),
       _c("total_match_count", "int64"), _c("edge_weight", "double")),
    _t("FactSankeyUserLocationFlow", "fact", "flow_key",
       _c("flow_key"), _c("flow_stage", "int64"), _c("source_node"),
       _c("target_node"), _c("source_type"), _c("target_type"),
       _c("department"), _c("user_id"), _c("user_display_name"),
       _c("location_id"), _c("location_name"), _c("sit_key"), _c("sit_name"),
       _c("risk_band"), _c("risk_score", "int64"), _c("qgiscf_dlm"),
       _c("pspf_classification"), _c("category"), _c("file_count", "int64"),
       _c("file_sit_pair_count", "int64"), _c("risk_pressure", "int64"),
       _c("high_confidence_match_count", "int64"),
       _c("flow_weight", "double"), _c("flow_rank", "int64")),
    _t("FactSankeySensitivityFlow", "fact", "flow_key",
       _c("flow_key"), _c("flow_stage", "int64"), _c("source_node"),
       _c("target_node"), _c("source_type"), _c("target_type"),
       _c("location_id"), _c("location_name"), _c("sit_key"), _c("sit_name"),
       _c("risk_band"), _c("risk_score", "int64"), _c("qgiscf_dlm"),
       _c("pspf_classification"), _c("category"), _c("file_count", "int64"),
       _c("file_sit_pair_count", "int64"), _c("risk_pressure", "int64"),
       _c("high_confidence_match_count", "int64"),
       _c("flow_weight", "double"), _c("flow_rank", "int64")),
    _t("FactGraphAdjacency", "fact", "adjacency_key",
       _c("adjacency_key"), _c("selected_node_key"), _c("connected_node_key"),
       _c("selected_node_type"), _c("connected_node_type"),
       _c("selected_node_label"), _c("connected_node_label"),
       _c("selected_visual_label"), _c("connected_visual_label"),
       _c("relationship_type"), _c("graph_cluster_key"),
       _c("file_count", "int64"), _c("file_sit_pair_count", "int64"),
       _c("distinct_sit_count", "int64"), _c("risk_pressure", "int64"),
       _c("high_confidence_match_count", "int64"),
       _c("total_match_count", "int64"), _c("edge_weight", "double")),
    _t("FactExecFinding", "fact", "finding_id",
       _c("finding_id"), _c("finding_rank", "int64"), _c("severity"),
       _c("finding_type"), _c("entity_type"), _c("entity_key"),
       _c("entity_label"), _c("area_key"), _c("location_id"), _c("sit_key"),
       _c("metric_name"), _c("metric_value", "double"),
       _c("secondary_metric_name"), _c("secondary_metric_value", "double"),
       _c("finding_title"), _c("finding_evidence"), _c("recommended_action")),
)

# 21 relationships — all active, single-direction (mirrors the legacy model).
CE_RELATIONSHIPS: tuple[RelationshipSpec, ...] = (
    RelationshipSpec("FactLocationSIT", "sit_key", "DimSIT", "sit_key"),
    RelationshipSpec("FactFileSITTag", "sit_key", "DimSIT", "sit_key"),
    RelationshipSpec("FactDetectedSITByLocation", "sit_key", "DimSIT", "sit_key"),
    RelationshipSpec("FactAreaSIT", "sit_key", "DimSIT", "sit_key"),
    RelationshipSpec("DimClassifier", "sit_key", "DimSIT", "sit_key"),
    RelationshipSpec("FactLocationSIT", "location_id", "DimLocation", "location_id"),
    RelationshipSpec("FactDetectedSITByLocation", "location_id", "DimLocation", "location_id"),
    RelationshipSpec("DimFile", "location_id", "DimLocation", "location_id"),
    RelationshipSpec("FactFileSITTag", "leaf_area_key", "DimArea", "area_key"),
    RelationshipSpec("FactArea", "area_key", "DimArea", "area_key"),
    RelationshipSpec("FactAreaSIT", "area_key", "DimArea", "area_key"),
    RelationshipSpec("FactFileSITTag", "file_id", "DimFile", "file_id"),
    RelationshipSpec("BridgeFileUser", "file_id", "DimFile", "file_id"),
    RelationshipSpec("BridgeFileUser", "user_id", "DimUser", "user_id"),
    RelationshipSpec("FactGraphAdjacency", "selected_node_key", "DimGraphNode", "graph_node_key"),
    RelationshipSpec("FactGraphEdgeFocus", "graph_cluster_key", "DimGraphCluster", "graph_cluster_key"),
    RelationshipSpec("FactSankeyUserLocationFlow", "user_id", "DimUserDepartment", "user_id"),
    RelationshipSpec("FactSankeyUserLocationFlow", "location_id", "DimLocation", "location_id"),
    RelationshipSpec("FactSankeyUserLocationFlow", "sit_key", "DimSIT", "sit_key"),
    RelationshipSpec("FactSankeySensitivityFlow", "location_id", "DimLocation", "location_id"),
    RelationshipSpec("FactSankeySensitivityFlow", "sit_key", "DimSIT", "sit_key"),
)

# Table name -> parquet file (legacy CE converter output names).
CE_PARQUET_FILES: dict[str, str] = {
    "DimSIT": "dim_sit.parquet",
    "DimQGISCFDLM": "dim_qgiscf_dlm.parquet",
    "DimLocation": "dim_location.parquet",
    "DimArea": "dim_area.parquet",
    "DimFile": "dim_file.parquet",
    "DimUser": "dim_user.parquet",
    "BridgeFileUser": "bridge_file_user.parquet",
    "DimClassifier": "dim_classifier.parquet",
    "FactLocationSIT": "fact_location_sit_aggregate.parquet",
    "FactFileSITTag": "fact_file_sit_tag.parquet",
    "FactArea": "fact_area.parquet",
    "FactAreaSIT": "fact_area_sit.parquet",
    "FactDetectedSITByLocation": "fact_detected_sit_by_location.parquet",
    "DimGraphCluster": "dim_graph_cluster.parquet",
    "DimGraphNode": "dim_graph_node.parquet",
    "DimUserDepartment": "dim_user_department.parquet",
    "FactGraphEdge": "fact_graph_edge.parquet",
    "FactGraphEdgeFocus": "fact_graph_edge_focus.parquet",
    "FactSankeyUserLocationFlow": "fact_sankey_user_location_flow.parquet",
    "FactSankeySensitivityFlow": "fact_sankey_sensitivity_flow.parquet",
    "FactGraphAdjacency": "fact_graph_adjacency.parquet",
    "FactExecFinding": "fact_exec_finding.parquet",
}

# Hidden columns beyond keys + relationship endpoints (legacy model parity:
# normalisation helpers, snowflake/grain keys not modelled as relationships).
CE_EXTRA_HIDDEN: dict[str, tuple[str, ...]] = {
    "DimSIT": ("sit_name_norm",),
    "DimUser": ("normalized_display_name",),
    "DimUserDepartment": ("normalized_display_name",),
    "DimFile": ("leaf_area_key",),
    "DimArea": ("location_id",),
    "DimGraphCluster": ("location_id",),
    "DimGraphNode": ("location_id", "area_key", "sit_key", "user_id",
                     "graph_cluster_key"),
    "FactFileSITTag": ("location_id",),
    "FactDetectedSITByLocation": ("detected_sit_id",),
    "FactGraphEdge": ("source_node_key", "target_node_key", "graph_cluster_key"),
    "FactGraphEdgeFocus": ("source_node_key", "target_node_key"),
    "FactGraphAdjacency": ("connected_node_key", "graph_cluster_key"),
    "FactExecFinding": ("area_key", "location_id", "sit_key"),
}


def ce_model_source() -> ModelSource:
    """The Content Explorer model source (no date table in the legacy model)."""
    return ModelSource(
        tables=CE_TABLES,
        relationships=CE_RELATIONSHIPS,
        parquet_files=CE_PARQUET_FILES,
        extra_hidden=CE_EXTRA_HIDDEN,
    )
