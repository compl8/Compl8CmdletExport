"""Schema-integrity tests for the AE star-schema v6 SSOT."""

from __future__ import annotations

import pyarrow as pa
import pytest

from parquet_builder.star import schema

EXPECTED_TABLES = {
    # facts
    "fact_activity", "fact_activity_sit", "fact_policy_activity",
    "fact_email_recipient", "fact_activity_detail", "fact_email_detail",
    "fact_copilot_interaction",
    # dims
    "dim_app_identity", "dim_source_page", "dim_user", "dim_department",
    "dim_sit", "dim_date", "dim_location", "dim_file", "dim_domain",
    "dim_email_address", "dim_policy", "dim_workload", "dim_activity_type",
    # aggs
    "agg_department_sit_day", "agg_user_sit_day", "agg_location_sit_day",
    "agg_activity_type_sit_day", "agg_domain_sit_day",
    # index / pipeline-only
    "activity_record_index", "archive_raw",
}


def test_all_v6_tables_declared() -> None:
    assert set(schema.TABLES) == EXPECTED_TABLES


def test_schema_validates_clean() -> None:
    assert schema.validate_schema() == []


def test_no_duplicate_columns_and_keys_present() -> None:
    for table in schema.TABLES.values():
        names = table.column_names()
        assert len(names) == len(set(names)), f"duplicate columns in {table.name}"
        if table.key is not None:
            assert table.key in names


def test_relationship_endpoints_exist_and_single_direction() -> None:
    for rel in schema.RELATIONSHIPS:
        assert rel.from_table in schema.TABLES
        assert rel.to_table in schema.TABLES
        assert rel.from_column in schema.TABLES[rel.from_table].column_names()
        assert rel.to_column in schema.TABLES[rel.to_table].column_names()
        # filter direction is never Both
        assert rel.cross_filter == "single"
        # relationships always land on the target table's key
        assert schema.TABLES[rel.to_table].key == rel.to_column


def _find_rel(from_table: str, from_column: str, to_table: str):
    for rel in schema.RELATIONSHIPS:
        if (rel.from_table, rel.from_column, rel.to_table) == (from_table, from_column, to_table):
            return rel
    return None


def test_domain_and_location_relationship_activity_flags() -> None:
    assert _find_rel("fact_activity", "target_domain_id", "dim_domain").active is True
    assert _find_rel("fact_activity", "originating_domain_id", "dim_domain").active is False
    assert _find_rel("fact_activity", "target_location_id", "dim_location").active is False
    assert _find_rel("fact_activity", "source_location_id", "dim_location").active is True
    assert _find_rel("fact_activity_sit", "target_domain_id", "dim_domain").active is True


def test_fact_activity_v6_additions() -> None:
    cols = schema.TABLES["fact_activity"].column_names()
    for name in ("user_type", "data_platform", "app_identity_id"):
        assert name in cols


def test_fact_activity_sit_v6_additions() -> None:
    cols = schema.TABLES["fact_activity_sit"].column_names()
    for name in ("classifier_type", "target_domain_id", "policy_rule_id"):
        assert name in cols


def test_fact_activity_detail_contract() -> None:
    cols = schema.TABLES["fact_activity_detail"].column_names()
    # v6 drops these (index/dim_file own them)
    assert "record_identity" not in cols
    assert "file_path" not in cols
    # catch-all and typed endpoint/DLP contract columns
    for name in (
        "extra_json", "enforcement_mode", "rms_encrypted", "previous_file_name",
        "target_printer_name", "mdatp_device_id", "jit_triggered", "evidence_file",
        "removable_media_device_attributes", "endpoint_operation", "authorized_group",
        "matched_policies", "dlp_audit_event_metadata", "session_metadata",
        "item_metadata", "parent_archive_hash", "agent_id", "agent_name",
        "target_agent_id", "target_agent_name", "platform_target_agent_id",
        "file_path_url", "source_location_type", "destination_location_type",
        "sha1", "sha256", "cold_scan_policy_id", "policy_version",
        "associated_admin_units", "sensitivity_label_ids_referenced",
    ):
        assert name in cols, f"missing detail contract column {name}"


def test_dim_sit_carries_all_reference_columns() -> None:
    cols = schema.TABLES["dim_sit"].column_names()
    for name in (
        "sit_name", "sit_id", "sit_slug", "category", "risk_description",
        "risk_score", "risk_band", "reference_url", "pspf_classification",
        "qgiscf", "qgiscf_dlm", "label_code", "sit_classifier_type", "source",
        "jurisdictions", "scope", "reference_confidence", "classification_tier",
        "generic_classification", "generic_dlm", "observed", "is_unrated",
    ):
        assert name in cols, f"missing dim_sit reference column {name}"


def test_dim_date_has_month_short_and_week_of_year() -> None:
    cols = schema.TABLES["dim_date"].column_names()
    assert "month_short" in cols
    assert "week_of_year" in cols


def test_activity_record_index_has_page_provenance() -> None:
    cols = schema.TABLES["activity_record_index"].column_names()
    assert "page_id" in cols
    assert _find_rel("activity_record_index", "page_id", "dim_source_page") is not None


def test_archive_raw_is_pipeline_only() -> None:
    assert schema.TABLES["archive_raw"].kind == "pipeline_only"


def test_pyarrow_schema_resolution() -> None:
    arrow = schema.pyarrow_schema("fact_activity")
    spec = schema.TABLES["fact_activity"]
    assert arrow.names == spec.column_names()
    assert arrow.field("activity_id").type == pa.int64()
    assert not arrow.field("activity_id").nullable
    assert arrow.field("happened_at").type == pa.timestamp("us")
    assert arrow.field("has_sit").type == pa.bool_()
    assert schema.pyarrow_schema("dim_date").field("date").type == pa.date32()


def test_emit_schema_json_payload(tmp_path) -> None:
    out = tmp_path / "schema.json"
    payload = schema.emit_schema_json(out)
    assert out.exists()
    assert payload["version"] == 6
    assert payload["profile"] == "powerbi_star"
    assert {t["name"] for t in payload["tables"]} == EXPECTED_TABLES
    assert len(payload["relationships"]) == len(schema.RELATIONSHIPS)
    valid_pbi = {"Int64", "String", "DateTime", "Boolean", "Double"}
    for table in payload["tables"]:
        for col in table["columns"]:
            assert col["pbi_type"] in valid_pbi, f"{table['name']}.{col['name']}"


def test_unknown_table_raises() -> None:
    with pytest.raises(KeyError):
        schema.pyarrow_schema("not_a_table")
