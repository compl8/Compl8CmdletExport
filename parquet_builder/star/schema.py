"""Single source of truth for the Activity Explorer star schema (v6).

Every table, column, type, key, and relationship in the Power BI-facing
parquet model is declared across `spec_types` / `dimensions` / `facts` and
assembled here. The parquet converter (star.convert) builds its pyarrow
schemas from this module, and the Power BI TMDL model builder generates the
semantic model from the same declarations, so the data and the report model
can never drift apart.

PBI metadata carried per column: pbi_type (Int64/String/DateTime/Boolean/
Double), optional format_string, summarize_by, description.
"""

from __future__ import annotations

import json
from pathlib import Path

from .dimensions import DIM_TABLES
from .facts import ACTIVITY_RECORD_INDEX, AGG_TABLES, ARCHIVE_RAW, FACT_TABLES
from .spec_types import (
    PBI_TYPE_BY_DTYPE,
    SCHEMA_PROFILE,
    SCHEMA_VERSION,
    VALID_KINDS,
    VALID_SUMMARIZE_BY,
    ColumnSpec,
    RelationshipSpec,
    TableSpec,
)

__all__ = [
    "SCHEMA_PROFILE",
    "SCHEMA_VERSION",
    "ColumnSpec",
    "RelationshipSpec",
    "TableSpec",
    "RELATIONSHIPS",
    "TABLES",
    "emit_schema_json",
    "model_relationships",
    "model_tables",
    "pyarrow_schema",
    "validate_schema",
]

# Table kinds that are NOT loaded into the Power BI semantic model.
NON_MODEL_KINDS = frozenset({"pipeline_only", "index"})

_ALL_TABLES = (*DIM_TABLES, *FACT_TABLES, *AGG_TABLES, ACTIVITY_RECORD_INDEX, ARCHIVE_RAW)

TABLES: dict[str, TableSpec] = {table.name: table for table in _ALL_TABLES}


def _rel(from_table: str, from_column: str, to_table: str, to_column: str,
         active: bool = True) -> RelationshipSpec:
    return RelationshipSpec(from_table, from_column, to_table, to_column, active)


RELATIONSHIPS: tuple[RelationshipSpec, ...] = (
    # fact_activity
    _rel("fact_activity", "date_key", "dim_date", "date_key"),
    _rel("fact_activity", "user_id", "dim_user", "user_id"),
    _rel("fact_activity", "department_id", "dim_department", "department_id"),
    _rel("fact_activity", "activity_type_id", "dim_activity_type", "activity_type_id"),
    _rel("fact_activity", "workload_id", "dim_workload", "workload_id"),
    _rel("fact_activity", "source_location_id", "dim_location", "location_id"),
    _rel("fact_activity", "target_location_id", "dim_location", "location_id", active=False),
    _rel("fact_activity", "file_id", "dim_file", "file_id"),
    _rel("fact_activity", "policy_rule_id", "dim_policy", "policy_rule_id"),
    _rel("fact_activity", "target_domain_id", "dim_domain", "domain_id"),
    _rel("fact_activity", "originating_domain_id", "dim_domain", "domain_id", active=False),
    _rel("fact_activity", "app_identity_id", "dim_app_identity", "app_identity_id"),
    # fact_activity_sit
    _rel("fact_activity_sit", "date_key", "dim_date", "date_key"),
    _rel("fact_activity_sit", "user_id", "dim_user", "user_id"),
    _rel("fact_activity_sit", "department_id", "dim_department", "department_id"),
    _rel("fact_activity_sit", "activity_type_id", "dim_activity_type", "activity_type_id"),
    _rel("fact_activity_sit", "workload_id", "dim_workload", "workload_id"),
    _rel("fact_activity_sit", "source_location_id", "dim_location", "location_id"),
    _rel("fact_activity_sit", "target_location_id", "dim_location", "location_id", active=False),
    _rel("fact_activity_sit", "file_id", "dim_file", "file_id"),
    _rel("fact_activity_sit", "target_domain_id", "dim_domain", "domain_id"),
    _rel("fact_activity_sit", "policy_rule_id", "dim_policy", "policy_rule_id"),
    _rel("fact_activity_sit", "sit_key", "dim_sit", "sit_key"),
    # Email-detail rollup: SIT-grain visuals/measures respond to
    # fact_email_detail column groupings (e.g. subject word cloud sized by
    # [Activities by SIT]). M:1 holds: activity_id is fact_email_detail's key;
    # SIT rows for non-email activities resolve to the blank member.
    # NOTE: fact_email_detail.date_key -> dim_date must stay INACTIVE while
    # this is active, or dim_date would have two active paths into
    # fact_activity_sit (direct vs via fact_email_detail) and Desktop would
    # reject the model as ambiguous.
    # NOT declared (same ambiguity, via every shared dim): rollups from
    # fact_activity_sit.activity_id to fact_activity or fact_activity_detail —
    # the denormalized FK relationships above already provide the dim paths,
    # so SIT-grain visuals bind dim columns (dim_file.file_name, dim_date.date)
    # instead of fact_activity / fact_activity_detail columns.
    _rel("fact_activity_sit", "activity_id", "fact_email_detail", "activity_id"),
    # fact_policy_activity
    _rel("fact_policy_activity", "date_key", "dim_date", "date_key"),
    _rel("fact_policy_activity", "user_id", "dim_user", "user_id"),
    _rel("fact_policy_activity", "department_id", "dim_department", "department_id"),
    _rel("fact_policy_activity", "policy_rule_id", "dim_policy", "policy_rule_id"),
    _rel("fact_policy_activity", "activity_type_id", "dim_activity_type", "activity_type_id"),
    _rel("fact_policy_activity", "workload_id", "dim_workload", "workload_id"),
    # fact_email_recipient
    _rel("fact_email_recipient", "date_key", "dim_date", "date_key"),
    _rel("fact_email_recipient", "recipient_email_address_id", "dim_email_address", "email_address_id"),
    _rel("fact_email_recipient", "sender_email_address_id", "dim_email_address", "email_address_id", active=False),
    _rel("fact_email_recipient", "recipient_domain_id", "dim_domain", "domain_id"),
    # fact_email_detail (date relationship inactive — see the email-detail
    # rollup note above; date filtering reaches SIT measures via dim_date ->
    # fact_activity_sit directly)
    _rel("fact_email_detail", "date_key", "dim_date", "date_key", active=False),
    _rel("fact_email_detail", "sender_email_address_id", "dim_email_address", "email_address_id"),
    # fact_copilot_interaction
    _rel("fact_copilot_interaction", "date_key", "dim_date", "date_key"),
    _rel("fact_copilot_interaction", "user_id", "dim_user", "user_id"),
    _rel("fact_copilot_interaction", "app_identity_id", "dim_app_identity", "app_identity_id"),
    # fact_activity_detail (1:1 drillthrough)
    _rel("fact_activity_detail", "activity_id", "fact_activity", "activity_id"),
    # aggregates
    _rel("agg_department_sit_day", "date_key", "dim_date", "date_key"),
    _rel("agg_department_sit_day", "department_id", "dim_department", "department_id"),
    _rel("agg_department_sit_day", "sit_key", "dim_sit", "sit_key"),
    _rel("agg_department_sit_day", "workload_id", "dim_workload", "workload_id"),
    _rel("agg_department_sit_day", "activity_type_id", "dim_activity_type", "activity_type_id"),
    _rel("agg_user_sit_day", "date_key", "dim_date", "date_key"),
    _rel("agg_user_sit_day", "user_id", "dim_user", "user_id"),
    _rel("agg_user_sit_day", "department_id", "dim_department", "department_id"),
    _rel("agg_user_sit_day", "sit_key", "dim_sit", "sit_key"),
    _rel("agg_user_sit_day", "workload_id", "dim_workload", "workload_id"),
    _rel("agg_user_sit_day", "activity_type_id", "dim_activity_type", "activity_type_id"),
    _rel("agg_location_sit_day", "date_key", "dim_date", "date_key"),
    _rel("agg_location_sit_day", "source_location_id", "dim_location", "location_id"),
    _rel("agg_location_sit_day", "sit_key", "dim_sit", "sit_key"),
    _rel("agg_location_sit_day", "workload_id", "dim_workload", "workload_id"),
    _rel("agg_location_sit_day", "activity_type_id", "dim_activity_type", "activity_type_id"),
    _rel("agg_activity_type_sit_day", "date_key", "dim_date", "date_key"),
    _rel("agg_activity_type_sit_day", "activity_type_id", "dim_activity_type", "activity_type_id"),
    _rel("agg_activity_type_sit_day", "sit_key", "dim_sit", "sit_key"),
    _rel("agg_activity_type_sit_day", "workload_id", "dim_workload", "workload_id"),
    _rel("agg_domain_sit_day", "date_key", "dim_date", "date_key"),
    _rel("agg_domain_sit_day", "target_domain_id", "dim_domain", "domain_id"),
    _rel("agg_domain_sit_day", "sit_key", "dim_sit", "sit_key"),
    _rel("agg_domain_sit_day", "workload_id", "dim_workload", "workload_id"),
    _rel("agg_domain_sit_day", "activity_type_id", "dim_activity_type", "activity_type_id"),
    # index provenance
    _rel("activity_record_index", "page_id", "dim_source_page", "page_id"),
)


def model_tables() -> list[TableSpec]:
    """Tables loaded into the Power BI semantic model (skips pipeline_only/index)."""
    return [table for table in TABLES.values() if table.kind not in NON_MODEL_KINDS]


def model_relationships() -> list[RelationshipSpec]:
    """Relationships whose both endpoints are loaded into the semantic model."""
    loaded = {table.name for table in model_tables()}
    return [
        rel for rel in RELATIONSHIPS
        if rel.from_table in loaded and rel.to_table in loaded
    ]


def pyarrow_schema(table_name: str):
    """Resolve a TableSpec to a concrete pyarrow Schema."""
    import pyarrow as pa

    dtype_map = {
        "int64": pa.int64(),
        "string": pa.string(),
        "bool": pa.bool_(),
        "timestamp_us": pa.timestamp("us"),
        "date32": pa.date32(),
        "double": pa.float64(),
    }
    table = TABLES[table_name]
    return pa.schema([
        pa.field(col.name, dtype_map[col.dtype], nullable=col.nullable)
        for col in table.columns
    ])


def validate_schema() -> list[str]:
    """Structural integrity checks; returns a list of problems (empty = OK)."""
    problems: list[str] = []
    for table in TABLES.values():
        if table.kind not in VALID_KINDS:
            problems.append(f"{table.name}: invalid kind {table.kind!r}")
        names = table.column_names()
        dupes = {n for n in names if names.count(n) > 1}
        if dupes:
            problems.append(f"{table.name}: duplicate columns {sorted(dupes)}")
        if table.key is not None and table.key not in names:
            problems.append(f"{table.name}: key column {table.key!r} missing")
        for col in table.columns:
            if col.dtype not in PBI_TYPE_BY_DTYPE:
                problems.append(f"{table.name}.{col.name}: unknown dtype {col.dtype!r}")
            if col.summarize_by not in VALID_SUMMARIZE_BY:
                problems.append(f"{table.name}.{col.name}: invalid summarize_by {col.summarize_by!r}")
    for rel in RELATIONSHIPS:
        for side, table_name, column_name in (
            ("from", rel.from_table, rel.from_column),
            ("to", rel.to_table, rel.to_column),
        ):
            table = TABLES.get(table_name)
            if table is None:
                problems.append(f"relationship {side} table missing: {table_name}")
                continue
            if column_name not in table.column_names():
                problems.append(
                    f"relationship {side} column missing: {table_name}.{column_name}"
                )
        if rel.cross_filter != "single":
            problems.append(
                f"relationship {rel.from_table}.{rel.from_column}: cross_filter must be single"
            )
        to_table = TABLES.get(rel.to_table)
        if to_table is not None and to_table.key != rel.to_column:
            problems.append(
                f"relationship target {rel.to_table}.{rel.to_column} is not the table key"
            )
    return problems


def emit_schema_json(output_path: Path | None = None) -> dict:
    """Machine-readable schema for Python/MCP consumers (and provenance)."""
    payload = {
        "version": SCHEMA_VERSION,
        "profile": SCHEMA_PROFILE,
        "tables": [
            {
                "name": table.name,
                "kind": table.kind,
                "key": table.key,
                "description": table.description,
                "columns": [
                    {
                        "name": col.name,
                        "dtype": col.dtype,
                        "nullable": col.nullable,
                        "pbi_type": col.resolved_pbi_type(),
                        "format_string": col.format_string,
                        "summarize_by": col.summarize_by,
                        "description": col.description,
                    }
                    for col in table.columns
                ],
            }
            for table in TABLES.values()
        ],
        "relationships": [
            {
                "from_table": rel.from_table,
                "from_column": rel.from_column,
                "to_table": rel.to_table,
                "to_column": rel.to_column,
                "active": rel.active,
                "cross_filter": rel.cross_filter,
            }
            for rel in RELATIONSHIPS
        ],
    }
    if output_path is not None:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    return payload
