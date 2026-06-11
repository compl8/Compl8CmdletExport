"""Query-expression and deterministic-id primitives for Power BI report JSON.

These helpers emit the `prototypeQuery` / filter expression shapes used by the
Power BI Desktop Layout format. The shapes mirror the hand-built reference
report (see tests/fixtures/*.json for golden examples).
"""

from __future__ import annotations

import uuid
from dataclasses import dataclass

_ID_NAMESPACE = "compl8-pbi"


def stable_uuid(seed: str) -> uuid.UUID:
    """Deterministic UUID for a fully-qualified seed (project/kind/name)."""
    return uuid.uuid5(uuid.NAMESPACE_URL, f"{_ID_NAMESPACE}/{seed}")


def hex_id(seed: str) -> str:
    """20-char hex id (the Layout format's visual/section/filter name style)."""
    return stable_uuid(seed).hex[:20]


def numeric_id(seed: str) -> int:
    """Deterministic small integer id for visualContainer.json."""
    return stable_uuid(seed).int % 10_000_000


@dataclass(frozen=True)
class Field:
    """A column or measure reference used by visuals and filters."""

    kind: str  # "column" | "measure"
    table: str
    name: str
    display_name: str | None = None

    def query_ref(self) -> str:
        return f"{self.table}.{self.name}"

    def shown_as(self) -> str:
        return self.display_name or self.name


def col(table: str, name: str, display_name: str | None = None) -> Field:
    return Field("column", table, name, display_name)


def meas(table: str, name: str, display_name: str | None = None) -> Field:
    return Field("measure", table, name, display_name)


def powerbi_literal(value: object) -> str:
    """Encode a python value as a Power BI query Literal value string."""
    if value is None:
        return "null"
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, (int, float)):
        return f"{value}D"
    return "'" + str(value).replace("'", "''") + "'"


def literal_expr(value: object) -> dict[str, object]:
    return {"expr": {"Literal": {"Value": powerbi_literal(value)}}}


def solid_color_expr(color: str) -> dict[str, object]:
    return {"solid": {"color": {"expr": {"Literal": {"Value": f"'{color}'"}}}}}


def entity_column_expression(table: str, column: str) -> dict[str, object]:
    """Filter `expression` shape: Column over an Entity SourceRef."""
    return {
        "Column": {
            "Expression": {"SourceRef": {"Entity": table}},
            "Property": column,
        }
    }


def entity_measure_expression(table: str, name: str) -> dict[str, object]:
    return {
        "Measure": {
            "Expression": {"SourceRef": {"Entity": table}},
            "Property": name,
        }
    }


def query_aliases(tables: list[str]) -> dict[str, str]:
    """Deterministic short aliases for the From clause of a prototypeQuery."""
    aliases: dict[str, str] = {}
    used: set[str] = set()
    for table in tables:
        candidates = [ch for ch in table.lower() if ch.isalpha()]
        alias = next((ch for ch in candidates if ch not in used), None)
        if alias is None:
            index = len(used)
            alias = f"t{index}"
            while alias in used:
                index += 1
                alias = f"t{index}"
        aliases[table] = alias
        used.add(alias)
    return aliases


def field_expression(field: Field, aliases: dict[str, str]) -> dict[str, object]:
    source_ref = {"SourceRef": {"Source": aliases[field.table]}}
    if field.kind == "measure":
        return {"Measure": {"Expression": source_ref, "Property": field.name}}
    return {"Column": {"Expression": source_ref, "Property": field.name}}


def field_select(field: Field, aliases: dict[str, str]) -> dict[str, object]:
    selected = field_expression(field, aliases)
    selected["Name"] = field.query_ref()
    return selected


def prototype_query(
    fields: list[Field],
    order_by: Field | None = None,
    order_direction: int = 2,
) -> dict[str, object]:
    """Build the prototypeQuery for a visual (Version 2 query shape)."""
    tables: list[str] = []
    for field in [*fields, *([order_by] if order_by else [])]:
        if field.table not in tables:
            tables.append(field.table)
    aliases = query_aliases(tables)
    query: dict[str, object] = {
        "Version": 2,
        "From": [
            {"Name": aliases[table], "Entity": table, "Type": 0}
            for table in tables
        ],
        "Select": [field_select(field, aliases) for field in fields],
    }
    if order_by:
        query["OrderBy"] = [
            {"Direction": order_direction, "Expression": field_expression(order_by, aliases)}
        ]
    return query


def column_properties(fields: list[Field]) -> dict[str, object]:
    """Per-field displayName overrides keyed by queryRef."""
    props: dict[str, object] = {}
    for field in fields:
        if field.display_name and field.display_name != field.name:
            props[field.query_ref()] = {"displayName": field.display_name}
    return props
