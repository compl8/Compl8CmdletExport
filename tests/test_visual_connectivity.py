"""Visual/model connectivity guard (the "Can't determine relationships
between the fields" bug class, found on 540_Location_Drill in T6).

Power BI can only render a multi-table visual when the engine can produce an
implicit join: some single "center" table must reach every bound COLUMN table
via a chain of ACTIVE many-to-one relationship hops (the center's grain
carries the join, RELATED-style). Two tables connected only through a fact's
many-side (dim <- fact -> dim) cannot be joined as bare columns; a measure can
mediate a multi-dim grouping (SUMMARIZECOLUMNS keeps non-blank combinations),
but a bare column table left dangling from every candidate center either
errors (no measure) or cross-joins (measure present).

Rule enforced per visual, over the project's own model relationships:
- column tables (including a column order_by) must all be reachable from one
  center via active M:1 hops;
- with no measure bound, the center must itself be one of the bound column
  tables (nothing else appears in the query to join through);
- with a measure bound, any model table may act as the center (the measure's
  expansion supplies it).

The companion schema test (test_star_schema) asserts the active relationship
graph itself is unambiguous, so a center that reaches a table does so along
exactly one path.
"""

from __future__ import annotations

from parquet_builder.star import schema
from parquet_builder.star.spec_types import RelationshipSpec
from PowerBI.builders.build_activity_explorer import ae_pages
from PowerBI.builders.build_content_explorer import ce_pages
from PowerBI.builders.build_smoke import smoke_project
from PowerBI.builders.ce_schema import CE_RELATIONSHIPS, CE_TABLES
from PowerBI.builders.report_layout import PageSpec
from PowerBI.builders.visual_factories import VisualSpec


def _m2o_adjacency(relationships: tuple[RelationshipSpec, ...] | list[RelationshipSpec],
                   ) -> dict[str, set[str]]:
    """from_table -> {to_table} edges for ACTIVE relationships (many -> one)."""
    adjacency: dict[str, set[str]] = {}
    for rel in relationships:
        if rel.active:
            adjacency.setdefault(rel.from_table, set()).add(rel.to_table)
    return adjacency


def _reachable(start: str, adjacency: dict[str, set[str]]) -> set[str]:
    seen = {start}
    stack = [start]
    while stack:
        for nxt in adjacency.get(stack.pop(), ()):
            if nxt not in seen:
                seen.add(nxt)
                stack.append(nxt)
    return seen


def _bound_fields(visual: VisualSpec):
    fields = list(visual.fields)
    if visual.order_by is not None:
        fields.append(visual.order_by)
    return fields


def _connectivity_problems(pages: list[PageSpec],
                           relationships,
                           model_table_names: set[str]) -> list[str]:
    adjacency = _m2o_adjacency(relationships)
    reach = {name: _reachable(name, adjacency) for name in model_table_names}
    problems: list[str] = []
    for page in pages:
        for visual in page.visuals:
            fields = _bound_fields(visual)
            column_tables = {f.table for f in fields if f.kind == "column"}
            if len(column_tables) < 2:
                continue
            has_measure = any(f.kind == "measure" for f in fields)
            candidates = model_table_names if has_measure else column_tables
            if not any(column_tables <= reach[center] for center in candidates):
                problems.append(
                    f"{page.folder}/{visual.seed}: no single table reaches all "
                    f"bound column tables {sorted(column_tables)} via active "
                    f"M:1 chains (has_measure={has_measure})")
    return problems


def test_ae_visuals_join_resolvable() -> None:
    tables = {table.name for table in schema.model_tables()}
    problems = _connectivity_problems(
        ae_pages(), schema.model_relationships(), tables)
    assert problems == []


def test_ce_visuals_join_resolvable() -> None:
    tables = {table.name for table in CE_TABLES}
    problems = _connectivity_problems(ce_pages(), CE_RELATIONSHIPS, tables)
    assert problems == []


def test_smoke_visuals_join_resolvable() -> None:
    tables = {table.name for table in schema.model_tables()}
    problems = _connectivity_problems(
        smoke_project().pages, schema.model_relationships(), tables)
    assert problems == []


def test_ce_active_relationship_graph_is_unambiguous() -> None:
    """Same guard as the star-SSOT ambiguity test, for the CE model."""
    edges: dict[str, list[str]] = {}
    for rel in CE_RELATIONSHIPS:
        if rel.active:
            edges.setdefault(rel.to_table, []).append(rel.from_table)

    def _paths(source: str) -> dict[str, int]:
        counts: dict[str, int] = {}

        def _walk(node: str, visited: frozenset[str]) -> None:
            for nxt in edges.get(node, []):
                if nxt in visited:
                    continue
                counts[nxt] = counts.get(nxt, 0) + 1
                _walk(nxt, visited | {nxt})

        _walk(source, frozenset({source}))
        return counts

    for source in (table.name for table in CE_TABLES):
        for target, count in _paths(source).items():
            assert count <= 1, (
                f"ambiguous active filter paths: {source} -> {target} "
                f"({count} distinct paths)")
