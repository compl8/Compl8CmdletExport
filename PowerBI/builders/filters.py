"""Report / page / visual filter emission (Power BI Layout filter shapes).

All shapes mirror the hand-built reference report:
- report-level Advanced NOT-IN (golden_report_not_in_filter.json)
- page/visual Categorical In
- visual-level measure threshold (Advanced comparison)
- drillthrough page field filters (howCreated: 5, golden_drillthrough_filters.json)

Filter dicts are built WITHOUT a `name`; the project writer assigns a
deterministic 20-hex name at emission time so two runs are identical.
"""

from __future__ import annotations

from .expressions import (
    Field,
    entity_column_expression,
    entity_measure_expression,
    powerbi_literal,
    query_aliases,
)


def _filter_from(table: str) -> tuple[list[dict[str, object]], str]:
    alias = query_aliases([table])[table]
    return [{"Name": alias, "Entity": table, "Type": 0}], alias


def _column_source(alias: str, column: str) -> dict[str, object]:
    return {
        "Column": {
            "Expression": {"SourceRef": {"Source": alias}},
            "Property": column,
        }
    }


def _values_list(values: list[object]) -> list[list[dict[str, object]]]:
    return [[{"Literal": {"Value": powerbi_literal(value)}}] for value in values]


def not_in_filter(table: str, column: str, values: list[object]) -> dict[str, object]:
    """Advanced NOT-IN exclusion filter (the report-level exclusion shape).

    `values` may include None, which encodes as the unquoted `null` literal.
    """
    from_clause, alias = _filter_from(table)
    return {
        "expression": entity_column_expression(table, column),
        "filter": {
            "Version": 2,
            "From": from_clause,
            "Where": [
                {
                    "Condition": {
                        "Not": {
                            "Expression": {
                                "In": {
                                    "Expressions": [_column_source(alias, column)],
                                    "Values": _values_list(values),
                                }
                            }
                        }
                    }
                }
            ],
        },
        "type": "Advanced",
        "howCreated": 1,
        "objects": {
            "general": [
                {
                    "properties": {
                        "isInvertedSelectionMode": {
                            "expr": {"Literal": {"Value": "true"}}
                        }
                    }
                }
            ]
        },
    }


def categorical_in_filter(table: str, column: str, values: list[object]) -> dict[str, object]:
    """Categorical In filter (page- or visual-level include list)."""
    from_clause, alias = _filter_from(table)
    return {
        "expression": entity_column_expression(table, column),
        "filter": {
            "Version": 2,
            "From": from_clause,
            "Where": [
                {
                    "Condition": {
                        "In": {
                            "Expressions": [_column_source(alias, column)],
                            "Values": _values_list(values),
                        }
                    }
                }
            ],
        },
        "type": "Categorical",
        "howCreated": 1,
    }


# ComparisonKind: 0 =, 1 >, 2 >=, 3 <, 4 <=
COMPARISON_GT = 1
COMPARISON_GE = 2
COMPARISON_LT = 3
COMPARISON_LE = 4


def measure_threshold_filter(
    field: Field,
    value: object,
    comparison_kind: int = COMPARISON_GE,
) -> dict[str, object]:
    """Visual-level Advanced threshold filter on a measure (or column)."""
    from_clause, alias = _filter_from(field.table)
    if field.kind == "measure":
        expression = entity_measure_expression(field.table, field.name)
        left = {
            "Measure": {
                "Expression": {"SourceRef": {"Source": alias}},
                "Property": field.name,
            }
        }
    else:
        expression = entity_column_expression(field.table, field.name)
        left = _column_source(alias, field.name)
    return {
        "expression": expression,
        "filter": {
            "Version": 2,
            "From": from_clause,
            "Where": [
                {
                    "Condition": {
                        "Comparison": {
                            "ComparisonKind": comparison_kind,
                            "Left": left,
                            "Right": {"Literal": {"Value": powerbi_literal(value)}},
                        }
                    }
                }
            ],
        },
        "type": "Advanced",
        "howCreated": 1,
    }


def drillthrough_field_filter(table: str, column: str) -> dict[str, object]:
    """Drillthrough page field (section filters.json entry, howCreated: 5)."""
    return {
        "expression": entity_column_expression(table, column),
        "type": "Categorical",
        "howCreated": 5,
    }


def drillthrough_pod_parameters(
    filters_with_names: list[tuple[str, dict[str, object]]],
) -> list[dict[str, object]]:
    """Pod `parameters` entries binding drillthrough fields to section filters.

    Takes (assigned_filter_name, drillthrough_filter_dict) pairs; the pod
    parameter `name` is assigned by the project writer.
    """
    return [
        {"boundFilter": filter_name, "fieldExpr": filter_dict["expression"]}
        for filter_name, filter_dict in filters_with_names
    ]
