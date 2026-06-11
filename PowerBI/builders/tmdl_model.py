"""TMDL semantic-model emission generated from `parquet_builder.star.schema`.

Every table, column, type, and relationship in the model comes from the star
schema SSOT — nothing is declared here. Measures are supplied declaratively
by report builders (T3+) via `MeasureSpec`.
"""

from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path

from parquet_builder.star.schema import (
    RelationshipSpec,
    TableSpec,
    model_relationships,
    model_tables,
)

from .expressions import stable_uuid

# Default ParquetRoot parameter value — intentionally a placeholder, never a
# machine-specific path. Users point it at their converter output folder.
DEFAULT_PARQUET_ROOT = "C:\\CHANGEME\\PowerBI-AE-Parquet-v6"

# The SSOT date table (dim_date.date carries "mark as PBI date table key").
DATE_TABLE = "dim_date"
DATE_COLUMN = "date"

# Builder-side display metadata not carried by the SSOT: sort-by-column pairs.
SORT_BY_COLUMN: dict[tuple[str, str], str] = {
    ("dim_date", "month_name"): "month",
    ("dim_date", "month_short"): "month",
    ("dim_date", "day_of_week"): "day_of_week_num",
}

# pyarrow dtype token -> (TMDL dataType, M conversion type)
_TMDL_TYPE_BY_DTYPE = {
    "int64": ("int64", "Int64.Type"),
    "string": ("string", "type text"),
    "bool": ("boolean", "type logical"),
    "timestamp_us": ("dateTime", "type datetime"),
    "date32": ("dateTime", "type date"),
    "double": ("double", "type number"),
}


@dataclass(frozen=True)
class MeasureSpec:
    """Declarative DAX measure consumed by the model emitter."""

    table: str
    name: str
    dax: str
    format_string: str | None = None
    display_folder: str | None = None
    description: str | None = None


def q(name: str) -> str:
    """Quote a TMDL identifier when it is not a bare word."""
    if name and all(ch.isalnum() or ch == "_" for ch in name) and not name[0].isdigit():
        return name
    return "'" + name.replace("'", "''") + "'"


def relationship_id(project_slug: str, rel: RelationshipSpec) -> str:
    seed = (
        f"{project_slug}/relationship/"
        f"{rel.from_table}.{rel.from_column}->{rel.to_table}.{rel.to_column}"
    )
    return str(stable_uuid(seed))


def hidden_columns(table: TableSpec) -> set[str]:
    """Columns hidden in the model: relationship FKs/keys and surrogate keys."""
    hidden: set[str] = set()
    for rel in model_relationships():
        if rel.from_table == table.name:
            hidden.add(rel.from_column)
        if rel.to_table == table.name:
            hidden.add(rel.to_column)
    if table.kind in {"fact", "agg"}:
        if table.key:
            hidden.add(table.key)
        if "activity_id" in table.column_names():
            hidden.add("activity_id")
    if table.kind == "dim" and table.key:
        hidden.add(table.key)
    return hidden


def _m_select_columns(table: TableSpec) -> str:
    return "{" + ", ".join(f'"{c.name}"' for c in table.columns) + "}"


def _m_transform_types(table: TableSpec) -> str:
    items = ", ".join(
        f'{{"{c.name}", {_TMDL_TYPE_BY_DTYPE[c.dtype][1]}}}' for c in table.columns
    )
    return "{" + items + "}"


def _column_lines(table: TableSpec) -> list[str]:
    hidden = hidden_columns(table)
    lines: list[str] = []
    for column in table.columns:
        tmdl_type, _m_type = _TMDL_TYPE_BY_DTYPE[column.dtype]
        if column.description:
            lines.append(f"\t/// {column.description}")
        lines.append(f"\tcolumn {q(column.name)}")
        lines.append(f"\t\tdataType: {tmdl_type}")
        if table.name == DATE_TABLE and column.name == DATE_COLUMN:
            lines.append("\t\tisKey")
        if column.name in hidden:
            lines.append("\t\tisHidden")
        if column.format_string:
            lines.append(f"\t\tformatString: {column.format_string}")
        elif column.dtype == "int64":
            lines.append("\t\tformatString: 0")
        sort_by = SORT_BY_COLUMN.get((table.name, column.name))
        if sort_by:
            lines.append(f"\t\tsortByColumn: {q(sort_by)}")
        lines.append(f"\t\tsummarizeBy: {column.summarize_by}")
        lines.append(f"\t\tsourceColumn: {column.name}")
        lines.append("")
        if column.dtype == "date32":
            lines.append("\t\tannotation UnderlyingDateTimeDataType = Date")
            lines.append("")
        lines.append("\t\tannotation SummarizationSetBy = Automatic")
        lines.append("")
    return lines


def _measure_lines(measures: list[MeasureSpec]) -> list[str]:
    lines: list[str] = []
    for measure in measures:
        if measure.description:
            lines.append(f"\t/// {measure.description}")
        lines.append(f"\tmeasure {q(measure.name)} =")
        for dax_line in measure.dax.splitlines() or [""]:
            lines.append(f"\t\t\t{dax_line}")
        if measure.format_string:
            lines.append(f"\t\tformatString: {measure.format_string}")
        if measure.display_folder:
            lines.append(f"\t\tdisplayFolder: {measure.display_folder}")
        lines.append("")
    return lines


def _partition_lines(table: TableSpec) -> list[str]:
    """M Parquet partition: load, prune to SSOT columns, set SSOT types."""
    return [
        f"\tpartition {q(table.name)} = m",
        "\t\tmode: import",
        "\t\tsource =",
        "\t\t\t\tlet",
        f'\t\t\t\t    Source = Parquet.Document(File.Contents(ParquetRoot & "\\{table.name}.parquet")),',
        f"\t\t\t\t    Columns = Table.SelectColumns(Source, {_m_select_columns(table)}, MissingField.UseNull),",
        f'\t\t\t\t    #"Changed Type" = Table.TransformColumnTypes(Columns, {_m_transform_types(table)})',
        "\t\t\t\tin",
        '\t\t\t\t    #"Changed Type"',
        "",
    ]


def table_tmdl(table: TableSpec, measures: list[MeasureSpec]) -> str:
    lines: list[str] = []
    if table.description:
        lines.append(f"/// {table.description}")
    lines.append(f"table {q(table.name)}")
    if table.name == DATE_TABLE:
        lines.append("\tdataCategory: Time")
    lines.append("")
    lines.extend(_measure_lines([m for m in measures if m.table == table.name]))
    lines.extend(_column_lines(table))
    lines.extend(_partition_lines(table))
    lines.append("\tannotation PBI_ResultType = Table")
    lines.append("")
    return "\n".join(lines)


def relationships_tmdl(project_slug: str) -> str:
    """All SSOT relationships between loaded tables, active and inactive.

    The declared fact_activity_detail -> fact_activity 1:1 is emitted as a
    regular single-direction many-to-one (no cardinality override): TMDL 1:1
    forces both-direction cross-filtering, which the SSOT forbids.
    """
    blocks: list[str] = []
    for rel in model_relationships():
        blocks.append(f"relationship {relationship_id(project_slug, rel)}")
        if not rel.active:
            blocks.append("\tisActive: false")
        blocks.append(f"\tfromColumn: {q(rel.from_table)}.{q(rel.from_column)}")
        blocks.append(f"\ttoColumn: {q(rel.to_table)}.{q(rel.to_column)}")
        blocks.append("")
    return "\n".join(blocks)


def expressions_tmdl(parquet_root: str) -> str:
    escaped = parquet_root.replace('"', '""')
    return (
        f'expression ParquetRoot = "{escaped}" '
        'meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]\n\n'
        "\tannotation PBI_NavigationStepName = Navigation\n\n"
        "\tannotation PBI_ResultType = Text\n"
    )


def model_tmdl() -> str:
    table_names = [table.name for table in model_tables()]
    refs = "\n".join(f"ref table {q(name)}" for name in table_names)
    query_order = json.dumps(["ParquetRoot", *table_names])
    return (
        "model Model\n"
        "\tculture: en-AU\n"
        "\tdefaultPowerBIDataSourceVersion: powerBI_V3\n"
        "\tsourceQueryCulture: en-AU\n"
        "\tdataAccessOptions\n"
        "\t\tlegacyRedirects\n"
        "\t\treturnErrorValuesAsNull\n\n"
        "annotation __PBI_TimeIntelligenceEnabled = 0\n\n"
        f"annotation PBI_QueryOrder = {query_order}\n\n"
        f"{refs}\n\n"
        "ref cultureInfo en-AU\n"
    )


def culture_tmdl() -> str:
    return (
        "cultureInfo en-AU\n\n"
        "\tlinguisticMetadata =\n"
        "\t\t\t{\n"
        '\t\t\t  "Version": "1.0.0",\n'
        '\t\t\t  "Language": "en-US"\n'
        "\t\t\t}\n"
        "\t\tcontentType: json\n"
    )


def database_tmdl(model_name: str) -> str:
    return f"database {q(model_name)}\n\tcompatibilityLevel: 1600\n"


def write_model(
    model_dir: Path,
    *,
    model_name: str,
    project_slug: str,
    parquet_root: str = DEFAULT_PARQUET_ROOT,
    measures: list[MeasureSpec] | None = None,
) -> None:
    """Emit the full Model/ folder (TMDL) from the star schema SSOT."""
    measures = measures or []
    _validate_measures(measures)
    _write(model_dir / "database.tmdl", database_tmdl(model_name))
    _write(model_dir / "expressions.tmdl", expressions_tmdl(parquet_root))
    _write(model_dir / "model.tmdl", model_tmdl())
    _write(model_dir / "relationships.tmdl", relationships_tmdl(project_slug))
    _write(model_dir / "cultures" / "en-AU.tmdl", culture_tmdl())
    for table in model_tables():
        _write(model_dir / "tables" / f"{table.name}.tmdl", table_tmdl(table, measures))


def _validate_measures(measures: list[MeasureSpec]) -> None:
    loaded = {table.name for table in model_tables()}
    names: set[str] = set()
    for measure in measures:
        if measure.table not in loaded:
            raise ValueError(f"measure {measure.name!r}: unknown home table {measure.table!r}")
        if measure.name in names:
            raise ValueError(f"duplicate measure name {measure.name!r}")
        names.add(measure.name)


def _write(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8", newline="\n")
