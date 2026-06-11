"""Spec dataclasses and helper constructors for the star schema SSOT.

Column dtypes are pyarrow type tokens kept as strings so the schema modules
import without pyarrow installed; `schema.pyarrow_schema()` resolves them
lazily.
"""

from __future__ import annotations

from dataclasses import dataclass, field

SCHEMA_VERSION = 6
SCHEMA_PROFILE = "powerbi_star"

# pyarrow dtype token -> Power BI / TMDL dataType
PBI_TYPE_BY_DTYPE = {
    "int64": "Int64",
    "string": "String",
    "bool": "Boolean",
    "timestamp_us": "DateTime",
    "date32": "DateTime",
    "double": "Double",
}

VALID_SUMMARIZE_BY = {"none", "sum", "count", "min", "max", "average"}
VALID_KINDS = {"dim", "fact", "agg", "index", "pipeline_only"}


@dataclass(frozen=True)
class ColumnSpec:
    name: str
    dtype: str  # pyarrow dtype token: int64 | string | bool | timestamp_us | date32 | double
    nullable: bool = True
    pbi_type: str = ""  # derived from dtype when empty
    format_string: str | None = None
    summarize_by: str = "none"
    description: str = ""

    def resolved_pbi_type(self) -> str:
        return self.pbi_type or PBI_TYPE_BY_DTYPE[self.dtype]


@dataclass(frozen=True)
class TableSpec:
    name: str
    kind: str  # dim | fact | agg | index | pipeline_only
    columns: tuple[ColumnSpec, ...]
    key: str | None = None
    description: str = ""

    def column(self, name: str) -> ColumnSpec:
        for col in self.columns:
            if col.name == name:
                return col
        raise KeyError(f"{self.name} has no column {name!r}")

    def column_names(self) -> list[str]:
        return [col.name for col in self.columns]


@dataclass(frozen=True)
class RelationshipSpec:
    """A single-direction filter relationship (cross-filter Both is never used)."""

    from_table: str
    from_column: str
    to_table: str
    to_column: str
    active: bool = True
    cross_filter: str = field(default="single", init=False)


def _c(name: str, dtype: str, desc: str = "", *, nullable: bool = True,
       fmt: str | None = None, agg: str = "none") -> ColumnSpec:
    return ColumnSpec(
        name=name, dtype=dtype, nullable=nullable,
        format_string=fmt, summarize_by=agg, description=desc,
    )


def _key(name: str, dtype: str = "int64", desc: str = "") -> ColumnSpec:
    return _c(name, dtype, desc or "Surrogate key.", nullable=False)


def _fk(name: str, desc: str = "") -> ColumnSpec:
    return _c(name, "int64", desc or "Dimension foreign key.")


def _metric(name: str, desc: str = "", dtype: str = "int64") -> ColumnSpec:
    return _c(name, dtype, desc, fmt="#,0", agg="sum")
