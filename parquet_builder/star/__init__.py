"""Activity Explorer star-schema v6 data layer (Power BI facing profile).

`schema` is the single source of truth for tables/columns/types/keys/
relationships and Power BI metadata. The parquet converter (`convert`) and the
Power BI TMDL model builder both generate from it.
"""

from .keys import stable_int_id
from .schema import (
    SCHEMA_VERSION,
    ColumnSpec,
    RelationshipSpec,
    TableSpec,
    RELATIONSHIPS,
    TABLES,
    emit_schema_json,
    pyarrow_schema,
    validate_schema,
)

__all__ = [
    "SCHEMA_VERSION",
    "ColumnSpec",
    "RelationshipSpec",
    "TableSpec",
    "RELATIONSHIPS",
    "TABLES",
    "emit_schema_json",
    "pyarrow_schema",
    "stable_int_id",
    "validate_schema",
]
