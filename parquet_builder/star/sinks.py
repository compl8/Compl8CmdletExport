"""Streaming parquet writers bound to the star schema SSOT."""

from __future__ import annotations

from pathlib import Path
from typing import Any

try:
    import pyarrow as pa
    import pyarrow.parquet as pq
except ModuleNotFoundError as exc:
    pa = None
    pq = None
    PYARROW_IMPORT_ERROR = exc
else:
    PYARROW_IMPORT_ERROR = None

from ..writers import PARQUET_WRITE_OPTS
from .schema import pyarrow_schema


def records_to_table(records: list[dict[str, Any]], schema) -> "pa.Table":
    arrays = [
        pa.array([row.get(field.name) for row in records], type=field.type)
        for field in schema
    ]
    return pa.Table.from_arrays(arrays, schema=schema)


class ParquetSink:
    """Buffered streaming writer for one SSOT table."""

    def __init__(self, output_dir: Path, table_name: str, batch_size: int) -> None:
        self.table_name = table_name
        self.schema = pyarrow_schema(table_name)
        self.batch_size = batch_size
        self.file_path = output_dir / f"{table_name}.parquet"
        self.rows: list[dict[str, Any]] = []
        self.writer: "pq.ParquetWriter | None" = None
        self.count = 0

    def write(self, row: dict[str, Any]) -> None:
        self.rows.append(row)
        if len(self.rows) >= self.batch_size:
            self.flush()

    def flush(self) -> None:
        if not self.rows:
            return
        if self.writer is None:
            self.file_path.parent.mkdir(parents=True, exist_ok=True)
            self.writer = pq.ParquetWriter(str(self.file_path), self.schema, **PARQUET_WRITE_OPTS)
        table = records_to_table(self.rows, self.schema)
        self.writer.write_table(table)
        self.count += table.num_rows
        self.rows = []

    def close(self) -> int:
        self.flush()
        if self.writer is not None:
            self.writer.close()
        elif not self.file_path.exists():
            self.file_path.parent.mkdir(parents=True, exist_ok=True)
            pq.write_table(records_to_table([], self.schema), str(self.file_path), **PARQUET_WRITE_OPTS)
        return self.count


def write_snapshot_table(output_dir: Path, table_name: str,
                         rows: list[dict[str, Any]]) -> int:
    """Write a full dimension/aggregate table in one shot."""
    output_dir.mkdir(parents=True, exist_ok=True)
    table = records_to_table(rows, pyarrow_schema(table_name))
    pq.write_table(table, str(output_dir / f"{table_name}.parquet"), **PARQUET_WRITE_OPTS)
    return table.num_rows
