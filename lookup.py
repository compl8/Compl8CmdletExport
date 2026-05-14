"""Lookup helper for the Compl8CmdletExport record_index parquet table.

Reads content/record_index/source=cmdletexport/*.parquet under a given C8 tuning
input root and answers four kinds of question:

    1. "Where is the record for this file?"
       python lookup.py --root <C8TuningInput> --file-url 'https://contoso.sharepoint.com/sites/HR/Shared%20Documents/x.docx'

    2. "What SITs touched this doc?"
       python lookup.py --root <C8TuningInput> --doc-id <sha1-or-known-doc-id>

    3. "What did we export under this site/folder?"
       python lookup.py --root <C8TuningInput> --site-prefix 'https://contoso.sharepoint.com/sites/HR'

    4. "What's in this exact page file?"
       python lookup.py --root <C8TuningInput> --page-file 'Data/ContentExplorer/SensitiveInformationType/CreditCard/SharePoint-001.json'

Pass --raw to also reopen the source page file and emit the original JSON record
for each hit (useful when extra_fields stripped a field you care about).

Tries DuckDB first (fastest, SQL-friendly), falls back to PyArrow when DuckDB
isn't installed. Output is one JSON object per matching row to stdout.
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

try:
    import duckdb  # noqa: F401
    HAS_DUCKDB = True
except ImportError:
    HAS_DUCKDB = False

try:
    import pyarrow.parquet as pq
    import pyarrow.compute as pc
    import pyarrow.dataset as ds
    HAS_PYARROW = True
except ImportError:
    HAS_PYARROW = False


def _find_index_paths(root: Path) -> list[Path]:
    """Locate the record_index parquet files under a C8 tuning input root."""
    idx_dir = root / "content" / "record_index"
    if not idx_dir.exists():
        return []
    return sorted(idx_dir.rglob("*.parquet"))


def _query_duckdb(index_paths: list[Path], where_clauses: list[str], params: list) -> list[dict]:
    """Run a parameterised query across the record_index parquet files via DuckDB."""
    import duckdb
    parquet_glob = "[" + ",".join(f"'{p.as_posix()}'" for p in index_paths) + "]"
    where_sql = " AND ".join(where_clauses) if where_clauses else "1=1"
    sql = f"""
        SELECT *
        FROM read_parquet({parquet_glob})
        WHERE {where_sql}
        ORDER BY page_file, page_offset
    """
    con = duckdb.connect(":memory:")
    try:
        cur = con.execute(sql, params)
        columns = [c[0] for c in cur.description]
        return [dict(zip(columns, row)) for row in cur.fetchall()]
    finally:
        con.close()


def _query_pyarrow(index_paths: list[Path], expr) -> list[dict]:
    """Fallback when DuckDB is unavailable. `expr` is a pyarrow.compute expression."""
    dataset = ds.dataset([str(p) for p in index_paths], format="parquet")
    table = dataset.to_table(filter=expr)
    return table.to_pylist()


def _emit_raw_record(row: dict, root: Path) -> dict | None:
    """Reopen the original page file and return the record at page_offset."""
    page_file = row.get("page_file")
    offset = row.get("page_offset")
    if not page_file or offset is None:
        return None
    # page_file is relative to the export root, which is stored on _source_export_dir
    export_dir = Path(row.get("_source_export_dir") or "")
    page_path = export_dir / page_file
    if not page_path.exists():
        # Older indexes may have stored absolute paths; try literal
        page_path = Path(page_file)
        if not page_path.exists():
            return None
    suffix = page_path.suffix.lower()
    try:
        if suffix == ".jsonl":
            with open(page_path, "r", encoding="utf-8-sig") as f:
                for i, line in enumerate(f, start=1):
                    if i == int(offset):
                        return json.loads(line)
            return None
        with open(page_path, "r", encoding="utf-8-sig") as f:
            data = json.load(f)
        if isinstance(data, dict) and "Records" in data:
            records = data["Records"]
        elif isinstance(data, list):
            records = data
        else:
            return None
        if 0 <= int(offset) < len(records):
            return records[int(offset)]
    except (OSError, json.JSONDecodeError):
        return None
    return None


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--root", required=True, help="C8 tuning input root (contains content/record_index/)")
    parser.add_argument("--file-url", help="Exact file_url to look up")
    parser.add_argument("--file-url-like", help="file_url LIKE pattern (use %% as wildcard, e.g. '%%budget%%')")
    parser.add_argument("--doc-id", help="Exact doc_id to look up")
    parser.add_argument("--site-prefix", help="Prefix match against file_url (path-style)")
    parser.add_argument("--page-file", help="List all records that came from this page file")
    parser.add_argument("--tag-name", help="Restrict to a specific tag_name (e.g. 'Credit Card')")
    parser.add_argument("--workload", help="Restrict to a workload (SharePoint/OneDrive/Exchange/Teams)")
    parser.add_argument("--limit", type=int, default=50, help="Max rows to return (default 50; 0 = unlimited)")
    parser.add_argument("--raw", action="store_true", help="Also emit the original source record for each hit")
    parser.add_argument("--format", choices=["json", "ndjson", "summary"], default="ndjson",
                        help="Output format (default ndjson)")
    args = parser.parse_args()

    root = Path(args.root).resolve()
    if not root.exists():
        print(f"ERROR: root does not exist: {root}", file=sys.stderr)
        return 1

    index_paths = _find_index_paths(root)
    if not index_paths:
        print(f"ERROR: no record_index parquet files under {root / 'content' / 'record_index'}", file=sys.stderr)
        print("       (Has build_unified_parquet.py been run against this export?)", file=sys.stderr)
        return 2

    if not any([args.file_url, args.file_url_like, args.doc_id, args.site_prefix, args.page_file]):
        print("ERROR: provide at least one of --file-url, --file-url-like, --doc-id, --site-prefix, --page-file",
              file=sys.stderr)
        return 1

    if HAS_DUCKDB:
        where = []
        params: list = []
        if args.file_url:
            where.append("file_url = ?"); params.append(args.file_url)
        if args.file_url_like:
            where.append("file_url LIKE ?"); params.append(args.file_url_like)
        if args.doc_id:
            where.append("doc_id = ?"); params.append(args.doc_id)
        if args.site_prefix:
            where.append("file_url LIKE ?"); params.append(args.site_prefix.rstrip("/") + "%")
        if args.page_file:
            where.append("page_file = ?"); params.append(args.page_file)
        if args.tag_name:
            where.append("tag_name = ?"); params.append(args.tag_name)
        if args.workload:
            where.append("workload = ?"); params.append(args.workload)
        rows = _query_duckdb(index_paths, where, params)
    elif HAS_PYARROW:
        # Build a PyArrow filter expression
        exprs = []
        if args.file_url:
            exprs.append(pc.field("file_url") == args.file_url)
        if args.file_url_like:
            # PyArrow has no LIKE; approximate with match_substring on the wildcard contents
            pat = args.file_url_like.strip("%")
            exprs.append(pc.match_substring(pc.field("file_url"), pat))
        if args.doc_id:
            exprs.append(pc.field("doc_id") == args.doc_id)
        if args.site_prefix:
            exprs.append(pc.starts_with(pc.field("file_url"), args.site_prefix.rstrip("/")))
        if args.page_file:
            exprs.append(pc.field("page_file") == args.page_file)
        if args.tag_name:
            exprs.append(pc.field("tag_name") == args.tag_name)
        if args.workload:
            exprs.append(pc.field("workload") == args.workload)
        from functools import reduce
        import operator
        combined = reduce(operator.and_, exprs) if exprs else None
        rows = _query_pyarrow(index_paths, combined)
    else:
        print("ERROR: neither duckdb nor pyarrow is installed. Install one with:", file=sys.stderr)
        print("       pip install duckdb   # recommended", file=sys.stderr)
        print("       pip install pyarrow  # fallback", file=sys.stderr)
        return 3

    if args.limit > 0:
        rows = rows[: args.limit]

    if args.raw:
        for row in rows:
            row["_raw_record"] = _emit_raw_record(row, root)

    if args.format == "summary":
        print(f"Matches: {len(rows)}")
        for row in rows:
            print(f"  {row.get('tag_name'):<30} {row.get('workload'):<12} "
                  f"{row.get('page_file')}#{row.get('page_offset')}  {row.get('file_url')}")
    elif args.format == "json":
        json.dump(rows, sys.stdout, indent=2, default=str)
        print()
    else:
        for row in rows:
            print(json.dumps(row, default=str))

    return 0


if __name__ == "__main__":
    sys.exit(main())
