# Scaling Plan — DuckDB central store + Power BI projections

_2026-06-12. Extends docs/ANALYTICS-V6-PLAN.md (T1-T5 complete) and the central-store design in
REVIEW-20260610 Section D. Decision context: Power BI import mode (VertiPaq) handles tens of
millions of AE events comfortably on 32GB; the string-heavy detail tables are the limiter. DuckDB
in front via plain ODBC is still import mode (no scale win); DirectQuery needs a custom connector
(fragile, breaks custom visuals' interactivity, relaxed security on every consumer machine)._

## Architecture (three layers, each scoped to what it scales)

1. **DuckDB central store — effectively unbounded.** One DuckDB database (or DuckDB views over
   the parquet store) holding every export run at full grain. This is the SAME store the
   merge/central-store track already planned: run-stamped star parquet ingested per run,
   compaction/dedup on record_id (AE) and generation/cell-assertion logic (CE) implemented as
   DuckDB SQL. Billions of rows on a workstation; survives the 30-day AE API horizon.
2. **Power BI import projections — bounded by design, not by data.** The report model never
   imports the raw store. It imports: (a) the pre-agg tables (O(dims × days) — overview pages
   never grow with event volume), (b) date-windowed fact/detail grain for drill pages (window
   size is a conversion/projection parameter), (c) dims in full. Same star schema, same report.
3. **Escalation paths, only if a tenant outgrows windowed import:** DirectQuery custom connector
   for detail tables only (composite model); or Fabric Direct Lake if the engagement is cloud —
   the parquet store promotes to OneLake/delta without redesign.

## Why this preserves the anti-drift property

The star schema SSOT (parquet_builder/star/schema.py) gains a third emitter: alongside parquet
writers and TMDL, it generates **DuckDB DDL/views** (CREATE VIEW per table over the run-stamped
parquet, plus the live-set/compaction views). Store, report model, and Python/MCP consumers stay
aligned by construction.

## Work items (sequenced; builds on existing code)

| # | Item | Notes | Effort |
|---|------|-------|--------|
| 1 | `star/store.py`: DuckDB DDL/view emitter from the SSOT + `ingest` CLI (register a run dir into the store with provenance from manifest.json) | store dir layout per REVIEW-20260610 D3 (ingest_log.jsonl, ae_watermark.json) | M |
| 2 | AE compaction: dedup on record_id per month partition; boundary-day overlap resolution | DuckDB SQL; one-time RecordIdentity stability check vs overlapping pulls (decision D6) | S-M |
| 3 | Projection extractor: `star/project.py --window <days> --output <dir>` emits the import-mode parquet set (full dims+aggs, windowed facts/details) from the store | the report's ParquetRoot points at a projection, unchanged TMDL | M |
| 4 | CE generations + cell_assertions in the store (per REVIEW Section D) | after the CE star schema phase | M-L |
| 5 | Optional: evaluate DuckDB DirectQuery custom connector for detail tables | only if a tenant outgrows windowed import (~100M+ match rows) | L |

Prerequisite reality check: at reference-tenant volume (~450K events/month) windowed import is years of
headroom; items 1-3 are about the scheduled-accumulation story (monthly runs appending to one
store), not rescue from a present limit.
