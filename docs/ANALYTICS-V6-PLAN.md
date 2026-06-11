# Analytics v6 — AE star schema + Power BI in this repo

_2026-06-11. Supersedes the location (not the content) of
`C:\claudecode\C8CmdLetExportReport\docs\AE-V6-SUPERSET-PLAN.md` — the v6 schema changes, bug fixes
(F1 enrichment loss, F2 record loss), transfer matrix, and builder capability list defined there are
implemented HERE. C8CmdLetExportReport becomes a legacy reference; ignore its duplicated tooling._

## Decisions (owner, 2026-06-11)

- All analytics/PBI work consolidates into Compl8CmdletExport. Fresh layout; do not port the mess.
- **Schema is a single source of truth**: one Python module declares tables/columns/types/keys/
  relationships/PBI metadata; the parquet converter AND the Power BI TMDL model are both generated
  from it. The current out-of-step PBI-vs-data problem becomes structurally impossible.
- Data formats serve three consumers: Power BI, Python, MCP (machine-readable `schema.json` +
  manifest emitted next to the parquet; MCP server itself is out of scope now).
- Indexing (`record_index`, `lookup.py`) stays separate — out of scope this phase.
- Testing: run conversions against the existing QFD export in place:
  `C:\claudecode\C8CmdLetExportReport\Export-20260609-162814\Export-20260609-162814`
  (447,017 raw records / 30 days; legacy `PowerBI_Data\` output there is the parity baseline).
- Column naming aligns with the existing `parquet_builder` conventions (snake_case, constants.py
  rename style) so c8_tuning_input and the star profile stay coherent.

## Layout

```
parquet_builder/
  star/                    # NEW: AE star schema v6 (PBI-facing profile)
    __init__.py
    schema.py              # SSOT: TableSpec/ColumnSpec/RelationshipSpec + PBI metadata + emit schema.json
    keys.py                # stable sha1-based int surrogate keys (ported from optimized v5)
    enrich.py              # risk workbook + GAL loaders; HARD FAIL unless --allow-unenriched (F1 fix)
    convert.py             # AE export dir -> v6 parquet (CLI: py -m parquet_builder.star.convert)
PowerBI/
  builders/
    pbi_project.py         # shared codegen core: TMDL emission FROM star.schema + visual factories,
                           # theme embedding, vcObjects titles, filters/drillthrough, layout constants
    build_activity_explorer.py
    build_content_explorer.py
  projects/ActivityExplorer/pbix/        # generated (tracked)
  projects/ContentExplorerSITRisk/pbix/  # ported then regenerated (tracked)
  themes/  CustomVisuals/                # CY26SU02-derived theme; ForceGraph, Sankey, WordCloud
  Build-PowerBI.ps1                      # wrapper -> pbi-tools.core (C:\Tools\pbi-tools-net9, needs
                                         # DOTNET_ROLL_FORWARD=Major)
```

## Reference sources (read-only; in C8CmdLetExportReport)

- `tools\convert_activity_explorer_optimized_to_parquet.py` — v5 star: `_stable_int_id`, SCHEMAS, aggs
- `Export-20260609-162814\Export-20260609-162814\ActivityExplorerOld\build_activity_explorer_old_powerbi_data.py`
  — the enhanced fork: `_REPORT_ACTIVITY_COLUMNS` (the definitive 70+ col raw contract),
  `_derive_target_domain`, robust JSON parsing (`tools\convert_activity_explorer_to_parquet.py` is 588
  lines STALE vs this — port from the fork)
- `tools\convert_content_explorer_to_parquet.py` — `load_risk_workbook` (line ~177)
- `tools\build_activity_explorer_powerbi_project.py` (64KB) + the hand-built old report
  `...\ActivityExplorerOld\pbix\` (theme CY26SU02, vcObjects patterns, drillthrough filters.json)
- `PowerBI\ContentExplorerSITRisk\` + its side Theme.json
- Enrichment inputs for QFD testing: `...\ActivityExplorerOld\{SIT-Risk-Analysis-v8.xlsx, GAL_Clean.csv}`

## v6 schema (summary — full detail in the C8CmdLetExportReport plan doc)

Star core: fact_activity (+user_type, data_platform, app_identity_id), fact_activity_sit
(+classifier_type, target_domain_id, policy_rule_id), fact_policy_activity, fact_email_recipient,
fact_activity_detail (drop record_identity/file_path; add ~30 typed endpoint/DLP contract cols +
extra_json catch-all + drift report), NEW fact_email_detail (subject/message_id/attachments), NEW
fact_copilot_interaction, NEW dim_app_identity, NEW dim_source_page (provenance), dims (sit: all 18+
reference cols, all 1,194 workbook rows w/ observed flag; user: full GAL + has_activity; date:
continuous + month_short/week_of_year; location/domain/email_address/policy/workload/activity_type),
4 agg tables (+ consider agg_domain_sit_day), activity_record_index (+page_id), optional pipeline-only
archive_raw (--archive-raw default on). Relationships declared in SSOT incl. target_domain (active),
originating_domain + target_location (inactive). 269-SIT exclusion applied at ETL. Anchor time-intel
to MAX(date). Fix F2 (2 records lost vs legacy: be6af93c-..., c0a27be5-... — diagnose during parity).

## Task checklist

- [ ] **T1 — Data layer**: `parquet_builder/star/` (schema.py SSOT, keys, enrich, convert CLI) + pytest
      (synthetic fixtures) + QFD parity run: 447,011 rows, nonzero risk scores, 30-day continuous
      dim_date, ~49,450 email subjects, superset checklist vs legacy columns, F2 diagnosed/fixed.
      Output to `<export>\PowerBI-AE-Parquet-v6\` + schema.json + manifest.
- [ ] **T2 — PBI builder core**: `PowerBI/builders/pbi_project.py` — TMDL model generated from
      star.schema; visual factories (incl. pie/line/pivotTable/clusteredColumn/actionButton/WordCloud);
      theme embedding; vcObjects titles (fix non-schema `title` key); report/page/visual filter +
      drillthrough emission; layout constants; LinguisticSchema/culture.
- [ ] **T3 — AE report superset**: 29-page transfer matrix implemented; 45 measures ported/rewritten
      (efficient star DAX per matrix); compile via pbi-tools.core.
- [ ] **T4 — CE report port**: bring ContentExplorerSITRisk builder+project in, rewire paths/theme,
      regenerate + compile (CE schema work itself is a later phase).
- [ ] **T5 — Wiring & docs**: Build-PowerBI.ps1, entry-script switch (e.g. -PowerBIParquet), README,
      CLAUDE.md architecture update, retire-list for the legacy converters in the old repo.
- [ ] **T6 — Owner Desktop verification pass** (screenshot loop), then iterate polish.

One commit per task; testing on QFD data in place; review between tasks.
