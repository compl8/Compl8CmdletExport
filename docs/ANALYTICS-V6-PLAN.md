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

- [x] **T1 — Data layer** (`688f073`): `parquet_builder/star/` — 27 tables, 65 relationships, CLI
      `py -m parquet_builder.star.convert`. QFD parity ALL PASS (447,011 exact; risk max 17,153 exact;
      30 continuous dates; 49,450 email details; 955,581 SIT matches pre-exclusion; 79/79 legacy
      columns mapped). F2 root cause: PowerShell ConvertTo-Json unwraps 1-element arrays → dict
      Records pages; fixed. Output: `<export>\PowerBI-AE-Parquet-v6\` + schema.json + manifest.json.
      Notes for T2/T3: exclusions applied at ETL (24.5% of SIT match rows on QFD; activities risk
      metrics still include them, faithful to legacy); fact_activity_detail→fact_activity declared 1:1
      relationship (confirm TMDL treatment); derive_target_domain ON by default (old report depends on
      it, incl. dotted-ItemName-derived "domains"); contract cols typed string; dim_sit = 1,194 workbook
      + observed-unknown rows with observed flag.
- [x] **T2 — PBI builder core** (`bbfb9a3`): PowerBI/builders/ engine — TMDL generated from star SSOT
      (25 tables/64 rels, date table marked, hidden FKs), all visual factories, theme
      (themes/Compl8.Theme.json = CY26SU02 + curated palette) embedded, correct vcObjects titles,
      filters + drillthrough emission, layout grid, en-AU/LinguisticSchema, package-on-demand custom
      visuals. Smoke project compile gate EXIT 0. API documented in builders/__init__ + build_smoke.py.
- [x] **T3 — AE report superset** (`22f736a`): build_activity_explorer.py + ae_*.py — 29 pages (full
      legacy mapping in module docstring), 214 titled visuals, 4 drill pages, 73 measures (45 legacy
      names kept, efficient star DAX, MAX-date anchoring, zero TODAY()). Compile EXIT 0; 80 tests.
      Owner Desktop command: `.\PowerBI\Build-PowerBI.ps1 -Project ActivityExplorer -ParquetRoot
      "C:\claudecode\C8CmdLetExportReport\Export-20260609-162814\Export-20260609-162814\PowerBI-AE-Parquet-v6"`
      then open PowerBI\projects\ActivityExplorer\ActivityExplorerRisk.pbit.
- [x] **T4 — CE report port**: build_content_explorer.py + ce_schema/ce_measures/ce_pages_* on the T2
      engine, slug `content-explorer-sit-risk`. ce_schema = declarative ModelSource mirroring the
      EXISTING CE parquet (22 tables/21 rels, legacy parquet file names, CHANGEME ParquetRoot; engine
      ModelSource refactor keeps AE/smoke byte-identical). 15 pages / 165 titled visuals (164 legacy +
      Back button; binding/displayName/sort parity verified against the legacy project), 71 measures
      (3 DimFile FILTER iterators -> column predicates, same names/semantics; display folders added).
      040_File_Drillthrough now REALLY wired (DimFile.file_name/file_extension, DimSIT.sit_name,
      DimLocation.location_name + Back button). Compile + generate-bim EXIT 0; 95 tests. NOTE: the QFD
      export has no CE data — owner verification needs a CE conversion run first (legacy repo
      convert_content_explorer_to_parquet.py), then
      `.\PowerBI\Build-PowerBI.ps1 -Project ContentExplorerSITRisk -ParquetRoot "<ce-parquet-dir>"`.
- [ ] **T5 — Wiring & docs**: Build-PowerBI.ps1, entry-script switch (e.g. -PowerBIParquet), README,
      CLAUDE.md architecture update, retire-list for the legacy converters in the old repo.
- [ ] **T6 — Owner Desktop verification pass** (screenshot loop), then iterate polish.

One commit per task; testing on QFD data in place; review between tasks.
