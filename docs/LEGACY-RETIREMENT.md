# Legacy retirement checklist — C8CmdLetExportReport

_2026-06-11, Analytics v6 T5. All analytics/PBI tooling now lives in this repo
(`parquet_builder/star/` + `PowerBI/`); `C:\claudecode\C8CmdLetExportReport` is a read-only legacy
reference. Nothing below may be retired until the owner's Desktop verification pass (T6) signs off
on the new AE and CE reports._

## Safe to retire AFTER T6 sign-off

- [ ] `tools\convert_activity_explorer_to_parquet.py` — stale legacy AE converter (588 lines behind
      the ActivityExplorerOld fork it was supposed to track); superseded twice over by star v6.
- [ ] `tools\convert_activity_explorer_optimized_to_parquet.py` — the v5 star converter; its
      surrogate keys and schemas were ported into `parquet_builder/star/` (keys.py, schema.py).
- [ ] `Export-20260609-162814\Export-20260609-162814\ActivityExplorerOld\build_activity_explorer_old_powerbi_data.py`
      — the enhanced fork; its raw-column contract, `_derive_target_domain`, and robust JSON parsing
      are absorbed into star v6 (T1 QFD parity ALL PASS).
- [ ] `tools\build_activity_explorer_powerbi_project.py` — monolithic AE report generator;
      superseded by `PowerBI/builders/build_activity_explorer.py` on the shared engine.
- [ ] `tools\build_content_explorer_powerbi_project.py` — CE report generator; superseded by
      `PowerBI/builders/build_content_explorer.py` (T4 binding/measure parity verified).
- [ ] `PowerBI\ActivityExplorer\` (project folder in the legacy repo) — superseded by the generated
      `PowerBI/projects/ActivityExplorer/` here; keeping two diverging projects recreates the drift
      problem v6 exists to kill.
- [ ] `PowerBI\ContentExplorerSITRisk\` (project folder in the legacy repo) — ported then
      regenerated into `PowerBI/projects/ContentExplorerSITRisk/` here; same drift rationale.
- [ ] `tools\Convert-ActivityExplorerForPowerBI.ps1` — wrapper around the retired AE converters;
      replaced by the `-PowerBIParquet` switch / `py -m parquet_builder.star.convert`.
- [ ] `tools\Build-ActivityExplorerPowerBI.ps1` — wrapper around the retired AE project generator;
      replaced by `.\PowerBI\Build-PowerBI.ps1 -Project ActivityExplorer`.
- [ ] `tools\Build-ContentExplorerPowerBI.ps1` — wrapper around the retired CE project generator;
      replaced by `.\PowerBI\Build-PowerBI.ps1 -Project ContentExplorerSITRisk`.

## Must NOT be retired

- `...\ActivityExplorerOld\pbix\` hand-built report (pbix + .pbit) — the transitional visual
  reference for T6 verification and polish; the only ground truth for theme/vcObjects/drillthrough
  intent until the new report is signed off.
- `tools\convert_content_explorer_to_parquet.py` — still the ONLY producer of the CE parquet model
  the new ContentExplorerSITRisk report consumes; stays until a CE star phase exists.
- `tools\derive_content_explorer_area_tables.py` — companion CE table deriver; same dependency,
  same condition.
- `tools\Convert-ContentExplorerForPowerBI.ps1` — operator wrapper for the still-active CE
  converter above; retires together with it when a CE star phase lands.
- `...\ActivityExplorerOld\SIT-Risk-Analysis-v8.xlsx` + `GAL_Clean.csv` — live enrichment inputs;
  the new star converter consumes them via `ConfigFiles/AEStarEnrichment.local.json`.
- `tools\Update-ActivityExplorerRollingDataset.ps1` — rolling-dataset maintenance; pending the
  central-store track decision, do not touch until that track is decided.

## Not classified (decide at retirement time)

- `tools\Convert-PurviewExportForPowerBI.ps1` — umbrella wrapper dispatching to both
  `Convert-ContentExplorerForPowerBI.ps1` (stays) and `Convert-ActivityExplorerForPowerBI.ps1`
  (retiring); its `-Mode ActivityExplorer`/`All` paths dangle after retirement — either trim it to
  CE-only or retire it together with the CE converter when a CE star phase lands.
