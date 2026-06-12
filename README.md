# Compl8 Cmdlet Export Tool

PowerShell utility for extracting Microsoft Purview compliance configuration via the ExchangeOnlineManagement module and Security & Compliance PowerShell.

Exports DLP policies, sensitivity labels, retention labels, Content Explorer data, Activity Explorer data, eDiscovery cases, and RBAC configuration.

## Prerequisites

- PowerShell 7+
- ExchangeOnlineManagement 3.2.0+

```powershell
Install-Module -Name ExchangeOnlineManagement -Force -Scope CurrentUser
```

## Quick Start

```powershell
# CLI help
.\Export-Compl8Configuration.ps1 --help

# Interactive menu
.\Export-Compl8Configuration.ps1

# Full export
.\Export-Compl8Configuration.ps1 -FullExport

# Full export, skipping slow Content/Activity Explorer
.\Export-Compl8Configuration.ps1 -FullExport -NoActivity -NoContent
```

## Repository Layout

The entry script is still operator-first, but the runtime is now split into a small set of stable folders. Keep this structure intact if you copy or package the tool:

```text
Export-Compl8Configuration.ps1
App/
  Host/
  Orchestrator/
  Providers/
Modules/
  Compl8ExportFunctions.psm1
  Compl8ExportFunctions/
ConfigFiles/
build_unified_parquet.py
parquet_builder/
PowerBI/
README.md
```

`Export-Compl8Configuration.ps1` expects `App/`, `Modules/`, and `ConfigFiles/` to sit beside it. Do not move individual files out of that layout. `parquet_builder/` is also required next to the entry script if you use `-PowerBIParquet`; `PowerBI/` only matters for building the reports.

## Run On Another Machine

1. Copy the repository root, or extract a portable zip, while preserving the folder structure above.
2. Install PowerShell 7 and `ExchangeOnlineManagement` 3.2.0+.
3. If you want certificate auth, copy `ConfigFiles/AuthConfig.example.json` to `ConfigFiles/AuthConfig.json` and populate it.
4. If you want Parquet conversion, install Python and `pyarrow`.

```powershell
Install-Module -Name ExchangeOnlineManagement -Force -Scope CurrentUser
```

```bash
pip install pyarrow
```

For a full development and test environment, use `pip install -r requirements.txt`.

## Export Modes

| Mode | Parameter | Description |
|------|-----------|-------------|
| Interactive Menu | `-Menu` (or no params) | Select export type from menu |
| Full Export | `-FullExport` | All data types |
| DLP Only | `-DlpOnly` | DLP policies, rules, and SITs |
| Labels Only | `-LabelsOnly` | Sensitivity and retention labels |
| Content Explorer | `-ContentExplorer` | Content by classifier/workload |
| Activity Explorer | `-ActivityExplorer` | Activity logs with filters |
| eDiscovery | `-eDiscoveryOnly` | eDiscovery cases and searches |
| RBAC | `-RbacOnly` | Role groups and members |

### Common Parameters

| Parameter | Range | Description |
|-----------|-------|-------------|
| `-OutputFormat` | JSON/CSV | Output file format (default: JSON) |
| `-OutputDirectory` | path | Custom output location (default: `./Output`) |
| `-UserPrincipalName` | UPN | Pre-fill authentication identity |
| `-PageSize` | 1-5000 | API pagination size |
| `-PastDays` | 1-30 | Activity Explorer history window |
| `-NoActivity` | switch | Skip Activity Explorer in full export |
| `-NoContent` | switch | Skip Content Explorer in full export |

### Content Explorer Parameters

| Parameter | Description |
|-----------|-------------|
| `-WorkerCount` | 2-16 parallel workers for multi-terminal export |
| `-CEAllSITs` | Auto-discover and export all SITs from tenant |
| `-CEWorkloads` | Limit to specific workloads (Exchange, SharePoint, OneDrive, Teams) |
| `-CEResumeDir` | Resume a previous incomplete export |
| `-CERetryDir` | Re-export only tasks with >2% count discrepancy |
| `-CETasksCsv` | Run from a specific task CSV file |

### Activity Explorer Parameters

| Parameter | Description |
|-----------|-------------|
| `-AEResume` | Resume from last successful page using run tracker |
| `-AEWorkerCount` | 2-16 parallel workers for multi-terminal export |
| `-AEResumeDir` | Resume a specific previous AE export |

### Post-Export Parameters

| Parameter | Description |
|-----------|-------------|
| `-UnifiedParquet` | Convert JSON output to unified Parquet format after export |
| `-UnifiedParquetDir` | Parquet output directory (default: the export run's `C8TuningInput` folder) |
| `-UsersCsv` | GAL Scraper or Entra user CSV for enrichment (repeatable) |
| `-PowerBIParquet` | Convert Activity Explorer output to the Power BI star-schema (v6) Parquet model |
| `-PowerBIParquetDir` | Star-schema output directory (default: the export run's `PowerBI-AE-Parquet-v6` folder) |

## Content Explorer

Content Explorer supports single-terminal and multi-terminal parallel export modes with full resumability.

### Single Terminal

```powershell
.\Export-Compl8Configuration.ps1 -ContentExplorer
```

### Multi-Terminal (Parallel)

Spawns worker terminals that process tasks in parallel, coordinated through a file-drop protocol with signed task payloads.

```powershell
# Spawn 4 parallel workers
.\Export-Compl8Configuration.ps1 -ContentExplorer -WorkerCount 4

# Resume a previous export (works even if all terminals crashed)
.\Export-Compl8Configuration.ps1 -CEResumeDir "Output\Export-20260130-100000"

# Retry only tasks with >2% count discrepancy
.\Export-Compl8Configuration.ps1 -CERetryDir "Output\Export-20260130-100000"
```

Or use the interactive menu (options 1-4) to configure Content Explorer modes.

### Aggregate Caching

On startup, the tool checks for recent aggregate CSV files (last 30 days) and offers to reuse them. This skips the aggregate query phase, which can take 30+ minutes on large tenants.

### SIT Filtering

`ConfigFiles/SITstoSkip.json` controls which Sensitive Information Types are excluded from Content Explorer exports. Set a SIT to `"True"` to skip it, `"False"` to include it.

### SIT Reference Snapshot

Each Content Explorer / Activity Explorer export run writes a SIT GUID-to-name reference snapshot to `<ExportDir>\CurrentTenantSITs.json` (via `Export-SitReferenceSnapshot`), built entirely from supported Security & Compliance cmdlets:

- `Get-DlpSensitiveInformationType` — the flat tenant SIT list (these names win on conflict).
- `Get-DlpSensitiveInformationTypeRulePackage` — every rule package's `ClassificationRuleCollection` XML. The raw XML is saved per pack under `<ExportDir>\Data\Reference\RulePackages\` and parsed for rule GUID-to-name pairs, which fills in the Microsoft-internal classification sub-entity GUIDs (Entities and Affinities) that the flat list does not surface but Activity Explorer DLP detections reference.

The snapshot lands where the star-schema converter auto-detects it, so SIT names ship with each tenant export and resolve automatically on conversion. To regenerate it for an existing export from a connected S&C session:

```powershell
Import-Module .\Modules\Compl8ExportFunctions.psm1
Export-SitReferenceSnapshot -ExportRunDirectory "Output\Export-20260130-100000" -Force
```

Copy names from the snapshot into `SITstoSkip.json` to exclude them from Content Explorer exports.

### Trainable Classifier Names — Provided Externally

Microsoft does not expose a public cmdlet or Graph API for enumerating trainable classifiers, so the export tool cannot discover them itself. The names come from a cache file that is produced externally — the tool owner's separate **GetTCs** utility (distributed independently of this repository) generates it; drop its output at:

```
ConfigFiles/CurrentTenantTCs.local.json    (gitignored — tenant data)
```

Expected JSON shape (only `Classifiers[].Name` is strictly required; extra properties are ignored):

```json
{
  "SchemaVersion": 1,
  "DiscoveredAt": "2026-05-14T07:00:00Z",
  "ClassifierCount": 2,
  "Classifiers": [
    { "Id": "<guid>", "Name": "Source code", "DisplayName": "Source code",
      "Type": "GlobalOOB", "ModelStatus": "Stable", "IsDeprecated": false }
  ]
}
```

`Get-TrainableClassifiersFromCache` reads this file to feed `TrainableClassifier` tag names into the Content Explorer task plan. If the file is missing the export logs a warning and proceeds without trainable classifiers; if `DiscoveredAt` is older than 30 days a staleness warning is logged. The `-RefreshTrainableClassifiers` switch reports cache status (it does not refresh anything itself).

GetTCs can also emit a CSV with `Id,DisplayName` columns — `py -m parquet_builder.star.extract_sit_names` accepts that shape directly as a `--input` artifact when building SIT/classifier name maps.

## Activity Explorer

```powershell
# Export last 7 days (default)
.\Export-Compl8Configuration.ps1 -ActivityExplorer

# Export last 30 days
.\Export-Compl8Configuration.ps1 -ActivityExplorer -PastDays 30

# Multi-terminal with 4 workers
.\Export-Compl8Configuration.ps1 -ActivityExplorer -AEWorkerCount 4 -PastDays 14

# Resume after failure
.\Export-Compl8Configuration.ps1 -ActivityExplorer -AEResume

# Resume a specific previous export
.\Export-Compl8Configuration.ps1 -AEResumeDir "Output\Export-20260206-100000"
```

### Monitoring Progress

```powershell
# Tail Content Explorer progress (in another terminal)
Get-Content -Path "Output\Export-*\_Logs\ContentExplorer-Progress.log" -Wait -Tail 20

# Tail Activity Explorer progress
Get-Content -Path "Output\Export-*\_Logs\ActivityExplorer-Progress.log" -Wait -Tail 20
```

## Parquet Conversion

Export output can be converted to unified Hive-partitioned Parquet format for downstream analytics.

### From PowerShell (post-export)

```powershell
# Automatic conversion after export.
# Writes to Output\Export-YYYYMMDD-HHMMSS\C8TuningInput by default.
.\Export-Compl8Configuration.ps1 -FullExport -UnifiedParquet

# With custom output directory and user enrichment
.\Export-Compl8Configuration.ps1 -ActivityExplorer -UnifiedParquet -UnifiedParquetDir "D:\C8TuningRuns\Run-001" -UsersCsv "users.csv"
```

### Standalone Python

```bash
pip install pyarrow

# Convert an export.
# Writes to Output/Export-20260301-090000/C8TuningInput by default.
python build_unified_parquet.py --input-dir Output/Export-20260301-090000

# With user enrichment (GAL Scraper or Entra CSV)
python build_unified_parquet.py --input-dir Output/Export-20260301-090000 --output-dir D:/C8TuningRuns/Run-001 --users-csv users.csv
```

The C8 tuning input root contains `content/content_files`, `content/sit_detections`, and a `c8_tuning_input_manifest.json` file that downstream run pickers can use to identify the export.

## Power BI Reports & Star Schema

The repository produces three analytics outputs:

| Output | How | Consumers |
|--------|-----|-----------|
| C8 tuning input (unified Parquet) | `-UnifiedParquet` / `build_unified_parquet.py` | Classifier tuning pipeline |
| Power BI star schema v6 (27 tables) | `-PowerBIParquet` / `py -m parquet_builder.star.convert` | Power BI, Python, MCP |
| Power BI templates (`.pbit`) | `.\PowerBI\Build-PowerBI.ps1` | Power BI Desktop |

### Single Source of Truth (anti-drift)

`parquet_builder/star/schema.py` declares every table, column, type, key, relationship, and Power BI metadata item once. Both the Parquet converter **and** the Power BI TMDL model are generated from it, so the data and the report model cannot drift apart. Never hand-edit the generated projects under `PowerBI/projects/` — change the schema or the builders and regenerate.

### Star-Schema Conversion

```powershell
# Post-export (Activity Explorer data only; skipped with a warning otherwise)
.\Export-Compl8Configuration.ps1 -ActivityExplorer -PastDays 30 -PowerBIParquet

# Standalone, against an existing export
py -m parquet_builder.star.convert --input-dir "Output\Export-20260611-090000"
```

Default output is the export run's `PowerBI-AE-Parquet-v6` folder. Alongside the Parquet files the converter writes `schema.json` (machine-readable model contract for Python/MCP consumers), `manifest.json` (row counts, enrichment provenance, exclusion and drift summaries), and `SchemaDrift.json` when unknown raw fields were seen.

**Enrichment:** copy `ConfigFiles/AEStarEnrichment.example.json` to `AEStarEnrichment.local.json` (gitignored) and point it at the SIT risk workbook and department CSV. Without it, `-PowerBIParquet` builds an unenriched model (`--allow-unenriched`) with a prominent warning — risk scores are 0 and departments unmapped. The standalone CLI hard-fails instead unless `--allow-unenriched` is passed explicitly.

**SIT display names:** `dim_sit.sit_name` resolves through a chain — risk-workbook GUID row > name carried in the raw AE payload > tenant GUID-to-name map (`CurrentTenantSITs.json`, the snapshot this tool auto-generates into each export run — see "SIT Reference Snapshot") > raw GUID fallback. Resolution order for the tenant map: `--sit-names <path>` CLI argument > `CurrentTenantSITs.json` in the export root (or one level below) > auto-detected `ConfigFiles/CurrentTenantSITs.json`. When a raw-payload/tenant-map name matches a workbook row by name (custom SITs live in the workbook as slugs while detections only carry GUIDs), the detection is bridged onto that workbook row, inheriting its risk metadata and exclusion behavior. Names are best-effort (a missing map is fine; GUID labels remain), but per-source resolution counts are logged and recorded under `sit_name_resolution` in `manifest.json`. Note: SIT exclusions key on the *resolved* name, so supplying a name map can correctly grow the excluded-row count.

**Sub-entity SIT names (`extract_sit_names`):** most AE detections carry Microsoft-internal classification sub-entity GUIDs that neither the flat `Get-DlpSensitiveInformationType` list nor the risk workbook know — they would fall back to GUID labels. The tenant's DLP rule packages name every classification rule GUID; the export tool resolves them automatically via the SIT Reference Snapshot above, so no extra step is needed for fresh exports. For older exports or extra sources, `py -m parquet_builder.star.extract_sit_names --input <artifact> [--input <artifact> ...]` distils name-bearing artifacts into one merged map (default output: `ConfigFiles/SITNames.local.json`, gitignored — tenant data, never track it) that you pass to the converter via `--sit-names`. Inputs are merged first-wins, so list the tenant's own `CurrentTenantSITs.json` before cross-tenant artifacts. Supported formats (content-sniffed): flat GUID-to-name map, rule-package `ClassificationRuleCollection` XML (`Data/Reference/RulePackages/*.xml`; Entity and Affinity GUIDs via LocalizedStrings, default-language name preferred), portal export artifacts (type-catalog dump with `TagRecords` groups — includes trainable classifiers; `{"Records": [{Id, Name}]}` aggregate probes; `sit_folder_index*.json`), warehouse `sit_catalog.parquet`, and any CSV with Id + DisplayName/Name columns. `--strip-prefix "QGISCF - "` (repeatable) removes a deployment-pack display prefix so pack SITs land on the risk workbook's base names and the bridge attaches their metadata.

**Org-field mapping:** how `dim_user`'s `division`/`region`/`job_title`/`is_leaver`/`is_generic_account` are sourced from the GAL is configurable (the schema itself is fixed). Copy `ConfigFiles/AEStarOrgMapping.example.json` to `AEStarOrgMapping.local.json` (gitignored) and adjust per tenant — e.g. point `Division` at `CompanyName` with a `Department` fallback. Resolution order: `--org-mapping <path>` CLI argument > auto-detected `ConfigFiles/AEStarOrgMapping.local.json` > built-in defaults (Division mirrors Department; Region/IsLeaver/IsGenericAccount derive from `OnPremisesDN` OUs and degrade to `Unknown`/false without DNs). A malformed config or a referenced column missing from the GAL is a hard error — never a silent fallback. The resolved mapping and its source are recorded in `manifest.json`.

### Building the Reports

Requires [pbi-tools.core](https://pbi.tools/) (expected at `C:\Tools\pbi-tools-net9`, run with `DOTNET_ROLL_FORWARD=Major` — the wrapper sets this) and Python with `pyarrow`.

```powershell
# Activity Explorer report (29 pages) from the star-schema output
.\PowerBI\Build-PowerBI.ps1 -Project ActivityExplorer -ParquetRoot "Output\Export-20260611-090000\PowerBI-AE-Parquet-v6"

# Content Explorer SIT risk report (15 pages) from a CE Parquet directory
.\PowerBI\Build-PowerBI.ps1 -Project ContentExplorerSITRisk -ParquetRoot "<ce-parquet-dir>"
```

The wrapper regenerates the PbixProj folder (`PowerBI/projects/<Project>/pbix/`) from the builders and compiles it to a `.pbit` template.

### Desktop Verification

1. Open the compiled `.pbit` (e.g. `PowerBI\projects\ActivityExplorer\ActivityExplorerRisk.pbit`) in Power BI Desktop.
2. If the data lives elsewhere, adjust the `ParquetRoot` parameter (Transform data > Edit parameters) and refresh.
3. Review pages and visuals; report issues against the builders, not the generated project files.

## Certificate Authentication

For unattended or multi-terminal operation, configure certificate-based auth:

1. Copy `ConfigFiles/AuthConfig.example.json` to `ConfigFiles/AuthConfig.json`
2. Fill in your Azure App registration details:

```json
{
  "UseCertificateAuth": "True",
  "AppId": "your-azure-app-id",
  "CertificateThumbprint": "YOUR_CERT_THUMBPRINT",
  "Organization": "contoso.onmicrosoft.com"
}
```

Without certificate auth, interactive (browser) authentication is used. In multi-terminal mode, each spawned worker opens a browser window requiring manual login.

## Configuration Files

| File | Purpose |
|------|---------|
| `ActivityExplorerSelector.json` | Toggle which activities and workloads to export |
| `AuthConfig.json` | Certificate auth for unattended operation (not tracked; copy from `.example`) |
| `ContentExplorerClassifiers.json` | Configure classifiers and auto-discovery settings |
| `SITstoSkip.json` | Exclude specific SITs from Content Explorer |
| `CurrentTenantSITs.json` | Optional GUID-to-Name SIT map fallback for the converter (not tracked; exports auto-generate their own snapshot at `<ExportDir>\CurrentTenantSITs.json`) |
| `CurrentTenantTCs.local.json` | Trainable-classifier name cache (not tracked; produced by the external GetTCs tool — see "Trainable Classifier Names") |
| `AEStarEnrichment.local.json` | Risk workbook + department CSV paths for `-PowerBIParquet` (not tracked; copy from `.example`) |
| `AEStarOrgMapping.local.json` | How dim_user org fields (division/region/flags) are sourced from the GAL (not tracked; copy from `.example`; built-in defaults apply without it) |
| `AEStarSITExclusions.json` | SIT names excluded from the star-schema fact and aggregate tables |

Config files use `"True"`/`"False"` strings for toggles. Properties starting with `_` are metadata and ignored by the export logic.

## Documentation Scope

`README.md` is the maintained operator-facing document.

Ad hoc root markdown files such as `progress.md`, `paginationanalysis.md`, `security_architecture_review.md`, and anything under `memory/` are working notes and design artifacts. They can be useful context, but they are not treated as authoritative runbooks and may lag behind the current code.

## Output Structure

Each export creates a timestamped subfolder under `Output/`:

```
Output/Export-YYYYMMDD-HHMMSS/
  Data/
    ContentExplorer/
      _manifest.json
      Aggregates/
        agg-*.csv
      SensitiveInformationType/
        CreditCardNumber/
          SharePoint-001.json
          OneDrive-001.json
          _task-SharePoint.json
      Sensitivity/...
      Retention/...
      TrainableClassifier/...
    ActivityExplorer/
      _manifest.json
      2026-03-15/
        Page-001.json
        Page-002.json
      2026-03-14/...
  _Coordination/
    ExportPhase.txt
    ExportType.txt
    ExportSettings.json
    RunSigningKey.txt
    AggregateTasks.csv
    DetailTasks.csv
    AEDayTasks.csv
    RunTracker.json
    Completions/
    Workers/
  _Logs/
    ExportProject-Errors.log
    ContentExplorer-Progress.log
    ActivityExplorer-Progress.log
  DLP-Config.json
  Labels-Config.json
  eDiscovery-Config.json
  RBAC-Config.json
```

## Resilience

- **Content Explorer**: Per-page file saves with run tracker for resumability. Adaptive page sizing based on volume and location distribution. Transient error retry with exponential backoff.
- **Activity Explorer**: Per-page file saves with watermark tracking. Resume from last page with `-AEResume`. Automatic retry on transient errors.
- **Multi-Terminal**: File-drop worker orchestration with stale-worker detection and automatic task reclaim for interrupted workers.
- **Token Expiry**: Certificate auth reconnects silently. Interactive auth prompts for re-login (orchestrator) or exits gracefully (workers).

## Security

- **Signed worker coordination**: Worker task payloads and completion signals are HMAC-SHA256 signed using a per-run key (`_Coordination/RunSigningKey.txt`). Tampered or unsigned payloads are quarantined.
- **Task schema validation**: Workers validate task payload schema and allowed enum values before execution.
- **CSV export hardening**: CSV output is sanitized to neutralize spreadsheet formula injection (`=`, `+`, `-`, `@` prefixes).

## License

[MIT](LICENSE)
