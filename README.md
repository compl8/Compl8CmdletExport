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
README.md
```

`Export-Compl8Configuration.ps1` expects `App/`, `Modules/`, and `ConfigFiles/` to sit beside it. Do not move individual files out of that layout.

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
| `-UnifiedParquetDir` | Parquet output directory (default: `C:\PurviewData`) |
| `-UsersCsv` | GAL Scraper or Entra user CSV for enrichment (repeatable) |

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

Or use the interactive menu (options 5-10) to configure Content Explorer modes.

### Aggregate Caching

On startup, the tool checks for recent aggregate CSV files (last 30 days) and offers to reuse them. This skips the aggregate query phase, which can take 30+ minutes on large tenants.

### SIT Filtering

`ConfigFiles/SITstoSkip.json` controls which Sensitive Information Types are excluded from Content Explorer exports. Set a SIT to `"True"` to skip it, `"False"` to include it.

On each export, `ConfigFiles/CurrentTenantSITs.json` is generated with a GUID-to-Name mapping of all tenant SITs. Copy names from this file into `SITstoSkip.json` to exclude them.

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
# Automatic conversion after export
.\Export-Compl8Configuration.ps1 -FullExport -UnifiedParquet

# With custom output directory and user enrichment
.\Export-Compl8Configuration.ps1 -ActivityExplorer -UnifiedParquet -UnifiedParquetDir "D:\PurviewData" -UsersCsv "users.csv"
```

### Standalone Python

```bash
pip install pyarrow

# Convert an export
python build_unified_parquet.py --input-dir Output/Export-20260301-090000 --output-dir C:/PurviewData

# With user enrichment (GAL Scraper or Entra CSV)
python build_unified_parquet.py --input-dir Output/Export-20260301-090000 --output-dir C:/PurviewData --users-csv users.csv
```

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
| `CurrentTenantSITs.json` | Auto-generated GUID-to-Name SIT mapping (not tracked) |

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
