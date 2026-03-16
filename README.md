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
# (also supported: -Help)

# Interactive menu
.\Export-Compl8Configuration.ps1

# Full export
.\Export-Compl8Configuration.ps1 -FullExport

# Full export, skipping slow Content/Activity Explorer
.\Export-Compl8Configuration.ps1 -FullExport -NoActivity -NoContent
```

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

### Additional Parameters

| Parameter | Range | Description |
|-----------|-------|-------------|
| `-PastDays` | 1-30 | Activity Explorer history window |
| `-PageSize` | 1-5000 | API pagination size |
| `-OutputFormat` | JSON/CSV | Output file format (default: JSON) |
| `-OutputDirectory` | path | Custom output location (default: `./Output`) |
| `-UserPrincipalName` | UPN | Pre-fill authentication identity |
| `-SkipContentCombine` | switch | Skip Content Explorer combined dedup file generation |
| `-SkipActivityCombine` | switch | Skip Activity Explorer combined file generation |
| `-NoActivity` | switch | Skip Activity Explorer in full export |
| `-NoContent` | switch | Skip Content Explorer in full export |

## Content Explorer

Content Explorer supports single-terminal and multi-terminal parallel export modes.

### Single Terminal

```powershell
.\Export-Compl8Configuration.ps1 -ContentExplorer
```

### Multi-Terminal (Parallel)

Spawns worker terminals that process tasks in parallel, coordinated through a file-drop protocol (`nexttask`/`currenttask` plus completion signal files in `Completions/`).

```powershell
# Spawn 4 parallel workers
.\Export-Compl8Configuration.ps1 -ContentExplorer -WorkerCount 4 -SkipContentCombine

# Resume a previous export (works even if all terminals crashed)
.\Export-Compl8Configuration.ps1 -CEResumeDir "Output\Export-20260130-100000"
```

Or use the interactive menu (option 7) to configure multi-terminal mode.

### Aggregate Caching

On startup, the tool checks for recent aggregate CSV files (last 30 days) and offers to reuse them. This skips the aggregate query phase, which can take 30+ minutes on large tenants.

### SIT Filtering

`ConfigFiles/SITstoSkip.json` controls which Sensitive Information Types are excluded from Content Explorer exports. Set a SIT to `"True"` to skip it, `"False"` to include it. This filter applies to all Content Explorer modes.

On each export, `ConfigFiles/CurrentTenantSITs.json` is generated with a GUID-to-Name mapping of all tenant SITs. Copy names from this file into `SITstoSkip.json` to exclude them.

## Activity Explorer

```powershell
# Export last 7 days (default)
.\Export-Compl8Configuration.ps1 -ActivityExplorer -SkipActivityCombine

# Export last 30 days
.\Export-Compl8Configuration.ps1 -ActivityExplorer -PastDays 30

# Resume after failure
.\Export-Compl8Configuration.ps1 -ActivityExplorer -AEResume
```

## Shard-Only Export (No Combined Files)

Use these switches when you want smaller JSON shards for BI/ETL imports:

```powershell
# CE shard-only
.\Export-Compl8Configuration.ps1 -ContentExplorer -WorkerCount 4 -SkipContentCombine

# AE shard-only
.\Export-Compl8Configuration.ps1 -ActivityExplorer -PastDays 30 -SkipActivityCombine

# Full export, shard-only for both explorers
.\Export-Compl8Configuration.ps1 -FullExport -SkipContentCombine -SkipActivityCombine
```

### Merging Exports

Deduplication across multiple exports is handled by the parquet pipeline:

```bash
# Append a second export (deduplicates on RecordIdentity automatically)
python PowerBI/build_powerbi_dataset.py "Output/Export-20260301" --dataset "Output/PowerBI_Dataset"
```

### Monitoring Progress

```powershell
# Tail Content Explorer progress (in another terminal)
Get-Content -Path "Output\Export-*\ContentExplorer-Progress.log" -Wait -Tail 20

# Tail Activity Explorer progress
Get-Content -Path "Output\Export-*\ActivityExplorer\Progress.log" -Wait -Tail 20
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

## Running with Signing

No new command-line flag is required. Signing is automatic for file-drop worker coordination.

```powershell
# Content Explorer multi-terminal (signed worker task/signal files enabled automatically)
.\Export-Compl8Configuration.ps1 -ContentExplorer -WorkerCount 4

# Activity Explorer multi-terminal
.\Export-Compl8Configuration.ps1 -ActivityExplorer -AEWorkerCount 4
```

- A per-run key file is created at `Output/Export-*/RunSigningKey.txt`.
- Worker `nexttask` payloads and completion/error signal files are signed and verified automatically.
- If a task payload is malformed or tampered, the worker quarantines it to `invalidtask-*.json` and continues polling.
- If you resume an older in-progress export and see signature warnings, clear stale unsigned signal files in `Completions/` and rerun resume.

## Configuration Files

| File | Purpose |
|------|---------|
| `ActivityExplorerSelector.json` | Toggle which activities and workloads to export |
| `AuthConfig.json` | Certificate auth for unattended operation (optional) |
| `ContentExplorerClassifiers.json` | Configure classifiers and auto-discovery settings |
| `SITstoSkip.json` | Exclude specific SITs from Content Explorer |
| `CurrentTenantSITs.json` | Auto-generated GUID-to-Name SIT mapping |

Config files use `"True"`/`"False"` strings for toggles. Properties starting with `_` are metadata and ignored by the export logic.

## Output Structure

Each export creates a timestamped subfolder under `Output/`:

```
Output/Export-20260130-100000/
    ExportPhase.txt                 # Current phase marker
    ExportType.txt                  # CE vs AE worker type marker
    RunSigningKey.txt               # Per-run HMAC key for worker task/signal integrity
    AggregateTasks.csv              # CE aggregate task state (when CE runs)
    DetailTasks.csv                 # CE detail task state (when CE runs)
    AEDayTasks.csv                  # AE day task state (when AE runs)
    Completions/                    # Worker completion/error signal files
    ExportProject-Errors.log        # Detailed error log
    ContentExplorer-Progress.log    # Tailable progress
    ContentExplorer-Combined.json   # Optional (omitted with -SkipContentCombine)
    ActivityExplorer/               # Per-page AE files
    ActivityExplorer-Combined.json  # Optional (omitted with -SkipActivityCombine)
    DLP-Config.json                 # DLP policies and rules
    Labels-Config.json              # Sensitivity and retention labels
    eDiscovery-Config.json          # eDiscovery cases
    RBAC-Config.json                # Role groups and members
```

## Resilience

- **Content Explorer**: Saves per-batch with run tracker for resumability. Adaptive page sizing based on volume and location distribution. Transient error retry with exponential backoff.
- **Activity Explorer**: Saves per-page immediately. Resume from last watermark with `-AEResume`. Automatic retry on transient errors.
- **Multi-Terminal**: File-drop worker orchestration with stale-worker detection and task reclaim for interrupted workers.
- **Token Expiry**: Certificate auth reconnects silently. Interactive auth prompts for re-login (orchestrator) or exits gracefully (workers).

## Security Notes

- **CSV export hardening**: CSV output is sanitized to neutralize spreadsheet formula injection (`=`, `+`, `-`, `@` prefixes).
- **Signed worker control files**: Worker tasks and completion/error signals are wrapped in signed envelopes (HMAC-SHA256) using a per-run key (`RunSigningKey.txt`).
- **Task schema validation**: Workers validate task payload schema and allowed enum values before execution.
- **Malformed task quarantine**: Invalid/tampered `nexttask` payloads are moved to `invalidtask-*.json` so workers do not get stuck in a busy state.

## Importing Shards (PowerBI / Scripts)

- **Activity Explorer shards**: `Output/Export-*/ActivityExplorer/Page-*.json` (single terminal) or `Output/Export-*/Worker-*/AE-Day-*/Page-*.json` (multi-terminal).
- **Content Explorer shards**: `Output/Export-*/Worker-*/detail-*.json` (multi-terminal) and/or `Output/Export-*/ContentExplorer-*.json` (single-terminal/retry runs).
- Combined files are optional convenience outputs and can be disabled with `-SkipContentCombine` and `-SkipActivityCombine`.
