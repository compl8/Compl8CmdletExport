#Requires -Version 7.0
<#
.SYNOPSIS
    Compl8 Cmdlet Export Tool

.DESCRIPTION
    Exports Microsoft Purview compliance configuration including:
    - DLP policies and rules
    - Sensitivity labels and policies
    - Retention labels and policies
    - Sensitive Information Types
    - Content Explorer data
    - Activity Explorer data
    - eDiscovery cases and searches
    - RBAC role groups and assignments

    Based on current Microsoft Learn documentation:
    - https://learn.microsoft.com/powershell/exchange/connect-to-scc-powershell
    - https://learn.microsoft.com/powershell/module/exchange/export-contentexplorerdata
    - https://learn.microsoft.com/powershell/module/exchange/export-activityexplorerdata

.PARAMETER FullExport
    Export all configuration types.

.PARAMETER DlpOnly
    Export only DLP policies, rules, and SITs.

.PARAMETER LabelsOnly
    Export only sensitivity and retention labels.

.PARAMETER ContentExplorer
    Export Content Explorer data.

.PARAMETER ActivityExplorer
    Export Activity Explorer data.

.PARAMETER eDiscoveryOnly
    Export only eDiscovery configuration.

.PARAMETER RbacOnly
    Export only RBAC configuration.

.PARAMETER OutputFormat
    Output format: JSON or CSV. Default: JSON

.PARAMETER OutputDirectory
    Directory for export files. Default: ./Output

.PARAMETER UserPrincipalName
    UPN for interactive authentication.

.PARAMETER PastDays
    Days of history for Activity Explorer (1-30). Default: 7

.PARAMETER PageSize
    Records per page for paginated exports. Default: 5000

.PARAMETER CEAllSITs
    Content Explorer: Auto-discover and export ALL Sensitive Information Types from tenant.
    Overrides config file settings. Runs aggregate queries first to show progress.

.PARAMETER AEResume
    Activity Explorer: Resume from last successful page using RunTracker.json.
    Use this after a failed or interrupted export to continue from where it left off.

.EXAMPLE
    .\Export-Compl8Configuration.ps1 -FullExport
    Exports all Purview configuration with interactive authentication.

.EXAMPLE
    .\Export-Compl8Configuration.ps1 -DlpOnly -OutputFormat CSV
    Exports only DLP configuration to CSV files.

.EXAMPLE
    .\Export-Compl8Configuration.ps1 -ContentExplorer -PastDays 14
    Exports Content Explorer data.

.EXAMPLE
    .\Export-Compl8Configuration.ps1 -ActivityExplorer -PastDays 30
    Exports Activity Explorer data for the past 30 days.
#>

[CmdletBinding(DefaultParameterSetName = "Full")]
param(
    [Parameter(ParameterSetName = "Full")]
    [switch]$FullExport,

    [Parameter(ParameterSetName = "DLP")]
    [switch]$DlpOnly,

    [Parameter(ParameterSetName = "Labels")]
    [switch]$LabelsOnly,

    [Parameter(ParameterSetName = "ContentExplorer")]
    [switch]$ContentExplorer,

    [Parameter(ParameterSetName = "ActivityExplorer")]
    [switch]$ActivityExplorer,

    [Parameter(ParameterSetName = "eDiscovery")]
    [switch]$eDiscoveryOnly,

    [Parameter(ParameterSetName = "RBAC")]
    [switch]$RbacOnly,

    [Parameter(ParameterSetName = "Menu")]
    [switch]$Menu,

    [switch]$Help,

    [ValidateSet("JSON", "CSV")]
    [string]$OutputFormat = "JSON",

    [string]$OutputDirectory,

    [string]$UserPrincipalName,

    [ValidateRange(1, 30)]
    [int]$PastDays = 7,

    [ValidateRange(1, 5000)]
    [int]$PageSize = 5000,

    # Skip options for Full Export
    [Parameter(ParameterSetName = "Full")]
    [switch]$NoActivity,

    [Parameter(ParameterSetName = "Full")]
    [switch]$NoContent,

    [Parameter(ParameterSetName = "ContentExplorer")]
    [switch]$CEAllSITs,  # Auto-discover and export ALL SITs

    [Parameter(ParameterSetName = "ContentExplorer")]
    [ValidateSet("Exchange", "SharePoint", "OneDrive", "Teams")]
    [string[]]$CEWorkloads,

    [Parameter(ParameterSetName = "ContentExplorer")]
    [ValidateRange(2, 16)]
    [int]$WorkerCount,

    [Parameter(ParameterSetName = "ContentExplorer")]
    [string]$CEResumeDir,

    [Parameter(ParameterSetName = "ContentExplorer")]
    [string]$CERetryDir,

    [Parameter(ParameterSetName = "ContentExplorer")]
    [string]$CETasksCsv,

    # Activity Explorer options
    [Parameter(ParameterSetName = "ActivityExplorer")]
    [Parameter(ParameterSetName = "Full")]
    [switch]$AEResume,  # Resume from last successful page

    [Parameter(ParameterSetName = "ActivityExplorer")]
    [ValidateRange(2, 16)]
    [int]$AEWorkerCount,

    [Parameter(ParameterSetName = "ActivityExplorer")]
    [string]$AEResumeDir,

    # Worker mode for multi-terminal orchestration (receives export run directory)
    [Parameter(ParameterSetName = "Worker", Mandatory)]
    [string]$WorkerExportDir,

    [Parameter(ParameterSetName = "Worker")]
    [switch]$WorkerMode,

    # Post-export: convert JSON output to unified Parquet format
    [string]$UnifiedParquetDir,

    [string[]]$UsersCsv,

    [switch]$UnifiedParquet,

    # Tenant prefix for the export directory (e.g. 'zava' -> Export-zava-20260514-...).
    # If omitted, resolved in this order: AuthConfig.json Organization (cert auth),
    # then the last-used value from ConfigFiles/LastTenant.local.json.
    [string]$TenantPrefix
)

#region Initialization

if ($Help -or ($args -contains '--help') -or ($args -contains '-h')) {
    Write-Host "Compl8 Cmdlet Export Tool" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Usage:" -ForegroundColor Yellow
    Write-Host "  .\Export-Compl8Configuration.ps1 [mode] [options]"
    Write-Host ""
    Write-Host "Modes:" -ForegroundColor Yellow
    Write-Host "  -FullExport"
    Write-Host "  -DlpOnly"
    Write-Host "  -LabelsOnly"
    Write-Host "  -ContentExplorer"
    Write-Host "  -ActivityExplorer"
    Write-Host "  -eDiscoveryOnly"
    Write-Host "  -RbacOnly"
    Write-Host "  -Menu"
    Write-Host ""
    Write-Host "Common Options:" -ForegroundColor Yellow
    Write-Host "  -OutputFormat JSON|CSV"
    Write-Host "  -OutputDirectory <path>"
    Write-Host "  -UserPrincipalName <upn>"
    Write-Host "  -PageSize <1-5000>"
    Write-Host "  -PastDays <1-30>"
    Write-Host ""
    Write-Host "Content Explorer Options:" -ForegroundColor Yellow
    Write-Host "  -WorkerCount <2-16>"
    Write-Host "  -CEAllSITs"
    Write-Host "  -CEResumeDir <export-dir>"
    Write-Host "  -CERetryDir <export-dir>"
    Write-Host "  -CETasksCsv <tasks.csv>"
    Write-Host ""
    Write-Host "Activity Explorer Options:" -ForegroundColor Yellow
    Write-Host "  -AEResume"
    Write-Host "  -AEWorkerCount <2-16>"
    Write-Host "  -AEResumeDir <export-dir>"
    Write-Host ""
    Write-Host "Post-Export Options:" -ForegroundColor Yellow
    Write-Host "  -UnifiedParquet              Convert output to unified Parquet format"
    Write-Host "  -UnifiedParquetDir <path>    Parquet output directory (default: <export-run>\C8TuningInput)"
    Write-Host "  -UsersCsv <path>             GAL Scraper or Entra user CSV (repeatable)"
    Write-Host ""
    Write-Host "Multi-Tenant Options:" -ForegroundColor Yellow
    Write-Host "  -TenantPrefix <name>         Prefix the export directory (Export-<name>-<timestamp>)."
    Write-Host "                               If omitted: auto-detected from cert-auth Organization,"
    Write-Host "                               else the last-used value, else no prefix."
    Write-Host ""
    Write-Host "Examples:" -ForegroundColor Yellow
    Write-Host "  .\Export-Compl8Configuration.ps1 -FullExport"
    Write-Host "  .\Export-Compl8Configuration.ps1 -ContentExplorer -WorkerCount 4 -TenantPrefix zava"
    Write-Host "  .\Export-Compl8Configuration.ps1 -ActivityExplorer -PastDays 30"
    Write-Host "  .\Export-Compl8Configuration.ps1 -CEResumeDir `"Output\Export-zava-20260514-100000`""
    Write-Host ""
    Write-Host "Tip: For full built-in help use: Get-Help .\Export-Compl8Configuration.ps1 -Detailed" -ForegroundColor DarkCyan
    exit 0
}

# Capture bound parameters count at script level (before any function calls)
$script:BoundParameterCount = $PSBoundParameters.Count
Write-Verbose "Script started with $($script:BoundParameterCount) bound parameters: $($PSBoundParameters.Keys -join ', ')"

$ErrorActionPreference = "Stop"
$scriptRoot = $PSScriptRoot
# Content Explorer defaults (centralized for consistency across worker/resume/retry/export)
$script:CEDefaultBatchSize = 10
$script:CEDefaultWorkloads = @("SharePoint", "OneDrive")
$script:CEWorkerInactivityMinutes = 35

# Set default base output directory
if (-not $OutputDirectory) {
    $OutputDirectory = Join-Path $scriptRoot "Output"
}

# Import the module (must happen before any helper functions like Get-LogsDir are called)
$modulePath = Join-Path $scriptRoot "Modules\Compl8ExportFunctions.psm1"
if (-not (Test-Path $modulePath)) {
    Write-Error "Module not found: $modulePath"
    exit 1
}

try {
    Import-Module $modulePath -Force -ErrorAction Stop
}
catch {
    Write-Error "Failed to import module: $($_.Exception.Message)"
    exit 1
}

function ConvertTo-SafeTenantPrefix {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return "" }
    # Strip tenant suffixes (e.g. 'zava.onmicrosoft.com' -> 'zava'), keep alnum/dash/underscore
    $first = ($Value -split '\.', 2)[0]
    $sanitized = ($first -replace '[^a-zA-Z0-9_-]', '').Trim('-').ToLower()
    return $sanitized
}

function Resolve-TenantPrefix {
    <#
    .SYNOPSIS
        Resolves the tenant prefix used in the export directory name.
    .DESCRIPTION
        Priority:
        1. Explicit -TenantPrefix value
        2. AuthConfig.json Organization (cert auth)
        3. Last-used value from ConfigFiles/LastTenant.local.json
        Returns empty string if no source provides a value.
    #>
    param(
        [string]$ExplicitPrefix,
        [string]$ScriptRoot
    )

    $resolved = ConvertTo-SafeTenantPrefix -Value $ExplicitPrefix
    if ($resolved) { return @{ Prefix = $resolved; Source = "parameter" } }

    $authConfigPath = Join-Path $ScriptRoot "ConfigFiles" "AuthConfig.json"
    if (Test-Path $authConfigPath) {
        try {
            $authConfig = Get-Content -Path $authConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
            if ($authConfig.UseCertificateAuth -eq "True" -and -not [string]::IsNullOrWhiteSpace($authConfig.Organization)) {
                $fromAuth = ConvertTo-SafeTenantPrefix -Value $authConfig.Organization
                if ($fromAuth) { return @{ Prefix = $fromAuth; Source = "AuthConfig.json" } }
            }
        } catch { <# malformed config — ignore #> }
    }

    $lastUsedPath = Join-Path $ScriptRoot "ConfigFiles" "LastTenant.local.json"
    if (Test-Path $lastUsedPath) {
        try {
            $lastUsed = Get-Content -Path $lastUsedPath -Raw -Encoding UTF8 | ConvertFrom-Json
            if (-not [string]::IsNullOrWhiteSpace($lastUsed.TenantPrefix)) {
                $fromLast = ConvertTo-SafeTenantPrefix -Value $lastUsed.TenantPrefix
                if ($fromLast) { return @{ Prefix = $fromLast; Source = "last run" } }
            }
        } catch { <# malformed — ignore #> }
    }

    return @{ Prefix = ""; Source = "none" }
}

function Save-LastTenantPrefix {
    param([string]$ScriptRoot, [string]$Prefix)
    if ([string]::IsNullOrWhiteSpace($Prefix)) { return }
    $lastUsedPath = Join-Path $ScriptRoot "ConfigFiles" "LastTenant.local.json"
    try {
        @{ TenantPrefix = $Prefix; UpdatedAt = (Get-Date).ToString("o") } |
            ConvertTo-Json | Set-Content -Path $lastUsedPath -Encoding UTF8
    } catch { <# best effort #> }
}

if ($WorkerExportDir) {
    # Worker mode: use the orchestrator's export directory, don't create a new one
    $script:ExportRunDirectory = $WorkerExportDir
    $script:ErrorLogPath = Join-Path (Get-LogsDir $WorkerExportDir) "ExportProject-Errors.log"
} else {
    # Orchestrator/interactive mode: create new export directory
    if (-not (Test-Path $OutputDirectory)) {
        New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
    }
    $script:ExportTimestamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $prefixResult = Resolve-TenantPrefix -ExplicitPrefix $TenantPrefix -ScriptRoot $scriptRoot
    $script:TenantPrefix = $prefixResult.Prefix
    $dirName = if ($script:TenantPrefix) { "Export-{0}-{1}" -f $script:TenantPrefix, $script:ExportTimestamp } else { "Export-$script:ExportTimestamp" }
    $script:ExportRunDirectory = Join-Path $OutputDirectory $dirName
    New-Item -ItemType Directory -Force -Path $script:ExportRunDirectory | Out-Null
    $script:ErrorLogPath = Join-Path (Get-LogsDir $script:ExportRunDirectory) "ExportProject-Errors.log"
    if ($script:TenantPrefix) {
        Save-LastTenantPrefix -ScriptRoot $scriptRoot -Prefix $script:TenantPrefix
    }
}

function Resolve-UnifiedParquetOutputDir {
    param(
        [string]$ConfiguredPath,
        [string]$ExportRunDirectory
    )

    if (-not [string]::IsNullOrWhiteSpace($ConfiguredPath)) {
        return $ConfiguredPath
    }

    return Join-Path $ExportRunDirectory "C8TuningInput"
}

# Initialize logging in the export run directory
$logFile = Initialize-ExportLog -LogDirectory (Get-LogsDir $script:ExportRunDirectory) -Prefix "ExportLog"

#endregion


$scriptPartRoot = Join-Path $PSScriptRoot "App"
$scriptPartFiles = @(
    'Host\Menu.ps1'
    'Providers\Exports.Core.ps1'
    'Orchestrator\ContentExplorer.ps1'
    'Orchestrator\ActivityExplorer.ps1'
    'Providers\Exports.Other.ps1'
    'MainExecution.ps1'
)

foreach ($part in $scriptPartFiles) {
    $partPath = Join-Path $scriptPartRoot $part
    if (-not (Test-Path $partPath)) {
        throw "Script section not found: $partPath"
    }

    . $partPath
}
