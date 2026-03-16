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

    [switch]$UnifiedParquet
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
    Write-Host "  -UnifiedParquetDir <path>    Parquet output directory (default: C:\PurviewData)"
    Write-Host "  -UsersCsv <path>             GAL Scraper or Entra user CSV (repeatable)"
    Write-Host ""
    Write-Host "Examples:" -ForegroundColor Yellow
    Write-Host "  .\Export-Compl8Configuration.ps1 -FullExport"
    Write-Host "  .\Export-Compl8Configuration.ps1 -ContentExplorer -WorkerCount 4"
    Write-Host "  .\Export-Compl8Configuration.ps1 -ActivityExplorer -PastDays 30"
    Write-Host "  .\Export-Compl8Configuration.ps1 -CEResumeDir `"Output\Export-20260130-100000`""
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
    $script:ExportRunDirectory = Join-Path $OutputDirectory "Export-$script:ExportTimestamp"
    New-Item -ItemType Directory -Force -Path $script:ExportRunDirectory | Out-Null
    $script:ErrorLogPath = Join-Path (Get-LogsDir $script:ExportRunDirectory) "ExportProject-Errors.log"
}

# Import the module
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

# Initialize logging in the export run directory
$logFile = Initialize-ExportLog -LogDirectory (Get-LogsDir $script:ExportRunDirectory) -Prefix "ExportLog"

#endregion

#region Interactive Menu

function Show-ExportMenu {
    <#
    .SYNOPSIS
        Shows an interactive menu for selecting export options.
    .OUTPUTS
        Hashtable with selected options including Mode, NoActivity, NoContent,
        PastDays, CEAllSITs, CEMultiTerminal, CEWorkerCount, CEResumePath,
        CERetryPath, OutputFormat, and Quit.
    #>
    [CmdletBinding()]
    param()

    # Initialize result
    $result = @{
        Mode             = $null
        NoActivity       = $false
        NoContent        = $false
        PastDays         = 7
        CEAllSITs        = $false
        OutputFormat     = "JSON"
        Quit             = $false
        CEMultiTerminal  = $false
        CEWorkerCount    = 0
        CEResumePath     = $null
        CERetryPath      = $null
        CETasksCsvPath   = $null
        CEWorkloads      = $null
        AEMultiTerminal  = $false
        AEWorkerCount    = 0
        AEResumePath     = $null
    }

    do {
        # Reset mode each iteration (so option 8 with no selection re-shows menu)
        $result.Mode = $null

        # Clear screen if running interactively (skip if non-interactive to avoid errors)
        if ($Host.UI.RawUI.CursorPosition) {
            try { Clear-Host } catch { <# Non-interactive host may not support Clear-Host #> }
        }

        Write-Host ""
        $w = Get-BoxInnerWidth -MaxWidth 62
        Write-BoxTop -InnerWidth $w -Color Cyan -Single
        Write-BoxLine -Text "       Compl8 Cmdlet Export" -InnerWidth $w -Color Cyan -Single
        Write-BoxSeparator -InnerWidth $w -Color Cyan -Single
        Write-BoxLine -InnerWidth $w -Color Cyan -Single
        Write-BoxLine -Text "FULL EXPORT" -InnerWidth $w -Color Cyan -Single
        Write-BoxLine -Text "  [1]  Full Export (all data types)" -InnerWidth $w -Color Cyan -Single
        Write-BoxLine -Text "  [2]  Full - Skip Activity Explorer" -InnerWidth $w -Color Cyan -Single
        Write-BoxLine -Text "  [3]  Full - Skip Content Explorer" -InnerWidth $w -Color Cyan -Single
        Write-BoxLine -Text "  [4]  Full - Skip Both Explorers" -InnerWidth $w -Color Cyan -Single
        Write-BoxSeparator -InnerWidth $w -Color Cyan -Single
        Write-BoxLine -Text "CONTENT EXPLORER" -InnerWidth $w -Color Cyan -Single
        Write-BoxLine -Text "  [5]  Content Explorer (from config)" -InnerWidth $w -Color Cyan -Single
        Write-BoxLine -Text "  [6]  Content Explorer - All SITs (full tenant scan)" -InnerWidth $w -Color Cyan -Single
        Write-BoxLine -Text "  [7]  Content Explorer - Multi-Terminal (parallel)" -InnerWidth $w -Color Cyan -Single
        Write-BoxLine -Text "  [8]  Content Explorer - Resume Previous Export" -InnerWidth $w -Color Cyan -Single
        Write-BoxLine -Text "  [9]  Content Explorer - Retry Discrepant Tasks" -InnerWidth $w -Color Cyan -Single
        Write-BoxLine -Text "  [10] Content Explorer - Run from Task CSV" -InnerWidth $w -Color Cyan -Single
        Write-BoxLine -Text "       * Filtered by SITstoSkip.json (all CE modes)" -InnerWidth $w -Color DarkCyan -Single
        Write-BoxSeparator -InnerWidth $w -Color Cyan -Single
        Write-BoxLine -Text "ACTIVITY EXPLORER" -InnerWidth $w -Color Cyan -Single
        Write-BoxLine -Text "  [11] Activity Explorer" -InnerWidth $w -Color Cyan -Single
        Write-BoxLine -Text "  [12] Activity Explorer - Multi-Terminal (parallel)" -InnerWidth $w -Color Cyan -Single
        Write-BoxLine -Text "  [13] Activity Explorer - Resume Previous Export" -InnerWidth $w -Color Cyan -Single
        Write-BoxSeparator -InnerWidth $w -Color Cyan -Single
        Write-BoxLine -Text "OTHER EXPORTS" -InnerWidth $w -Color Cyan -Single
        Write-BoxLine -Text "  [14] DLP Policies & Rules" -InnerWidth $w -Color Cyan -Single
        Write-BoxLine -Text "  [15] Sensitivity & Retention Labels" -InnerWidth $w -Color Cyan -Single
        Write-BoxLine -Text "  [16] eDiscovery Cases" -InnerWidth $w -Color Cyan -Single
        Write-BoxLine -Text "  [17] RBAC Configuration" -InnerWidth $w -Color Cyan -Single
        Write-BoxLine -InnerWidth $w -Color Cyan -Single
        Write-BoxLine -Text "  [Q]  Quit" -InnerWidth $w -Color Cyan -Single
        Write-BoxLine -InnerWidth $w -Color Cyan -Single
        Write-BoxBottom -InnerWidth $w -Color Cyan -Single
        Write-Host ""

        $selection = Read-Host "  Enter selection [1-17, Q]"

        if ([string]::IsNullOrEmpty($selection)) { $selection = "" }
        $selectionUpper = $selection.Trim().ToUpper()

        switch ($selectionUpper) {
            "1" { $result.Mode = "Full" }
            "2" { $result.Mode = "Full"; $result.NoActivity = $true }
            "3" { $result.Mode = "Full"; $result.NoContent = $true }
            "4" { $result.Mode = "Full"; $result.NoContent = $true; $result.NoActivity = $true }
            "5" {
                $result.Mode = "ContentExplorer"
                $result.CEWorkloads = Get-CEWorkloadSelection
            }
            "6" {
                $result.Mode = "ContentExplorer"
                $result.CEAllSITs = $true
                $result.CEWorkloads = Get-CEWorkloadSelection
            }
            "7" {
                # Multi-Terminal Content Explorer
                $result.Mode = "ContentExplorer"
                $result.CEMultiTerminal = $true

                # Worker count prompt
                Write-Host ""
                Write-Host "  How many worker terminals?" -ForegroundColor Yellow
                $workerInput = Read-Host "  Enter count [4]"
                if ([string]::IsNullOrEmpty($workerInput)) { $workerInput = "4" }
                $workerCount = $workerInput -as [int]
                if (-not $workerCount -or $workerCount -lt 2 -or $workerCount -gt 16) {
                    Write-Host "  Invalid count. Must be 2-16. Using default of 4." -ForegroundColor Yellow
                    $workerCount = 4
                }
                $result.CEWorkerCount = $workerCount

                # All SITs prompt
                Write-Host ""
                $allSitsInput = Read-Host "  Scan ALL Sensitive Info Types? [Y/n]"
                Write-Host "    * Filtered by SITstoSkip.json (excluded SITs are always skipped)" -ForegroundColor DarkCyan
                if ([string]::IsNullOrEmpty($allSitsInput) -or $allSitsInput.Trim().ToUpper() -ne "N") {
                    $result.CEAllSITs = $true
                }

                # Auth mode detection
                $authParams = Build-AuthParameters
                $isCertAuth = $authParams.ContainsKey('AppId')
                if ($isCertAuth) {
                    Write-Host ""
                    Write-Host "  Authentication: Certificate (unattended)" -ForegroundColor Green
                    Write-Host "  AppId: $($authParams.AppId.Substring(0, 8))..." -ForegroundColor Gray
                    Write-Host "  Workers will authenticate automatically." -ForegroundColor Green
                } else {
                    Write-Host ""
                    Write-Host "  Authentication: Interactive (browser login)" -ForegroundColor Yellow
                    Write-Host "  NOTE: Each spawned terminal will open a browser window." -ForegroundColor Yellow
                    Write-Host "  You must authenticate in each window manually." -ForegroundColor Yellow
                }

                $result.CEWorkloads = Get-CEWorkloadSelection
            }
            "8" {
                # Resume Previous Export - scan for ExportPhase.txt in Export-* directories
                $baseOutputDir = if ($OutputDirectory) { $OutputDirectory } else { Join-Path $PSScriptRoot "Output" }
                $resumableDirs = @()
                if (Test-Path $baseOutputDir) {
                    $exportFolders = Get-ChildItem -Path $baseOutputDir -Directory -Filter "Export-*" -ErrorAction SilentlyContinue |
                        Sort-Object LastWriteTime -Descending
                    foreach ($folder in $exportFolders) {
                        $phasePath = Join-Path (Get-CoordinationDir $folder.FullName) "ExportPhase.txt"
                        if (Test-Path $phasePath) {
                            $phase = ([System.IO.File]::ReadAllText($phasePath)).Trim()
                            if ($phase -notin @("Completed")) {
                                $resumableDirs += [PSCustomObject]@{
                                    ExportDir = $folder.FullName
                                    DirName = $folder.Name
                                    Phase = $phase
                                    LastWrite = $folder.LastWriteTime
                                }
                            }
                        }
                    }
                }

                if ($resumableDirs.Count -eq 0) {
                    Write-Host ""
                    Write-Host "  No resumable exports found in Output directory." -ForegroundColor Yellow
                    Write-Host "  Press Enter to return to menu..." -ForegroundColor Gray
                    $null = Read-Host
                } else {
                    Write-Host ""
                    Write-SectionHeader -Text "Resumable Exports Found" -Color Cyan
                    for ($i = 0; $i -lt $resumableDirs.Count; $i++) {
                        $rd = $resumableDirs[$i]
                        $elapsed = (Get-Date) - $rd.LastWrite
                        $agoText = if ($elapsed.TotalHours -lt 1) { "{0}m ago" -f [int]$elapsed.TotalMinutes }
                                   elseif ($elapsed.TotalHours -lt 24) { "{0}h ago" -f [math]::Round($elapsed.TotalHours, 1) }
                                   else { "{0}d ago" -f [math]::Round($elapsed.TotalDays, 1) }

                        Write-Host ("  [{0}] {1}" -f ($i+1), $rd.DirName) -ForegroundColor White
                        Write-Host ("      Phase: {0} | Last activity: {1}" -f $rd.Phase, $agoText) -ForegroundColor Gray
                    }
                    Write-Host ""
                    $resumeInput = Read-Host ("  Select export to resume [1-{0}, N for new export]" -f $resumableDirs.Count)

                    if (-not [string]::IsNullOrEmpty($resumeInput)) {
                        $resumeIndex = ($resumeInput -as [int])
                        if ($resumeIndex -and $resumeIndex -ge 1 -and $resumeIndex -le $resumableDirs.Count) {
                            $result.Mode = "ContentExplorerResume"
                            $result.CEResumePath = $resumableDirs[$resumeIndex - 1].ExportDir

                            # Ask for worker count
                            Write-Host ""
                            Write-Host "  Workers: [1] Single terminal (default)  [2-16] Multi-terminal with N workers" -ForegroundColor Gray
                            $wcInput = Read-Host "  Number of workers"
                            $wcVal = $wcInput -as [int]
                            if ($wcVal -and $wcVal -ge 2 -and $wcVal -le 16) {
                                $result.CEWorkerCount = $wcVal
                            }
                            else {
                                $result.CEWorkerCount = 0
                            }
                        }
                    }
                }
            }
            "9" {
                # Retry Discrepant Tasks - scan for RetryTasks.csv in Export-* directories
                $baseOutputDir = if ($OutputDirectory) { $OutputDirectory } else { Join-Path $PSScriptRoot "Output" }
                $retryableDirs = @()
                if (Test-Path $baseOutputDir) {
                    $exportFolders = Get-ChildItem -Path $baseOutputDir -Directory -Filter "Export-*" -ErrorAction SilentlyContinue |
                        Sort-Object LastWriteTime -Descending
                    foreach ($folder in $exportFolders) {
                        $retryPath = Join-Path (Get-CoordinationDir $folder.FullName) "RetryTasks.csv"
                        if (Test-Path $retryPath) {
                            $retryTaskCount = @(Import-Csv -Path $retryPath -Encoding UTF8).Count
                            $retryableDirs += [PSCustomObject]@{
                                ExportDir = $folder.FullName
                                DirName   = $folder.Name
                                TaskCount = $retryTaskCount
                                LastWrite = $folder.LastWriteTime
                            }
                        }
                    }
                }

                if ($retryableDirs.Count -eq 0) {
                    Write-Host ""
                    Write-Host "  No exports with retry tasks found in Output directory." -ForegroundColor Yellow
                    Write-Host "  Press Enter to return to menu..." -ForegroundColor Gray
                    $null = Read-Host
                } else {
                    Write-Host ""
                    Write-SectionHeader -Text "Exports with Retry Tasks" -Color Cyan
                    for ($i = 0; $i -lt $retryableDirs.Count; $i++) {
                        $rd = $retryableDirs[$i]
                        $elapsed = (Get-Date) - $rd.LastWrite
                        $agoText = if ($elapsed.TotalHours -lt 1) { "{0}m ago" -f [int]$elapsed.TotalMinutes }
                                   elseif ($elapsed.TotalHours -lt 24) { "{0}h ago" -f [math]::Round($elapsed.TotalHours, 1) }
                                   else { "{0}d ago" -f [math]::Round($elapsed.TotalDays, 1) }

                        Write-Host ("  [{0}] {1}" -f ($i+1), $rd.DirName) -ForegroundColor White
                        Write-Host ("      Retry tasks: {0} | Last activity: {1}" -f $rd.TaskCount, $agoText) -ForegroundColor Gray
                    }
                    Write-Host ""
                    $retryInput = Read-Host ("  Select export to retry [1-{0}, N to cancel]" -f $retryableDirs.Count)

                    if (-not [string]::IsNullOrEmpty($retryInput)) {
                        $retryIndex = ($retryInput -as [int])
                        if ($retryIndex -and $retryIndex -ge 1 -and $retryIndex -le $retryableDirs.Count) {
                            $result.Mode = "ContentExplorerRetry"
                            $result.CERetryPath = $retryableDirs[$retryIndex - 1].ExportDir
                        }
                    }
                }
            }
            "10" {
                # Run from Task CSV - scan for RemainingTasks.csv or accept a path
                Write-Host ""
                Write-Host "  Enter path to a task CSV file, or press Enter to scan Output directory:" -ForegroundColor Yellow
                $csvInput = Read-Host "  CSV path"

                if ([string]::IsNullOrEmpty($csvInput)) {
                    # Scan for RemainingTasks.csv in Output directories
                    $baseOutputDir = if ($OutputDirectory) { $OutputDirectory } else { Join-Path $PSScriptRoot "Output" }
                    $taskCsvDirs = @()
                    if (Test-Path $baseOutputDir) {
                        $exportFolders = Get-ChildItem -Path $baseOutputDir -Directory -Filter "Export-*" -ErrorAction SilentlyContinue |
                            Sort-Object LastWriteTime -Descending
                        foreach ($folder in $exportFolders) {
                            $remainingPath = Join-Path $folder.FullName "RemainingTasks.csv"
                            if (Test-Path $remainingPath) {
                                $taskCount = @(Import-Csv -Path $remainingPath -Encoding UTF8).Count
                                $taskCsvDirs += [PSCustomObject]@{
                                    CsvPath   = $remainingPath
                                    DirName   = $folder.Name
                                    TaskCount = $taskCount
                                    LastWrite = $folder.LastWriteTime
                                }
                            }
                        }
                    }

                    if ($taskCsvDirs.Count -eq 0) {
                        Write-Host ""
                        Write-Host "  No RemainingTasks.csv files found in Output directory." -ForegroundColor Yellow
                        Write-Host "  Press Enter to return to menu..." -ForegroundColor Gray
                        $null = Read-Host
                    } else {
                        Write-Host ""
                        Write-SectionHeader -Text "Exports with Remaining Tasks" -Color Cyan
                        for ($i = 0; $i -lt $taskCsvDirs.Count; $i++) {
                            $td = $taskCsvDirs[$i]
                            $elapsed = (Get-Date) - $td.LastWrite
                            $agoText = if ($elapsed.TotalHours -lt 1) { "{0}m ago" -f [int]$elapsed.TotalMinutes }
                                       elseif ($elapsed.TotalHours -lt 24) { "{0}h ago" -f [math]::Round($elapsed.TotalHours, 1) }
                                       else { "{0}d ago" -f [math]::Round($elapsed.TotalDays, 1) }

                            Write-Host ("  [{0}] {1}" -f ($i+1), $td.DirName) -ForegroundColor White
                            Write-Host ("      Remaining tasks: {0} | Last activity: {1}" -f $td.TaskCount, $agoText) -ForegroundColor Gray
                        }
                        Write-Host ""
                        $pickInput = Read-Host ("  Select [1-{0}, N to cancel]" -f $taskCsvDirs.Count)

                        if (-not [string]::IsNullOrEmpty($pickInput)) {
                            $pickIndex = ($pickInput -as [int])
                            if ($pickIndex -and $pickIndex -ge 1 -and $pickIndex -le $taskCsvDirs.Count) {
                                $result.Mode = "ContentExplorerTasksCsv"
                                $result.CETasksCsvPath = $taskCsvDirs[$pickIndex - 1].CsvPath
                            }
                        }
                    }
                }
                else {
                    # User provided a path
                    $csvPath = $csvInput.Trim().Trim('"')
                    if (Test-Path $csvPath) {
                        $result.Mode = "ContentExplorerTasksCsv"
                        $result.CETasksCsvPath = $csvPath
                    } else {
                        Write-Host ("  File not found: {0}" -f $csvPath) -ForegroundColor Red
                        Write-Host "  Press Enter to return to menu..." -ForegroundColor Gray
                        $null = Read-Host
                    }
                }

                # Ask for worker count if a CSV was selected
                if ($result.CETasksCsvPath) {
                    Write-Host ""
                    Write-Host "  Workers: [1] Single terminal (default)  [2-16] Multi-terminal with N workers" -ForegroundColor Gray
                    $wcInput = Read-Host "  Number of workers"
                    $wcVal = $wcInput -as [int]
                    if ($wcVal -and $wcVal -ge 2 -and $wcVal -le 16) {
                        $result.CEWorkerCount = $wcVal
                    }
                    else {
                        $result.CEWorkerCount = 0
                    }

                    $result.CEWorkloads = Get-CEWorkloadSelection
                }
            }
            "11" { $result.Mode = "ActivityExplorer" }
            "12" {
                # Multi-Terminal Activity Explorer
                $result.Mode = "ActivityExplorer"
                $result.AEMultiTerminal = $true

                # Worker count prompt
                Write-Host ""
                Write-Host "  How many worker terminals?" -ForegroundColor Yellow
                $workerInput = Read-Host "  Enter count [4]"
                if ([string]::IsNullOrEmpty($workerInput)) { $workerInput = "4" }
                $aeWc = $workerInput -as [int]
                if (-not $aeWc -or $aeWc -lt 2 -or $aeWc -gt 16) {
                    Write-Host "  Invalid count. Must be 2-16. Using default of 4." -ForegroundColor Yellow
                    $aeWc = 4
                }
                $result.AEWorkerCount = $aeWc

                # Auth mode detection
                $authParams = Build-AuthParameters
                $isCertAuth = $authParams.ContainsKey('AppId')
                if ($isCertAuth) {
                    Write-Host ""
                    Write-Host "  Authentication: Certificate (unattended)" -ForegroundColor Green
                    Write-Host "  Workers will authenticate automatically." -ForegroundColor Green
                } else {
                    Write-Host ""
                    Write-Host "  Authentication: Interactive (browser login)" -ForegroundColor Yellow
                    Write-Host "  NOTE: Each spawned terminal will open a browser window." -ForegroundColor Yellow
                    Write-Host "  You must authenticate in each window manually." -ForegroundColor Yellow
                }
            }
            "13" {
                # Resume Previous AE Export - scan for ExportType.txt = "ActivityExplorer" + incomplete phase
                $baseOutputDir = if ($OutputDirectory) { $OutputDirectory } else { Join-Path $PSScriptRoot "Output" }
                $aeResumableDirs = @()
                if (Test-Path $baseOutputDir) {
                    $exportFolders = Get-ChildItem -Path $baseOutputDir -Directory -Filter "Export-*" -ErrorAction SilentlyContinue |
                        Sort-Object LastWriteTime -Descending
                    foreach ($folder in $exportFolders) {
                        $coordDir = Get-CoordinationDir $folder.FullName
                        $typePath = Join-Path $coordDir "ExportType.txt"
                        $phasePath = Join-Path $coordDir "ExportPhase.txt"
                        if ((Test-Path $typePath) -and (Test-Path $phasePath)) {
                            $exportType = ([System.IO.File]::ReadAllText($typePath)).Trim()
                            $phase = ([System.IO.File]::ReadAllText($phasePath)).Trim()
                            if ($exportType -eq "ActivityExplorer" -and $phase -notin @("AECompleted")) {
                                $taskPath = Join-Path $coordDir "AEDayTasks.csv"
                                $taskCount = 0
                                $pendingCount = 0
                                if (Test-Path $taskPath) {
                                    $tasks = @(Import-Csv -Path $taskPath -Encoding UTF8)
                                    $taskCount = $tasks.Count
                                    $pendingCount = @($tasks | Where-Object { $_.Status -in @("Pending","Error","InProgress") }).Count
                                }
                                $aeResumableDirs += [PSCustomObject]@{
                                    ExportDir    = $folder.FullName
                                    DirName      = $folder.Name
                                    Phase        = $phase
                                    TaskCount    = $taskCount
                                    PendingCount = $pendingCount
                                    LastWrite    = $folder.LastWriteTime
                                }
                            }
                        }
                    }
                }

                if ($aeResumableDirs.Count -eq 0) {
                    Write-Host ""
                    Write-Host "  No resumable Activity Explorer exports found in Output directory." -ForegroundColor Yellow
                    Write-Host "  Press Enter to return to menu..." -ForegroundColor Gray
                    $null = Read-Host
                } else {
                    Write-Host ""
                    Write-SectionHeader -Text "Resumable Activity Explorer Exports" -Color Cyan
                    for ($i = 0; $i -lt $aeResumableDirs.Count; $i++) {
                        $rd = $aeResumableDirs[$i]
                        $elapsed = (Get-Date) - $rd.LastWrite
                        $agoText = if ($elapsed.TotalHours -lt 1) { "{0}m ago" -f [int]$elapsed.TotalMinutes }
                                   elseif ($elapsed.TotalHours -lt 24) { "{0}h ago" -f [math]::Round($elapsed.TotalHours, 1) }
                                   else { "{0}d ago" -f [math]::Round($elapsed.TotalDays, 1) }

                        Write-Host ("  [{0}] {1}" -f ($i+1), $rd.DirName) -ForegroundColor White
                        Write-Host ("      Phase: {0} | {1}/{2} tasks remaining | Last activity: {3}" -f $rd.Phase, $rd.PendingCount, $rd.TaskCount, $agoText) -ForegroundColor Gray
                    }
                    Write-Host ""
                    $resumeInput = Read-Host ("  Select export to resume [1-{0}, N for new export]" -f $aeResumableDirs.Count)

                    if (-not [string]::IsNullOrEmpty($resumeInput)) {
                        $resumeIndex = ($resumeInput -as [int])
                        if ($resumeIndex -and $resumeIndex -ge 1 -and $resumeIndex -le $aeResumableDirs.Count) {
                            $result.Mode = "ActivityExplorerResume"
                            $result.AEResumePath = $aeResumableDirs[$resumeIndex - 1].ExportDir

                            # Ask for worker count
                            Write-Host ""
                            Write-Host "  Workers: [1] Single terminal (default)  [2-16] Multi-terminal with N workers" -ForegroundColor Gray
                            $wcInput = Read-Host "  Number of workers"
                            $wcVal = $wcInput -as [int]
                            if ($wcVal -and $wcVal -ge 2 -and $wcVal -le 16) {
                                $result.AEWorkerCount = $wcVal
                            }
                        }
                    }
                }
            }
            "14" { $result.Mode = "DLP" }
            "15" { $result.Mode = "Labels" }
            "16" { $result.Mode = "eDiscovery" }
            "17" { $result.Mode = "RBAC" }
            "Q" { $result.Quit = $true; return $result }
            "" { $result.Quit = $true; return $result }  # Default to quit on empty input
            default {
                Write-Host "`n  Invalid selection. Exiting." -ForegroundColor Yellow
                $result.Quit = $true
                return $result
            }
        }

    } while ($null -eq $result.Mode -and -not $result.Quit)

    # Show additional options for Activity Explorer
    if ($result.Mode -eq "ActivityExplorer" -or ($result.Mode -eq "Full" -and -not $result.NoActivity)) {
        Write-Host ""
        Write-Host "  Activity Explorer Options:" -ForegroundColor Yellow
        Write-Host ""

        $pastDaysInput = Read-Host "    Past days to export (1-30) [7]"
        if ($pastDaysInput -match '^\d+$') {
            $days = [int]$pastDaysInput
            if ($days -ge 1 -and $days -le 30) {
                $result.PastDays = $days
            }
            else {
                Write-Host "    Invalid range. Using default: 7 days" -ForegroundColor Yellow
            }
        }
    }

    # Output format selection (skip for resume mode since output already exists)
    if ($result.Mode -ne "ContentExplorerResume") {
        Write-Host ""
        Write-Host "  Output Options:" -ForegroundColor Yellow
        Write-Host ""
        $formatInput = Read-Host "    Output format (JSON/CSV) [JSON]"
        if (-not [string]::IsNullOrEmpty($formatInput) -and $formatInput.Trim().ToUpper() -eq "CSV") {
            $result.OutputFormat = "CSV"
        }
    }

    return $result
}

function Get-CEWorkloadSelection {
    <#
    .SYNOPSIS
        Prompts the user to select Content Explorer workloads.
    .OUTPUTS
        Array of workload names, or $null to use config file settings.
    #>
    [CmdletBinding()]
    param()

    Write-Host ""
    Write-SectionHeader -Text "Workload Selection" -Color Yellow
    Write-Host "    [1] As per Config File (default)" -ForegroundColor White
    Write-Host "    [2] SharePoint only" -ForegroundColor White
    Write-Host "    [3] SharePoint and OneDrive" -ForegroundColor White
    Write-Host "    [4] SharePoint, OneDrive and Exchange" -ForegroundColor White
    Write-Host "    [5] All Workloads (Exchange, SharePoint, OneDrive, Teams)" -ForegroundColor White
    $wlInput = Read-Host "  Select workloads [1]"

    switch ($wlInput) {
        "2" { return @("SharePoint") }
        "3" { return @("SharePoint", "OneDrive") }
        "4" { return @("SharePoint", "OneDrive", "Exchange") }
        "5" { return @("Exchange", "SharePoint", "OneDrive", "Teams") }
        default { return $null }
    }
}

function Test-NoParametersProvided {
    <#
    .SYNOPSIS
        Checks if the script was invoked with no explicit parameters.
    #>
    [CmdletBinding()]
    param()

    # Use script-scoped copy of bound parameters (set during initialization)
    if ($script:BoundParameterCount -eq 0) {
        return $true
    }

    # Check if any export mode switch was explicitly set
    $explicitParams = @(
        $script:FullExport.IsPresent,
        $script:DlpOnly.IsPresent,
        $script:LabelsOnly.IsPresent,
        $script:ContentExplorer.IsPresent,
        $script:ActivityExplorer.IsPresent,
        $script:eDiscoveryOnly.IsPresent,
        $script:RbacOnly.IsPresent,
        $script:Menu.IsPresent,
        $script:NoActivity.IsPresent,
        $script:NoContent.IsPresent,
        $script:AEResume.IsPresent
    )

    # Return true if no explicit switches
    $noExplicitParams = -not ($explicitParams -contains $true)

    Write-Verbose "Test-NoParametersProvided: BoundCount=$($script:BoundParameterCount), explicitParams=$($explicitParams -join ','), result=$noExplicitParams"

    return $noExplicitParams
}

function Build-AuthParameters {
    <#
    .SYNOPSIS
        Builds authentication parameters from AuthConfig.json or script parameters.
    .DESCRIPTION
        Reads ConfigFiles/AuthConfig.json. If certificate-based auth is configured and all
        required fields are present, returns certificate auth parameters. Otherwise falls
        back to UserPrincipalName if set, or returns an empty hashtable for interactive auth.
    .OUTPUTS
        Hashtable with authentication parameters for Connect-Compl8Compliance.
    #>
    [CmdletBinding()]
    param()

    $authConfigPath = Join-Path $PSScriptRoot "ConfigFiles\AuthConfig.json"

    if (Test-Path $authConfigPath) {
        $authConfig = Read-JsonConfig -Path $authConfigPath

        if ($authConfig.UseCertificateAuth -eq "True" -and
            -not [string]::IsNullOrEmpty($authConfig.AppId) -and
            -not [string]::IsNullOrEmpty($authConfig.CertificateThumbprint) -and
            -not [string]::IsNullOrEmpty($authConfig.Organization)) {

            Write-Verbose "Build-AuthParameters: Using certificate-based authentication"
            return @{
                AppId                 = $authConfig.AppId
                CertificateThumbprint = $authConfig.CertificateThumbprint
                Organization          = $authConfig.Organization
            }
        }
    }

    if ($UserPrincipalName) {
        Write-Verbose "Build-AuthParameters: Using UserPrincipalName=$UserPrincipalName"
        return @{
            UserPrincipalName = $UserPrincipalName
        }
    }

    Write-Verbose "Build-AuthParameters: Using interactive authentication (no config or UPN)"
    return @{}
}

function Start-WorkerTerminals {
    <#
    .SYNOPSIS
        Spawns worker terminal processes to join a Content Explorer export.
    .PARAMETER ExportRunDirectory
        Absolute path to the export run directory (e.g., Output/Export-20260131-...).
    .PARAMETER Count
        Number of worker terminals to spawn.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$ExportRunDirectory,

        [Parameter(Mandatory)]
        [ValidateRange(1, 16)]
        [int]$Count
    )

    $scriptPath = Join-Path $PSScriptRoot "Export-Compl8Configuration.ps1"

    # Check auth mode from config
    $authParams = Build-AuthParameters
    $isCertAuth = $authParams.ContainsKey('AppId')

    if (-not $isCertAuth) {
        Write-Host ""
        Write-Host "  ╔════════════════════════════════════════════════════════════════╗" -ForegroundColor Yellow
        Write-Host "  ║  INTERACTIVE AUTH: Each terminal needs manual browser login.  ║" -ForegroundColor Yellow
        Write-Host "  ║  You will be prompted before each terminal is spawned.       ║" -ForegroundColor Yellow
        Write-Host "  ║  Press Enter to spawn each terminal, or Q to stop spawning.  ║" -ForegroundColor Yellow
        Write-Host "  ╚════════════════════════════════════════════════════════════════╝" -ForegroundColor Yellow
        Write-Host ""
    } else {
        Write-Host ""
        Write-Host "  ╔════════════════════════════════════════════════════════════════╗" -ForegroundColor Green
        Write-Host "  ║  CERTIFICATE AUTH: Workers will authenticate automatically.   ║" -ForegroundColor Green
        Write-Host ("  ║  Spawning {0} worker terminal(s)...                          ║" -f $Count) -ForegroundColor Green
        Write-Host "  ╚════════════════════════════════════════════════════════════════╝" -ForegroundColor Green
        Write-Host ""
    }

    # Build base arguments for spawned processes
    $baseArgs = @(
        "-NoProfile",
        "-NoExit",
        "-File", ("`"{0}`"" -f $scriptPath),
        "-WorkerExportDir", ("`"{0}`"" -f $ExportRunDirectory),
        "-WorkerMode"
    )

    # Add PageSize if non-default
    if ($PageSize -and $PageSize -ne 5000) {
        $baseArgs += @("-PageSize", $PageSize.ToString())
    }

    $workerProcesses = [System.Collections.ArrayList]::new()

    for ($i = 1; $i -le $Count; $i++) {
        if (-not $isCertAuth) {
            $input = Read-Host "  Press Enter to spawn worker $i/$Count (Q to stop)"
            if ($input -and $input.Trim().ToUpper() -eq 'Q') {
                Write-ExportLog -Message "  Worker spawning stopped by user at $($i-1)/$Count" -Level Warning
                break
            }
        }

        try {
            $proc = Start-Process pwsh -ArgumentList $baseArgs -PassThru
            $workerDir = Get-WorkerCoordDir $ExportRunDirectory $proc.Id
            if (-not (Test-Path $workerDir)) {
                New-Item -ItemType Directory -Force -Path $workerDir | Out-Null
            }
            [void]$workerProcesses.Add(@{
                WorkerNumber = $i
                PID = $proc.Id
                Process = $proc
                WorkerDir = $workerDir
                SpawnedTime = (Get-Date).ToString("o")
            })
            Write-ExportLog -Message ("  Spawned worker {0}/{1} (PID: {2})" -f $i, $Count, $proc.Id) -Level Success

            # Brief delay between spawns to avoid auth collisions
            Start-Sleep -Seconds 2
        }
        catch {
            Write-ExportLog -Message "  Failed to spawn worker $i/$Count`: $($_.Exception.Message)" -Level Error
        }
    }

    if ($workerProcesses.Count -gt 0) {
        Write-ExportLog -Message "  $($workerProcesses.Count) worker(s) spawned successfully" -Level Success
    } else {
        Write-ExportLog -Message "  No workers were spawned" -Level Warning
    }

    return $workerProcesses
}

function Add-WorkerToExport {
    <#
    .SYNOPSIS
        Spawns a single additional worker terminal and returns a worker-process hashtable.
    .DESCRIPTION
        Used for dynamic worker spawning mid-export (e.g., via the W hotkey).
        Matches the hashtable format used in $workerProcesses.
    .PARAMETER ExportRunDirectory
        Absolute path to the export run directory.
    .PARAMETER NextWorkerNumber
        The worker number to assign (for display purposes).
    #>
    param(
        [Parameter(Mandatory)]
        [string]$ExportRunDirectory,

        [Parameter(Mandatory)]
        [int]$NextWorkerNumber
    )

    $scriptPath = Join-Path $PSScriptRoot "Export-Compl8Configuration.ps1"
    $authParams = Build-AuthParameters
    $isCertAuth = $authParams.ContainsKey('AppId')

    if (-not $isCertAuth) {
        Write-Host ""
        Write-Host "  Adding worker $NextWorkerNumber (interactive auth - browser window will open)..." -ForegroundColor Yellow
    }

    $baseArgs = @(
        "-NoProfile",
        "-NoExit",
        "-File", ("`"{0}`"" -f $scriptPath),
        "-WorkerExportDir", ("`"{0}`"" -f $ExportRunDirectory),
        "-WorkerMode"
    )

    try {
        $proc = Start-Process pwsh -ArgumentList $baseArgs -PassThru
        $workerDir = Get-WorkerCoordDir $ExportRunDirectory $proc.Id
        if (-not (Test-Path $workerDir)) {
            New-Item -ItemType Directory -Force -Path $workerDir | Out-Null
        }
        $worker = @{
            WorkerNumber = $NextWorkerNumber
            PID          = $proc.Id
            Process      = $proc
            WorkerDir    = $workerDir
            SpawnedTime  = (Get-Date).ToString("o")
        }
        Write-ExportLog -Message ("  Dynamic worker #{0} spawned (PID: {1})" -f $NextWorkerNumber, $proc.Id) -Level Success
        return $worker
    }
    catch {
        Write-ExportLog -Message ("  Failed to spawn dynamic worker #{0}: {1}" -f $NextWorkerNumber, $_.Exception.Message) -Level Error
        return $null
    }
}

function Test-AddWorkerKeypress {
    <#
    .SYNOPSIS
        Non-blocking check for the W key to dynamically add a worker mid-export.
    .DESCRIPTION
        If the user presses W, spawns a new worker via Add-WorkerToExport and appends
        it to the workerProcesses array. Non-W keys are consumed and discarded.
    #>
    param(
        [Parameter(Mandatory)][string]$ExportRunDirectory,
        [Parameter(Mandatory)][ref]$WorkerProcesses,
        [Parameter(Mandatory)][ref]$NextWorkerNumber
    )

    while ([Console]::KeyAvailable) {
        $key = [Console]::ReadKey($true)
        if ($key.Key -eq 'W') {
            $newWorker = Add-WorkerToExport -ExportRunDirectory $ExportRunDirectory -NextWorkerNumber $NextWorkerNumber.Value
            if ($newWorker) {
                [void]$WorkerProcesses.Value.Add($newWorker)
                $NextWorkerNumber.Value++
            }
        }
        # else: consume and discard non-W keys
    }
}

#endregion

#region Main Export Functions

function Save-ExportData {
    param(
        [Parameter(Mandatory)]
        $Data,

        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter(Mandatory)]
        [string]$Format,

        [Parameter(Mandatory)]
        [string]$Directory
    )

    # No timestamp in filename since directory already has timestamp
    if ($Format -eq "JSON") {
        $path = Join-Path $Directory "$Name.json"
        Export-ToJsonFile -Data $Data -Path $path
    }
    else {
        $path = Join-Path $Directory "$Name.csv"
        Export-ToCsvFile -Data $Data -Path $path
    }
}

function Invoke-FullExport {
    Write-ExportLog -Message "`n========== Full Configuration Export ==========" -Level Info
    Write-ExportLog -Message "Each module will be saved to a separate file" -Level Info

    $sectionsCompleted = 0
    $sectionsFailed = 0

    # DLP
    Write-ExportLog -Message "`n--- DLP Configuration ---" -Level Info
    try {
        $dlp = Export-DlpPolicies
        $sits = Export-SensitiveInfoTypes
        $dlpExport = @{
            ExportTimestamp = Get-Date -Format "yyyy-MM-ddTHH:mm:ssZ"
            Policies = $dlp.Policies
            Rules = $dlp.Rules
            SensitiveInfoTypes = $sits
        }
        Save-ExportData -Data $dlpExport -Name "DLP-Config" -Format $OutputFormat -Directory $script:ExportRunDirectory
        $sectionsCompleted++
    }
    catch {
        Write-ExportLog -Message "DLP export failed: $($_.Exception.Message)" -Level Error
        $sectionsFailed++
    }

    # Sensitivity Labels
    Write-ExportLog -Message "`n--- Sensitivity Labels ---" -Level Info
    try {
        $sensLabels = Export-SensitivityLabels
        $sensExport = @{
            ExportTimestamp = Get-Date -Format "yyyy-MM-ddTHH:mm:ssZ"
            Labels = $sensLabels.Labels
            Policies = $sensLabels.Policies
        }
        Save-ExportData -Data $sensExport -Name "SensitivityLabels-Config" -Format $OutputFormat -Directory $script:ExportRunDirectory
        $sectionsCompleted++
    }
    catch {
        Write-ExportLog -Message "Sensitivity Labels export failed: $($_.Exception.Message)" -Level Error
        $sectionsFailed++
    }

    # Retention Labels
    Write-ExportLog -Message "`n--- Retention Labels ---" -Level Info
    try {
        $retLabels = Export-RetentionLabels
        $retExport = @{
            ExportTimestamp = Get-Date -Format "yyyy-MM-ddTHH:mm:ssZ"
            Labels = $retLabels.Labels
            Policies = $retLabels.Policies
            Rules = $retLabels.Rules
        }
        Save-ExportData -Data $retExport -Name "RetentionLabels-Config" -Format $OutputFormat -Directory $script:ExportRunDirectory
        $sectionsCompleted++
    }
    catch {
        Write-ExportLog -Message "Retention Labels export failed: $($_.Exception.Message)" -Level Error
        $sectionsFailed++
    }

    # eDiscovery
    Write-ExportLog -Message "`n--- eDiscovery ---" -Level Info
    try {
        $ediscovery = Export-eDiscoveryCases
        $ediscExport = @{
            ExportTimestamp = Get-Date -Format "yyyy-MM-ddTHH:mm:ssZ"
            Cases = $ediscovery.Cases
            Searches = $ediscovery.Searches
        }
        Save-ExportData -Data $ediscExport -Name "eDiscovery-Config" -Format $OutputFormat -Directory $script:ExportRunDirectory
        $sectionsCompleted++
    }
    catch {
        Write-ExportLog -Message "eDiscovery export failed: $($_.Exception.Message)" -Level Error
        $sectionsFailed++
    }

    # RBAC
    Write-ExportLog -Message "`n--- RBAC ---" -Level Info
    try {
        $rbac = Export-RbacConfiguration
        $rbacExport = @{
            ExportTimestamp = Get-Date -Format "yyyy-MM-ddTHH:mm:ssZ"
            RoleGroups = $rbac.RoleGroups
            Members = $rbac.Members
        }
        Save-ExportData -Data $rbacExport -Name "RBAC-Config" -Format $OutputFormat -Directory $script:ExportRunDirectory
        $sectionsCompleted++
    }
    catch {
        Write-ExportLog -Message "RBAC export failed: $($_.Exception.Message)" -Level Error
        $sectionsFailed++
    }

    # Content Explorer
    if ($NoContent) {
        Write-ExportLog -Message "`n--- Content Explorer ---" -Level Info
        Write-ExportLog -Message "  Skipped (-NoContent specified)" -Level Warning
    }
    else {
        Write-ExportLog -Message "`n--- Content Explorer ---" -Level Info
        try {
            Invoke-ContentExplorerExport
            $sectionsCompleted++
        }
        catch {
            Write-ExportLog -Message "Content Explorer export failed: $($_.Exception.Message)" -Level Error
            $sectionsFailed++
        }
    }

    # Activity Explorer
    if ($NoActivity) {
        Write-ExportLog -Message "`n--- Activity Explorer ---" -Level Info
        Write-ExportLog -Message "  Skipped (-NoActivity specified)" -Level Warning
    }
    else {
        Write-ExportLog -Message "`n--- Activity Explorer ---" -Level Info
        try {
            Invoke-ActivityExplorerExport
            $sectionsCompleted++
        }
        catch {
            Write-ExportLog -Message "Activity Explorer export failed: $($_.Exception.Message)" -Level Error
            $sectionsFailed++
        }
    }

    Write-ExportLog -Message "`nSections completed: $sectionsCompleted, Failed: $sectionsFailed" -Level Info
}

function Invoke-DlpExport {
    Write-ExportLog -Message "`n========== DLP Export ==========" -Level Info

    $exportResult = @{
        ExportTimestamp = Get-Date -Format "yyyy-MM-ddTHH:mm:ssZ"
        Policies = @()
        Rules = @()
        SensitiveInfoTypes = @()
    }

    # DLP Policies
    try {
        $dlp = Export-DlpPolicies
        $exportResult.Policies = $dlp.Policies
        $exportResult.Rules = $dlp.Rules
    }
    catch {
        Write-ExportLog -Message "DLP Policies export failed: $($_.Exception.Message)" -Level Error
    }

    # Sensitive Info Types
    try {
        $sits = Export-SensitiveInfoTypes
        $exportResult.SensitiveInfoTypes = $sits
    }
    catch {
        Write-ExportLog -Message "SITs export failed: $($_.Exception.Message)" -Level Error
    }

    if ($OutputFormat -eq "JSON") {
        Save-ExportData -Data $exportResult -Name "DLP-Config" -Format $OutputFormat -Directory $script:ExportRunDirectory
    }
    else {
        # CSV: separate files for each type
        if (@($exportResult.Policies).Count -gt 0) {
            Save-ExportData -Data $exportResult.Policies -Name "DLP-Policies" -Format $OutputFormat -Directory $script:ExportRunDirectory
        }
        if (@($exportResult.Rules).Count -gt 0) {
            Save-ExportData -Data $exportResult.Rules -Name "DLP-Rules" -Format $OutputFormat -Directory $script:ExportRunDirectory
        }
        if (@($exportResult.SensitiveInfoTypes).Count -gt 0) {
            Save-ExportData -Data $exportResult.SensitiveInfoTypes -Name "SensitiveInfoTypes" -Format $OutputFormat -Directory $script:ExportRunDirectory
        }
    }
}

function Invoke-LabelsExport {
    Write-ExportLog -Message "`n========== Labels Export ==========" -Level Info

    $exportResult = @{
        ExportTimestamp = Get-Date -Format "yyyy-MM-ddTHH:mm:ssZ"
        SensitivityLabels = @{ Labels = @(); Policies = @() }
        RetentionLabels = @{ Labels = @(); Policies = @() }
    }

    # Sensitivity Labels
    try {
        $sensLabels = Export-SensitivityLabels
        $exportResult.SensitivityLabels = $sensLabels
    }
    catch {
        Write-ExportLog -Message "Sensitivity Labels export failed: $($_.Exception.Message)" -Level Error
    }

    # Retention Labels
    try {
        $retLabels = Export-RetentionLabels
        $exportResult.RetentionLabels = $retLabels
    }
    catch {
        Write-ExportLog -Message "Retention Labels export failed: $($_.Exception.Message)" -Level Error
    }

    if ($OutputFormat -eq "JSON") {
        Save-ExportData -Data $exportResult -Name "Labels-Config" -Format $OutputFormat -Directory $script:ExportRunDirectory
    }
    else {
        # CSV: separate files
        if (@($exportResult.SensitivityLabels.Labels).Count -gt 0) {
            Save-ExportData -Data $exportResult.SensitivityLabels.Labels -Name "SensitivityLabels" -Format $OutputFormat -Directory $script:ExportRunDirectory
        }
        if (@($exportResult.SensitivityLabels.Policies).Count -gt 0) {
            Save-ExportData -Data $exportResult.SensitivityLabels.Policies -Name "LabelPolicies" -Format $OutputFormat -Directory $script:ExportRunDirectory
        }
        if (@($exportResult.RetentionLabels.Labels).Count -gt 0) {
            Save-ExportData -Data $exportResult.RetentionLabels.Labels -Name "RetentionLabels" -Format $OutputFormat -Directory $script:ExportRunDirectory
        }
        if (@($exportResult.RetentionLabels.Policies).Count -gt 0) {
            Save-ExportData -Data $exportResult.RetentionLabels.Policies -Name "RetentionPolicies" -Format $OutputFormat -Directory $script:ExportRunDirectory
        }
    }
}

function Invoke-ContentExplorerWorker {
    <#
    .SYNOPSIS
        Content Explorer worker using file-drop coordination.
    .DESCRIPTION
        Runs in a spawned worker terminal. Receives tasks from the orchestrator via
        file-drop (nexttask/currenttask files) and writes output to its Worker-PID/
        subfolder. Handles both Aggregate and Detail phases.

        Coordination protocol:
        - Receive-WorkerTask reads nexttask, renames to currenttask, returns hashtable
        - Complete-WorkerTask deletes currenttask file
        - Read-ExportPhase reads ExportPhase.txt from export directory
        - Orchestrator assigns tasks and monitors worker liveness via Get-Process
    .PARAMETER WorkerExportDir
        The export run directory (e.g., Output/Export-20260131-...).
        Worker creates its own Worker-PID/ subfolder within this directory.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$WorkerExportDir
    )

    $exportDir = $WorkerExportDir

    Write-Host "`n  Worker starting in file-drop mode..." -ForegroundColor Yellow
    Write-Host ("  Export directory: {0}" -f $exportDir) -ForegroundColor Gray

    # Create Worker coordination subfolder
    $workerDir = Get-WorkerCoordDir $exportDir $PID
    if (-not (Test-Path $workerDir)) {
        New-Item -ItemType Directory -Force -Path $workerDir | Out-Null
    }

    # Initialize logging
    Initialize-ExportLog -LogDirectory (Get-LogsDir $exportDir)

    # Set up error log paths (shared + per-worker)
    $script:ErrorLogPath = Join-Path (Get-LogsDir $exportDir) "ExportProject-Errors.log"
    $script:WorkerErrorLogPath = Join-Path (Get-LogsDir $exportDir) "ExportErrors-$PID.log"

    # Worker-specific progress log
    $progressLogPath = Join-Path $workerDir "Progress.log"
    $signalSigningKey = Get-ExportRunSigningKey -ExportDir $exportDir -CreateIfMissing

    Write-ExportLog -Message "Worker PID $PID started (file-drop), output folder: $workerDir" -Level Info
    Write-ProgressEntry -Path $progressLogPath -Message "Worker PID $PID started"

    # Load Content Explorer configuration (prefer saved manifest for consistency)
    $savedSettings = Get-ExportSettings -ExportRunDirectory $exportDir
    $configPath = Join-Path $PSScriptRoot "ConfigFiles" "ContentExplorerClassifiers.json"
    $ceConfig = Read-JsonConfig -Path $configPath
    if (-not $ceConfig -and -not $savedSettings) {
        Write-ExportLog -Message "ERROR: Cannot read Content Explorer config" -Level Error
        return
    }

    # Extract config settings (manifest overrides config file)
    if ($savedSettings) {
        Write-ExportLog -Message "CE Worker using saved settings from ExportSettings.json" -Level Info
        $batchSize = if ($savedSettings.BatchSize) { $savedSettings.BatchSize -as [int] } else { $script:CEDefaultBatchSize }
        $workloads = if ($savedSettings.Workloads) { @($savedSettings.Workloads) } else { @($script:CEDefaultWorkloads) }
        $cePageSize = if ($savedSettings.PageSize) { $savedSettings.PageSize -as [int] } else { $PageSize }
    }
    else {
        Write-ExportLog -Message "CE Worker using current config file (no manifest found)" -Level Warning
        $batchSize = if ($ceConfig._BatchSize) { ($ceConfig._BatchSize -as [int]) } else { $null }
        $workloads = if ($ceConfig._Workloads) { @($ceConfig._Workloads) } else { @($script:CEDefaultWorkloads) }
        $cePageSize = if ($ceConfig._PageSize) { ($ceConfig._PageSize -as [int]) } else { $null }
    }
    if (-not $batchSize -or $batchSize -lt 1) { $batchSize = $script:CEDefaultBatchSize }
    if (-not $cePageSize -or $cePageSize -lt 1) { $cePageSize = $PageSize }

    # Per-worker run tracker (in worker folder)
    $trackerPath = Join-Path $workerDir "RunTracker.json"
    $tracker = Get-ContentExplorerRunTracker -TrackerPath $trackerPath

    # Telemetry setup
    $telemetryDbPath = Join-Path $PSScriptRoot "TelemetryDB" "content-explorer-telemetry.jsonl"

    # Aggregate data (loaded once when entering Detail phase)
    $aggregateDataLoaded = $false
    $aggregateData = $null

    $tasksExported = 0
    $lastWorkerActivity = Get-Date
    $workerInactivityLimit = New-TimeSpan -Minutes $script:CEWorkerInactivityMinutes

    Write-ExportLog -Message "Worker entering file-drop task loop" -Level Info

    #region Worker Task Loop
    while ($true) {
        # Try to receive a task from orchestrator
        $task = Receive-WorkerTask -WorkerDir $workerDir -ExportDir $exportDir

        if (-not $task) {
            # No task available -- check if we should exit or wait
            $phase = Read-ExportPhase -ExportDir $exportDir

            if ($phase -eq "Completed") {
                Write-ExportLog -Message "Project phase is $phase - worker exiting cleanly" -Level Info
                Write-ProgressEntry -Path $progressLogPath -Message "Phase is $phase - exiting"
                break
            }

            # Inactivity timeout — but check one last time for a task before exiting
            if (((Get-Date) - $lastWorkerActivity) -gt $workerInactivityLimit) {
                $finalCheck = Receive-WorkerTask -WorkerDir $workerDir -ExportDir $exportDir
                if ($finalCheck) {
                    $lastWorkerActivity = Get-Date
                    $task = $finalCheck
                    # Fall through to task processing below
                } else {
                    Write-ExportLog -Message "No activity for 35 minutes - worker exiting" -Level Warning
                    Write-ProgressEntry -Path $progressLogPath -Message "Inactivity timeout - exiting"
                    break
                }
            }

            if (-not $task) {
                Start-Sleep -Seconds 2
                continue
            }
        }

        # We have a task - reset activity timer
        $lastWorkerActivity = Get-Date
        $taskPhase = $task.Phase

        # Build sanitized key for filenames: replace illegal chars with underscore
        $taskTagType = if ($task.TagType) { $task.TagType } else { "Unknown" }
        $taskTagName = if ($task.TagName) { $task.TagName } else { "Unknown" }
        $taskWorkload = if ($task.Workload) { $task.Workload } else { "Unknown" }
        $taskKey = "{0}|{1}|{2}" -f $taskTagType, $taskTagName, $taskWorkload
        $sanitizedKey = ("{0}-{1}-{2}" -f $taskTagType, $taskTagName, $taskWorkload) -replace '[\\/:\*\?"<>\|]', '_'

        Write-ExportLog -Message ("  Received task: {0} (Phase: {1})" -f $taskKey, $taskPhase) -Level Info
        Write-ProgressEntry -Path $progressLogPath -Message ("Received task: {0} (Phase: {1})" -f $taskKey, $taskPhase)

        # ── Aggregate phase ──
        if ($taskPhase -eq "Aggregate") {
            $aggDir = Join-Path (Get-CEDataDir $exportDir) "Aggregates"
            if (-not (Test-Path $aggDir)) { New-Item -ItemType Directory -Force -Path $aggDir | Out-Null }
            $aggCsvPath = Join-Path $aggDir ("agg-{0}.csv" -f $sanitizedKey)
            $aggErrorPath = Join-Path $workerDir ("error-agg-{0}.txt" -f $sanitizedKey)

            $aggSuccess = $false
            $aggLocations = @()
            $totalCount = 0
            $aggError = $null
            $maxAggRetries = 3
            $aggFinalAttemptDelay = 120

            try {
                $allAggregates = @()
                $pageCookie = $null
                $aggPageNum = 0

                do {
                    $aggParams = @{
                        TagType     = $taskTagType
                        TagName     = $taskTagName
                        Workload    = $taskWorkload
                        PageSize    = 5000
                        Aggregate   = $true
                        ErrorAction = 'Stop'
                    }
                    if ($pageCookie) { $aggParams['PageCookie'] = $pageCookie }

                    $pageSuccess = $false
                    $pageRetry = 0

                    while (-not $pageSuccess -and $pageRetry -le $maxAggRetries) {
                        try {
                            $aggResult = Export-ContentExplorerData @aggParams
                            $pageSuccess = $true
                        }
                        catch {
                            $lastPageError = $_
                            $pageRetry++
                            $errorInfo = Get-HttpErrorExplanation -ErrorMessage $_.Exception.Message -ErrorRecord $_
                            $statusStr = if ($errorInfo.StatusCode) { "HTTP $($errorInfo.StatusCode)" } else { $errorInfo.Category }

                            # Connection lost - cmdlet not available (session dropped or never established)
                            if ($_.Exception -is [System.Management.Automation.CommandNotFoundException]) {
                                Write-ExportLog -Message "    CONNECTION LOST: S&C cmdlet not available - attempting reconnection..." -Level Warning
                                try {
                                    Disconnect-Compl8Compliance
                                    if ($script:AuthParams -and $script:AuthParams.Count -gt 0) {
                                        $reAuthResult = Connect-Compl8Compliance @script:AuthParams
                                        if ($reAuthResult) {
                                            Write-ExportLog -Message "    Reconnection successful - retrying" -Level Success
                                            $pageRetry--
                                            continue
                                        }
                                    }
                                }
                                catch {
                                    Write-ExportLog -Message ("    Reconnection failed: {0}" -f $_.Exception.Message) -Level Error
                                }
                                # Cannot recover - throw to outer catch (bad session tracking will exit the worker)
                                throw $lastPageError
                            }

                            # Auth recovery
                            if ($errorInfo.Category -eq "AuthError") {
                                Write-ExportLog -Message "    AUTH EXPIRED during aggregate - attempting recovery..." -Level Warning
                                try {
                                    Disconnect-Compl8Compliance
                                    if ($script:AuthParams -and $script:AuthParams.Count -gt 0) {
                                        $reAuthResult = Connect-Compl8Compliance @script:AuthParams
                                        if ($reAuthResult) {
                                            Write-ExportLog -Message "    Re-authentication successful - retrying" -Level Success
                                            $pageRetry--
                                            continue
                                        }
                                    }
                                }
                                catch {
                                    Write-ExportLog -Message ("    Re-authentication failed: {0}" -f $_.Exception.Message) -Level Error
                                }
                                throw $lastPageError
                            }

                            # Log the error
                            if ($script:ErrorLogPath) {
                                Write-ExportErrorLog -ErrorLogPath $script:ErrorLogPath -Context "Worker Aggregate (Page Retry)" -TaskKey $taskKey -ErrorRecord $_ -AdditionalData @{ RetryCount = $pageRetry; MaxRetries = $maxAggRetries; Page = $aggPageNum }
                            }
                            if ($script:WorkerErrorLogPath) {
                                Write-ExportErrorLog -ErrorLogPath $script:WorkerErrorLogPath -Context "Worker Aggregate (Page Retry)" -TaskKey $taskKey -ErrorRecord $_ -AdditionalData @{ RetryCount = $pageRetry; MaxRetries = $maxAggRetries; Page = $aggPageNum }
                            }

                            if ($errorInfo.IsTransient -and $pageRetry -le $maxAggRetries) {
                                if ($pageRetry -eq $maxAggRetries) {
                                    $msg = "    Aggregate TRANSIENT ERROR [{0}] (attempt {1}/{2}) - final attempt in {3}s" -f $statusStr, $pageRetry, $maxAggRetries, $aggFinalAttemptDelay
                                    Write-ExportLog -Message $msg -Level Warning
                                    Start-Sleep -Seconds $aggFinalAttemptDelay
                                }
                                else {
                                    $retryDelay = 60 * $pageRetry
                                    $msg = "    Aggregate TRANSIENT ERROR [{0}] (attempt {1}/{2}) - waiting {3}s" -f $statusStr, $pageRetry, $maxAggRetries, $retryDelay
                                    Write-ExportLog -Message $msg -Level Warning
                                    Start-Sleep -Seconds $retryDelay
                                }
                            }
                            else {
                                $msg = "    Aggregate FAILED [{0}] after {1} attempts" -f $statusStr, $pageRetry
                                Write-ExportLog -Message $msg -Level Error
                                throw
                            }
                        }
                    }

                    $aggPageNum++

                    if ($null -eq $aggResult -or $aggResult.Count -eq 0) { break }

                    $metadata = $aggResult[0]
                    if ($metadata.RecordsReturned -gt 0) {
                        $allAggregates += $aggResult[1..$metadata.RecordsReturned]
                    }

                    if ($metadata.MorePagesAvailable -eq $true -or $metadata.MorePagesAvailable -eq "True") {
                        $pageCookie = $metadata.PageCookie
                    }
                    else { break }
                } while ($true)

                $aggSuccess = $true
                $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                $matchCount = 0

                # Build CSV content for this aggregate task
                $csvSb = [System.Text.StringBuilder]::new()
                [void]$csvSb.AppendLine("Timestamp,TagType,TagName,Workload,Location,Count,Error")
                foreach ($agg in $allAggregates) {
                    $aggLocations += @{ Name = $agg.Name; ExpectedCount = $agg.Count; ExportedCount = 0 }
                    $matchCount += $agg.Count
                    $locationName = $agg.Name -replace '"', '""'
                    if ($locationName -match '[,"]') { $locationName = ('"' + $locationName + '"') }
                    $csvLine = "{0},{1},{2},{3},{4},{5}," -f $timestamp, $taskTagType, $taskTagName, $taskWorkload, $locationName, $agg.Count
                    [void]$csvSb.AppendLine($csvLine)
                }

                # Probe detail API for actual file count (aggregate returns match count, not file count)
                $fileCount = $matchCount
                if ($matchCount -gt 0) {
                    try {
                        $probeResult = Export-ContentExplorerData -TagType $taskTagType -TagName $taskTagName -Workload $taskWorkload -PageSize 1 -ErrorAction Stop
                        if ($probeResult -and $probeResult[0] -and $null -ne $probeResult[0].TotalCount) {
                            $probedCount = $probeResult[0].TotalCount -as [int]
                            if ($probedCount -and $probedCount -gt 0) {
                                $fileCount = $probedCount
                                if ($fileCount -ne $matchCount) {
                                    Write-ExportLog -Message ("    File count probe: {0} files vs {1} matches" -f $fileCount, $matchCount) -Level Info
                                }
                            }
                        }
                    }
                    catch {
                        Write-ExportLog -Message ("    File count probe failed (using match count): {0}" -f $_.Exception.Message) -Level Warning
                    }
                }
                $totalCount = $fileCount

                # Write _FILECOUNT row so planning phase can use file-level counts
                $csvLine = "{0},{1},{2},{3},_FILECOUNT,{4}," -f $timestamp, $taskTagType, $taskTagName, $taskWorkload, $fileCount
                [void]$csvSb.AppendLine($csvLine)

                # Write marker row for zero-result tasks so orchestrator can identify and complete them
                if ($allAggregates.Count -eq 0) {
                    $csvLine = "{0},{1},{2},{3},NONE,0," -f $timestamp, $taskTagType, $taskTagName, $taskWorkload
                    [void]$csvSb.AppendLine($csvLine)
                }

                # Write per-worker aggregate CSV
                [System.IO.File]::WriteAllText($aggCsvPath, $csvSb.ToString(), [System.Text.Encoding]::UTF8)

                $msg = "    -> {0} files ({1} matches) in {2} locations" -f $fileCount, $matchCount, $allAggregates.Count
                Write-ExportLog -Message $msg -Level Success
                Write-ProgressEntry -Path $progressLogPath -Message ("Aggregate complete: {0} -> {1} files ({2} matches), {3} locations" -f $taskKey, $fileCount, $matchCount, $allAggregates.Count)
            }
            catch {
                $aggError = $_.Exception.Message
                $isCmdletNotFound = $_.Exception -is [System.Management.Automation.CommandNotFoundException]

                if ($isCmdletNotFound) {
                    # Bad session: don't write error output so orchestrator reclaims this task as Pending
                    $cmdletNotFoundCount++
                    Write-ExportLog -Message ("    AGGREGATE SKIPPED (bad session #{0}): {1} - task will be reassigned" -f $cmdletNotFoundCount, $taskKey) -Level Warning
                    Write-ProgressEntry -Path $progressLogPath -Message ("Bad session #{0}: {1} -> task returned to queue" -f $cmdletNotFoundCount, $taskKey)
                }
                else {
                    Write-ExportLog -Message ("    AGGREGATE FAILED: {0}" -f $aggError) -Level Error
                    Write-ProgressEntry -Path $progressLogPath -Message ("Aggregate FAILED: {0} -> {1}" -f $taskKey, $aggError)

                    # Write per-worker aggregate CSV with error row
                    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                    $escapedError = $aggError -replace '"', '""'
                    $csvSb = [System.Text.StringBuilder]::new()
                    [void]$csvSb.AppendLine("Timestamp,TagType,TagName,Workload,Location,Count,Error")
                    $errorCsvLine = '{0},{1},{2},{3},ERROR,0,"{4}"' -f $timestamp, $taskTagType, $taskTagName, $taskWorkload, $escapedError
                    [void]$csvSb.AppendLine($errorCsvLine)
                    [System.IO.File]::WriteAllText($aggCsvPath, $csvSb.ToString(), [System.Text.Encoding]::UTF8)

                    # Write error file as JSON (orchestrator parses with ConvertFrom-Json)
                    $errorPayload = @{
                        Timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff")
                        TaskKey   = $taskKey
                        TagType   = $taskTagType
                        TagName   = $taskTagName
                        Workload  = $taskWorkload
                        Error     = $aggError
                    }
                    $errorJson = ConvertTo-SignedEnvelopeJson -Payload $errorPayload -SigningKey $signalSigningKey
                    [System.IO.File]::WriteAllText($aggErrorPath, $errorJson, [System.Text.Encoding]::UTF8)

                    # Log to shared and per-worker error logs
                    if ($script:ErrorLogPath) {
                        Write-ExportErrorLog -ErrorLogPath $script:ErrorLogPath -Context "Worker Aggregate" -TaskKey $taskKey -ErrorRecord $_
                    }
                    if ($script:WorkerErrorLogPath) {
                        Write-ExportErrorLog -ErrorLogPath $script:WorkerErrorLogPath -Context "Worker Aggregate" -TaskKey $taskKey -ErrorRecord $_
                    }
                }
            }

            Complete-WorkerTask -WorkerDir $workerDir
            $lastWorkerActivity = Get-Date

            # Bad session: exit worker after more than 3 tasks fail with cmdlet not found
            # Orchestrator will detect dead worker and reclaim all InProgress tasks as Pending
            if ($cmdletNotFoundCount -gt 3) {
                Write-ExportLog -Message ("Worker exiting: {0} tasks failed with cmdlet not recognized - bad session (tasks returned to queue)" -f $cmdletNotFoundCount) -Level Error
                Write-ProgressEntry -Path $progressLogPath -Message ("Exiting: bad session ({0} cmdlet-not-found failures)" -f $cmdletNotFoundCount)
                break
            }
            continue
        }

        # ── Detail phase ──
        if ($taskPhase -eq "Detail") {
            $locationSuffix = if ($task.Location) { "-" + ([Math]::Abs($task.Location.GetHashCode())).ToString("X8") } else { "" }
            $classifierDir = Get-CEClassifierDir $exportDir $taskTagType $taskTagName
            $completionsDir = Get-CompletionsDir $exportDir
            if (-not (Test-Path $completionsDir)) { New-Item -Path $completionsDir -ItemType Directory -Force | Out-Null }
            $detailErrorPath = Join-Path $completionsDir ("error-detail-{0}{1}-{2}.txt" -f $sanitizedKey, $locationSuffix, $PID)
            $detailDonePath = Join-Path $completionsDir ("detail-done-{0}{1}-{2}.txt" -f $sanitizedKey, $locationSuffix, $PID)

            # Load aggregate data once for building detail task context
            if (-not $aggregateDataLoaded) {
                # Check for central aggregate CSV in coordination dir
                $coordDir = Get-CoordinationDir $exportDir
                $rootAggCsv = Join-Path $coordDir "ContentExplorer-Aggregates.csv"
                $aggregateCsvPath = $rootAggCsv

                if (-not (Test-Path $aggregateCsvPath)) {
                    # Fall back to scanning Data/ContentExplorer/Aggregates/ for per-worker agg files
                    $aggDataDir = Join-Path (Get-CEDataDir $exportDir) "Aggregates"
                    if (Test-Path $aggDataDir) {
                        $aggCsvFiles = Get-ChildItem -Path $aggDataDir -Filter "agg-*.csv" -ErrorAction SilentlyContinue
                        if ($aggCsvFiles -and $aggCsvFiles.Count -gt 0) {
                            $aggregateCsvPath = $aggCsvFiles | Sort-Object LastWriteTime -Descending | Select-Object -First 1 -ExpandProperty FullName
                        }
                    }
                }

                if ($aggregateCsvPath -and (Test-Path $aggregateCsvPath)) {
                    $aggregateData = Import-AggregateDataFromCsv -CsvPath $aggregateCsvPath
                    $aggregateDataLoaded = $true
                    Write-ExportLog -Message "Worker loaded aggregate data for detail export" -Level Info
                }
            }

            # Build task context for Export-ContentExplorerWithProgress
            $expectedCount = if ($task.ExpectedCount) { ($task.ExpectedCount -as [int]) } else { 0 }
            $taskLocations = @()

            # Try to get location data from aggregate data
            if ($aggregateData -and $aggregateData.TaskData -and $aggregateData.TaskData.ContainsKey($taskKey)) {
                $taskAggEntry = $aggregateData.TaskData[$taskKey]
                if (-not $expectedCount -or $expectedCount -eq 0) {
                    $expectedCount = ($taskAggEntry.TotalCount -as [int])
                }
                $taskLocations = @($taskAggEntry.Locations)
            }

            # Also check if task itself carries location data
            if ($taskLocations.Count -eq 0 -and $task.Locations) {
                $taskLocations = @($task.Locations)
            }

            $detailTask = @{
                TaskId        = $taskKey
                TagType       = $taskTagType
                TagName       = $taskTagName
                Workload      = $taskWorkload
                Location      = if ($task.Location) { $task.Location } else { "" }
                LocationType  = if ($task.LocationType) { $task.LocationType } else { "" }
                Status        = "Pending"
                ExpectedCount = $expectedCount
                Locations     = $taskLocations
            }

            # Use orchestrator-computed page size from file-drop if available, otherwise fall back to default
            $taskPageSize = if ($task.PageSize -and ($task.PageSize -as [int]) -gt 0) { ($task.PageSize -as [int]) } else { $cePageSize }

            Write-ExportLog -Message ("  Detail export: {0} (Expected: {1}, PageSize: {2})" -f $taskKey, $expectedCount, $taskPageSize) -Level Info
            Write-ProgressEntry -Path $progressLogPath -Message ("Detail export starting: {0} (Expected: {1}, PageSize: {2})" -f $taskKey, $expectedCount, $taskPageSize)

            # Build location filter params for location-based tasks
            $locationParams = @{}
            if ($task.LocationType -eq "SiteUrl" -and $task.Location) {
                $locationParams["SiteUrl"] = $task.Location
            } elseif ($task.LocationType -eq "UPN" -and $task.Location) {
                $locationParams["UserPrincipalName"] = $task.Location
            }
            # WorkloadFallback: no location filter (existing behavior)

            $detailTaskStartTime = Get-Date
            try {
                $telemetry = New-ContentExplorerTelemetry -TagType $taskTagType -TagName $taskTagName -Workload $taskWorkload
                Export-ContentExplorerWithProgress -Task $detailTask -PageSize $taskPageSize -ProgressLogPath $progressLogPath -Telemetry $telemetry -TelemetryDatabasePath $telemetryDbPath -OutputDirectory $classifierDir @locationParams | Out-Null

                $exportedCount = if ($detailTask.ExportedCount) { $detailTask.ExportedCount } else { 0 }

                if ($exportedCount -gt 0) {
                    Write-ExportLog -Message ("    -> Exported {0} records to {1}" -f $exportedCount, $classifierDir) -Level Success
                }
                else {
                    Write-ExportLog -Message ("    -> No records exported for {0}" -f $taskKey) -Level Info
                }

                Write-ProgressEntry -Path $progressLogPath -Message ("Detail complete: {0} -> {1} records" -f $taskKey, $exportedCount)

                # Update run tracker
                if (-not $tracker.CompletedTasks) { $tracker.CompletedTasks = @() }
                $tracker.CompletedTasks += $taskKey
                $tracker.TotalExported = ($tracker.TotalExported -as [int]) + $exportedCount

                if (-not $tracker.OutputFiles) { $tracker.OutputFiles = @() }
                if ($exportedCount -gt 0) {
                    $tracker.OutputFiles += @{
                        TaskKey         = $taskKey
                        OutputDirectory = $classifierDir
                        RecordCount     = $exportedCount
                        Pages           = $detailTask.TotalPages
                        CompletedTime   = (Get-Date).ToString("o")
                    }
                }

                Save-ContentExplorerRunTracker -Tracker $tracker -TrackerPath $trackerPath

                # Write detail-done signal file (orchestrator watches for these)
                $detailTaskElapsed = ((Get-Date) - $detailTaskStartTime).TotalSeconds
                $donePayload = @{
                    TagType        = $taskTagType
                    TagName        = $taskTagName
                    Workload       = $taskWorkload
                    Location       = if ($task.Location) { $task.Location } else { "" }
                    LocationType   = if ($task.LocationType) { $task.LocationType } else { "" }
                    RecordCount    = $exportedCount
                    ElapsedSeconds = [Math]::Round($detailTaskElapsed, 1)
                    Timestamp      = (Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff")
                }
                $doneJson = ConvertTo-SignedEnvelopeJson -Payload $donePayload -SigningKey $signalSigningKey
                [System.IO.File]::WriteAllText($detailDonePath, $doneJson, [System.Text.Encoding]::UTF8)

                $tasksExported++
            }
            catch {
                $detailError = $_.Exception.Message
                $isCmdletNotFound = $_.Exception -is [System.Management.Automation.CommandNotFoundException]

                if ($isCmdletNotFound) {
                    # Bad session: don't write error output so orchestrator reclaims this task as Pending
                    $cmdletNotFoundCount++
                    Write-ExportLog -Message ("    DETAIL SKIPPED (bad session #{0}): {1} - task will be reassigned" -f $cmdletNotFoundCount, $taskKey) -Level Warning
                    Write-ProgressEntry -Path $progressLogPath -Message ("Bad session #{0}: {1} -> task returned to queue" -f $cmdletNotFoundCount, $taskKey)
                }
                else {
                    Write-ExportLog -Message ("    DETAIL EXPORT FAILED: {0}" -f $detailError) -Level Error
                    Write-ProgressEntry -Path $progressLogPath -Message ("Detail FAILED: {0} -> {1}" -f $taskKey, $detailError)

                    # Write error file as JSON (orchestrator parses with ConvertFrom-Json)
                    $errorPayload = @{
                        Timestamp    = (Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff")
                        TaskKey      = $taskKey
                        TagType      = $taskTagType
                        TagName      = $taskTagName
                        Workload     = $taskWorkload
                        Location     = if ($task.Location) { $task.Location } else { "" }
                        LocationType = if ($task.LocationType) { $task.LocationType } else { "" }
                        Error        = $detailError
                    }
                    $errorJson = ConvertTo-SignedEnvelopeJson -Payload $errorPayload -SigningKey $signalSigningKey
                    [System.IO.File]::WriteAllText($detailErrorPath, $errorJson, [System.Text.Encoding]::UTF8)

                    # Log to shared and per-worker error logs
                    if ($script:ErrorLogPath) {
                        Write-ExportErrorLog -ErrorLogPath $script:ErrorLogPath -Context "Worker Detail Export" -TaskKey $taskKey -ErrorRecord $_
                    }
                    if ($script:WorkerErrorLogPath) {
                        Write-ExportErrorLog -ErrorLogPath $script:WorkerErrorLogPath -Context "Worker Detail Export" -TaskKey $taskKey -ErrorRecord $_
                    }
                }
            }

            Complete-WorkerTask -WorkerDir $workerDir
            $lastWorkerActivity = Get-Date

            # Bad session: exit worker after more than 3 tasks fail with cmdlet not found
            if ($cmdletNotFoundCount -gt 3) {
                Write-ExportLog -Message ("Worker exiting: {0} tasks failed with cmdlet not recognized - bad session (tasks returned to queue)" -f $cmdletNotFoundCount) -Level Error
                Write-ProgressEntry -Path $progressLogPath -Message ("Exiting: bad session ({0} cmdlet-not-found failures)" -f $cmdletNotFoundCount)
                break
            }
            continue
        }

        # ── Unknown phase: acknowledge and move on ──
        Write-ExportLog -Message ("  Unknown task phase: {0} - skipping" -f $taskPhase) -Level Warning
        Complete-WorkerTask -WorkerDir $workerDir
    }
    #endregion

    # Worker summary
    $summaryMsg = "Worker completed. Tasks exported: {0}, Total records: {1}" -f $tasksExported, ($tracker.TotalExported -as [int])
    Write-ExportLog -Message $summaryMsg -Level Success
    Write-ProgressEntry -Path $progressLogPath -Message $summaryMsg
}

function Invoke-ContentExplorerResume {
    <#
    .SYNOPSIS
        Resumes an incomplete Content Explorer export from its last phase.
    .DESCRIPTION
        Reads ExportPhase.txt and task CSVs to determine progress, then continues
        from where the previous export left off. When WorkerCount > 0, spawns workers
        and dispatches remaining detail tasks via the file-drop protocol.
    .PARAMETER ExportDir
        Path to the export run directory containing ExportPhase.txt.
    .PARAMETER WorkerCount
        Number of worker terminals to spawn for multi-terminal resume.
        0 = single-terminal (default, existing behavior).
    #>
    param(
        [Parameter(Mandatory)]
        [string]$ExportDir,

        [int]$WorkerCount = 0
    )

    # Read current phase
    $phase = Read-ExportPhase -ExportDir $ExportDir
    if (-not $phase) {
        Write-Host "  ERROR: No ExportPhase.txt found in $ExportDir" -ForegroundColor Red
        return
    }

    # Display resume summary - build dynamic lines
    $resumeLines = @(
        ("Export: {0}" -f (Split-Path $ExportDir -Leaf)),
        ("Phase:  {0}" -f $phase)
    )

    # Read task CSVs for progress
    $aggCsvPath = Join-Path (Get-CoordinationDir $ExportDir) "AggregateTasks.csv"
    $detCsvPath = Join-Path (Get-CoordinationDir $ExportDir) "DetailTasks.csv"
    $aggTasks = @()
    $detTasks = @()

    if (Test-Path $aggCsvPath) {
        $aggTasks = @(Read-TaskCsv -Path $aggCsvPath)
        $aggDone = @($aggTasks | Where-Object { $_.Status -eq "Completed" }).Count
        $aggErr = @($aggTasks | Where-Object { $_.Status -eq "Error" }).Count
        $resumeLines += ("Aggregate: {0}/{1} done, {2} errors" -f $aggDone, $aggTasks.Count, $aggErr)
    }
    if (Test-Path $detCsvPath) {
        $detTasks = @(Read-TaskCsv -Path $detCsvPath)
        $detDone = @($detTasks | Where-Object { $_.Status -eq "Completed" }).Count
        $detErr = @($detTasks | Where-Object { $_.Status -eq "Error" }).Count
        $resumeLines += ("Detail:    {0}/{1} done, {2} errors" -f $detDone, $detTasks.Count, $detErr)
    }

    # Check last activity from ExportPhase.txt timestamp
    $phaseFile = Join-Path (Get-CoordinationDir $ExportDir) "ExportPhase.txt"
    if (Test-Path $phaseFile) {
        $lastWrite = (Get-Item $phaseFile).LastWriteTime
        $elapsed = (Get-Date) - $lastWrite
        $agoText = if ($elapsed.TotalHours -lt 1) { "{0} minutes ago" -f [int]$elapsed.TotalMinutes }
                   elseif ($elapsed.TotalHours -lt 24) { "{0} hours ago" -f [math]::Round($elapsed.TotalHours, 1) }
                   else { "{0} days ago" -f [math]::Round($elapsed.TotalDays, 1) }
        $resumeLines += ("Last activity: {0}" -f $agoText)
    }

    Write-Host ""
    Write-Banner -Title 'CONTENT EXPLORER - RESUME MODE' -Lines $resumeLines -Color 'Cyan'

    # Confirm
    Write-Host ""
    $confirm = Read-Host "  Resume this export? [Y/n]"
    if (-not [string]::IsNullOrEmpty($confirm) -and $confirm.Trim().ToUpper() -ne "Y") {
        Write-Host "  Resume cancelled." -ForegroundColor Yellow
        return
    }

    # Initialize logging
    Initialize-ExportLog -LogDirectory (Get-LogsDir $ExportDir)
    $script:ExportRunDirectory = $ExportDir
    $script:SharedExportDirectory = $ExportDir
    $script:ErrorLogPath = Join-Path (Get-LogsDir $ExportDir) "ExportProject-Errors.log"

    Write-ExportLog -Message ("Resuming export from phase: {0}" -f $phase) -Level Info

    # Load Content Explorer page size (manifest overrides config file for consistency)
    $configPath = Join-Path $PSScriptRoot "ConfigFiles" "ContentExplorerClassifiers.json"
    $resolved = Resolve-CEPageSize -ExportRunDirectory $ExportDir -ConfigPath $configPath -FallbackPageSize $PageSize
    $cePageSize = $resolved.PageSize

    $sitsToSkipPath = Join-Path $PSScriptRoot "ConfigFiles" "SITstoSkip.json"
    $sitsToSkip = Get-SITsToSkip -ConfigPath $sitsToSkipPath

    $progressLogPath = Join-Path (Get-LogsDir $ExportDir) "ContentExplorer-Progress.log"
    $aggregateCsvPath = Join-Path (Get-CoordinationDir $ExportDir) "ContentExplorer-Aggregates.csv"
    $trackerPath = Join-Path (Get-CoordinationDir $ExportDir) "RunTracker.json"
    $tracker = Get-ContentExplorerRunTracker -TrackerPath $trackerPath
    $telemetryDbPath = Join-Path $PSScriptRoot "TelemetryDB" "content-explorer-telemetry.jsonl"

    # Route based on current phase
    switch ($phase) {
        "Aggregate" {
            Write-ExportLog -Message "Resuming from AGGREGATE phase" -Level Info

            # Reset any InProgress or Error tasks to Pending (from crashed workers or transient failures)
            $changed = $false
            foreach ($t in $aggTasks) {
                if ($t.Status -in @("InProgress", "Error")) {
                    $t.Status = "Pending"
                    $t.ErrorMessage = ""
                    $changed = $true
                }
            }
            if ($changed -and $aggTasks.Count -gt 0) {
                Write-TaskCsv -Path $aggCsvPath -Tasks $aggTasks
                Write-ExportLog -Message "Reset stale InProgress/Error aggregate tasks to Pending" -Level Info
            }

            # Run pending aggregate tasks directly
            $pendingAgg = @($aggTasks | Where-Object { $_.Status -eq "Pending" })
            Write-ExportLog -Message ("Running {0} pending aggregate tasks..." -f $pendingAgg.Count) -Level Info

            foreach ($aggTask in $pendingAgg) {
                $taskKey = "{0}|{1}|{2}" -f $aggTask.TagType, $aggTask.TagName, $aggTask.Workload
                Write-ExportLog -Message ("  Aggregate: {0}" -f $taskKey) -Level Info

                try {
                    $allAggregates = @()
                    $pageCookie = $null
                    do {
                        $aggParams = @{
                            TagType     = $aggTask.TagType
                            TagName     = $aggTask.TagName
                            Workload    = $aggTask.Workload
                            PageSize    = 5000
                            Aggregate   = $true
                            ErrorAction = 'Stop'
                        }
                        if ($pageCookie) { $aggParams['PageCookie'] = $pageCookie }

                        $aggResult = Export-ContentExplorerData @aggParams
                        if ($null -eq $aggResult -or $aggResult.Count -eq 0) { break }

                        $metadata = $aggResult[0]
                        if ($metadata.RecordsReturned -gt 0) {
                            $allAggregates += $aggResult[1..$metadata.RecordsReturned]
                        }
                        if ($metadata.MorePagesAvailable -eq $true -or $metadata.MorePagesAvailable -eq "True") {
                            $pageCookie = $metadata.PageCookie
                        } else { break }
                    } while ($true)

                    # Write aggregate results to central CSV
                    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                    foreach ($agg in $allAggregates) {
                        $locationName = $agg.Name -replace '"', '""'
                        if ($locationName -match '[,"]') { $locationName = ('"' + $locationName + '"') }
                        $csvLine = "{0},{1},{2},{3},{4},{5}," -f $timestamp, $aggTask.TagType, $aggTask.TagName, $aggTask.Workload, $locationName, $agg.Count
                        [System.IO.File]::AppendAllText($aggregateCsvPath, ($csvLine + "`r`n"), [System.Text.Encoding]::UTF8)
                    }

                    $matchCount = ($allAggregates | Measure-Object -Property Count -Sum).Sum
                    if (-not $matchCount) { $matchCount = 0 }

                    # Probe detail API for actual file count
                    $fileCount = $matchCount
                    if ($matchCount -gt 0) {
                        try {
                            $probeResult = Export-ContentExplorerData -TagType $aggTask.TagType -TagName $aggTask.TagName -Workload $aggTask.Workload -PageSize 1 -ErrorAction Stop
                            if ($probeResult -and $probeResult[0] -and $null -ne $probeResult[0].TotalCount) {
                                $probedCount = $probeResult[0].TotalCount -as [int]
                                if ($probedCount -and $probedCount -gt 0) {
                                    $fileCount = $probedCount
                                    if ($fileCount -ne $matchCount) {
                                        Write-ExportLog -Message ("    File count probe: {0} files vs {1} matches" -f $fileCount, $matchCount) -Level Info
                                    }
                                }
                            }
                        }
                        catch {
                            Write-ExportLog -Message ("    File count probe failed (using match count): {0}" -f $_.Exception.Message) -Level Warning
                        }
                    }

                    # Write _FILECOUNT row to central CSV
                    $fcLine = "{0},{1},{2},{3},_FILECOUNT,{4}," -f $timestamp, $aggTask.TagType, $aggTask.TagName, $aggTask.Workload, $fileCount
                    [System.IO.File]::AppendAllText($aggregateCsvPath, ($fcLine + "`r`n"), [System.Text.Encoding]::UTF8)

                    $aggTask.Status = "Completed"
                    Write-ExportLog -Message ("    -> {0} files ({1} matches) in {2} locations" -f $fileCount, $matchCount, $allAggregates.Count) -Level Success
                }
                catch {
                    $aggTask.Status = "Error"
                    $aggTask.ErrorMessage = $_.Exception.Message
                    Write-ExportLog -Message ("    AGGREGATE FAILED: {0}" -f $_.Exception.Message) -Level Error
                }

                Write-TaskCsv -Path $aggCsvPath -Tasks $aggTasks
            }

            # Fall through to Detail
            Write-ExportPhase -ExportDir $ExportDir -Phase "Detail"
            $phase = "Detail"
        }

        { $_ -in @("Detail") } {
            # Build or reload work plan from aggregate CSV
            if (Test-Path $aggregateCsvPath) {
                Write-ExportLog -Message "Loading aggregate data for detail task planning..." -Level Info

                # Get tag type/name combos from the aggregate CSV
                $aggCsvContent = Import-Csv -Path $aggregateCsvPath -ErrorAction SilentlyContinue
                if (-not $aggCsvContent) {
                    # Try reading with header
                    $aggCsvContent = @()
                    $csvLines = [System.IO.File]::ReadAllLines($aggregateCsvPath)
                    if ($csvLines.Count -gt 1) {
                        # Parse manually if Import-Csv fails
                        Write-ExportLog -Message "Using Import-AggregateDataFromCsv for aggregate data" -Level Info
                    }
                }

                # Use module function to import aggregate data
                $tagTypes = @($aggTasks | Select-Object -ExpandProperty TagType -Unique)
                $workloads = @($aggTasks | Select-Object -ExpandProperty Workload -Unique) | Select-Object -Unique

                $script:AllWorkPlanTasks = @()
                foreach ($tagType in $tagTypes) {
                    $tagNames = @($aggTasks | Where-Object { $_.TagType -eq $tagType } | Select-Object -ExpandProperty TagName -Unique)
                    if ($tagNames.Count -eq 0) { continue }
                    $cachedResult = Import-AggregateDataFromCsv -CsvPath $aggregateCsvPath -TagType $tagType -TagNames $tagNames -Workloads $workloads
                    if ($cachedResult -and $cachedResult.Tasks) {
                        $script:AllWorkPlanTasks += @($cachedResult.Tasks)
                    }
                }

                Write-ExportLog -Message ("Built work plan: {0} detail tasks" -f $script:AllWorkPlanTasks.Count) -Level Info
            }
            else {
                Write-ExportLog -Message "No aggregate CSV found - cannot build detail task list" -Level Error
                return
            }

            # If we have existing DetailTasks.csv, check which are already done
            $completedTaskKeys = @{}
            if ($detTasks.Count -gt 0) {
                foreach ($dt in $detTasks) {
                    if ($dt.Status -eq "Completed") {
                        $dtKey = "{0}|{1}|{2}" -f $dt.TagType, $dt.TagName, $dt.Workload
                        $completedTaskKeys[$dtKey] = $true
                    }
                }
                Write-ExportLog -Message ("Found {0} already-completed detail tasks" -f $completedTaskKeys.Count) -Level Info
            }

            Write-ExportPhase -ExportDir $ExportDir -Phase "Detail"

            if ($WorkerCount -gt 0) {
                # -- Multi-terminal resume: spawn workers and dispatch remaining detail tasks --
                Write-ExportLog -Message ("Multi-terminal resume: spawning {0} worker(s) for detail phase" -f $WorkerCount) -Level Info

                # Build detail tasks: prefer DetailTasks.csv (already has Location/LocationType)
                $detCsvPath2 = Join-Path (Get-CoordinationDir $ExportDir) "DetailTasks.csv"
                if ($detTasks.Count -gt 0) {
                    # Use existing per-location tasks from DetailTasks.csv
                    $detailTasks = @($detTasks)
                    # Reset InProgress/Error tasks to Pending (from crashed workers or transient failures)
                    foreach ($t in $detailTasks) {
                        if ($t.Status -in @("InProgress", "Error")) {
                            $t.Status = "Pending"
                            $t.AssignedPID = 0
                            $t.ErrorMessage = ""
                        }
                    }
                    Write-ExportLog -Message ("  Loaded {0} detail tasks from existing DetailTasks.csv" -f $detailTasks.Count) -Level Info
                }
                else {
                    # Fallback: rebuild from work plan (no DetailTasks.csv available)
                    $detailTasks = @()
                    foreach ($task in $script:AllWorkPlanTasks) {
                        $taskKey = "{0}|{1}|{2}" -f $task.TagType, $task.TagName, $task.Workload
                        if ($completedTaskKeys.ContainsKey($taskKey)) { continue }
                        if ((-not $task.ExpectedCount -or $task.ExpectedCount -eq 0) -and $task.Status -ne "Error") { continue }

                        $detailTasks += @{
                            TagType               = $task.TagType
                            TagName               = $task.TagName
                            Workload              = $task.Workload
                            Location              = ""
                            LocationType          = ""
                            ExpectedCount         = ($task.ExpectedCount -as [int])
                            OriginalExpectedCount = if ($task.OriginalExpectedCount) { ($task.OriginalExpectedCount -as [int]) } else { ($task.ExpectedCount -as [int]) }
                            PageSize              = if ($task.PageSize) { ($task.PageSize -as [int]) } else { $cePageSize }
                            AssignedPID           = 0
                            Status                = "Pending"
                            ErrorMessage          = ""
                        }
                    }
                    Write-ExportLog -Message ("  Rebuilt {0} detail tasks from work plan (no DetailTasks.csv)" -f $detailTasks.Count) -Level Info
                    # Sort tasks: largest first for optimal scheduling (heavy tasks start early, small tasks fill gaps)
                    $detailTasks = @($detailTasks | Sort-Object { [int]$_.ExpectedCount } -Descending)
                    Write-ExportLog -Message "  Sorted detail tasks by ExpectedCount descending (largest first)" -Level Info
                }

                Write-TaskCsv -Path (Join-Path (Get-CoordinationDir $ExportDir) "DetailTasks.csv") -Tasks $detailTasks
                $pendingDetailTasks = @($detailTasks | Where-Object { $_.Status -eq "Pending" })
                Write-ExportLog -Message ("  {0} detail tasks total, {1} pending" -f $detailTasks.Count, $pendingDetailTasks.Count) -Level Info

                if ($pendingDetailTasks.Count -eq 0) {
                    Write-ExportLog -Message "  No pending detail tasks - skipping to completion" -Level Info
                }
                else {
                    # Spawn workers
                    $workerProcesses = Start-WorkerTerminals -ExportRunDirectory $ExportDir -Count $WorkerCount
                    if ($workerProcesses.Count -eq 0) {
                        Write-ExportLog -Message "  No workers spawned - aborting multi-terminal resume" -Level Error
                        Write-Host "  ERROR: No workers were spawned. Re-run without -WorkerCount for single-terminal mode." -ForegroundColor Red
                        return
                    }

                    # Dispatch via Invoke-DispatchLoop engine
                    $detailTaskCsvPath = Join-Path (Get-CoordinationDir $ExportDir) "DetailTasks.csv"
                    $totalDetailItems = ($detailTasks | ForEach-Object { ($_.ExpectedCount -as [long]) } | Measure-Object -Sum).Sum
                    if (-not $totalDetailItems) { $totalDetailItems = 0 }
                    Reset-OrchestratorDashboard
                    $detailPhaseStartTime = Get-Date
                    $nextWorkerNumber = $workerProcesses.Count + 1

                    # Baseline: tasks already completed before this resume session (for ETA accuracy)
                    $preResumeCompleted = @($detailTasks | Where-Object { $_.Status -eq "Completed" }).Count
                    $preResumeCompletedItems = ($detailTasks | Where-Object { $_.Status -eq "Completed" } | ForEach-Object { ($_.ExpectedCount -as [long]) } | Measure-Object -Sum).Sum
                    if (-not $preResumeCompletedItems) { $preResumeCompletedItems = 0 }

                    Write-ExportLog -Message "  Entering resume dispatch loop..." -Level Info

                    # Build tasks ArrayList for the engine
                    $dispatchTasks = [System.Collections.ArrayList]::new()
                    foreach ($t in $detailTasks) {
                        [void]$dispatchTasks.Add($t)
                    }

                    # --- Shared CE Detail Callback: OnScanCompletions ---
                    $ceDetailOnScan = {
                        param($ExportDir, $WorkerDirs, $Context)
                        $completed = @()
                        $errors = @()
                        $completionsDir = Get-CompletionsDir $ExportDir
                        $signingKey = Get-ExportRunSigningKey -ExportDir $ExportDir

                        # Scan central Completions/ directory
                        $doneSignalFiles = @(Get-ChildItem -Path $completionsDir -Filter "detail-done-*.txt" -File -ErrorAction SilentlyContinue)
                        $errorSignalFiles = @(Get-ChildItem -Path $completionsDir -Filter "error-detail-*.txt" -File -ErrorAction SilentlyContinue)

                        # Also scan worker dirs (backward compat)
                        foreach ($wDir in $WorkerDirs) {
                            $doneSignalFiles += @(Get-ChildItem -Path $wDir -Filter "detail-done-*.txt" -File -ErrorAction SilentlyContinue)
                            $doneSignalFiles += @(Get-ChildItem -Path $wDir -Filter "done-detail-*.txt" -File -ErrorAction SilentlyContinue)
                            $errorSignalFiles += @(Get-ChildItem -Path $wDir -Filter "error-detail-*.txt" -File -ErrorAction SilentlyContinue)
                        }

                        foreach ($doneFile in $doneSignalFiles) {
                            try {
                                $doneContent = [System.IO.File]::ReadAllText($doneFile.FullName)
                                $doneData = ConvertFrom-SignedEnvelopeJson -Json $doneContent -SigningKey $signingKey -RequireSignature:([bool]$signingKey) -Context ("detail completion file {0}" -f $doneFile.Name)
                                if ($null -eq $doneData) {
                                    Write-ExportLog -Message ("  Warning: Empty/null detail done file {0}" -f $doneFile.Name) -Level Warning -LogOnly
                                    Remove-Item -Path $doneFile.FullName -Force -ErrorAction SilentlyContinue
                                    continue
                                }
                                $doneLocation = if ($doneData.Location) { $doneData.Location } else { "" }
                                $completed += @{
                                    TagType     = $doneData.TagType
                                    TagName     = $doneData.TagName
                                    Workload    = $doneData.Workload
                                    Location    = $doneLocation
                                    RecordCount = $doneData.RecordCount
                                    Message     = "{0}/{1}{2} -> {3} records" -f $doneData.TagName, $doneData.Workload, $(if ($doneLocation) { "/$doneLocation" } else { "" }), $doneData.RecordCount
                                }
                                Remove-Item -Path $doneFile.FullName -Force -ErrorAction SilentlyContinue
                            }
                            catch {
                                Write-ExportLog -Message ("  Warning: Could not parse detail done file {0}" -f $doneFile.Name) -Level Warning -LogOnly
                            }
                        }

                        foreach ($errFile in $errorSignalFiles) {
                            try {
                                $errContent = [System.IO.File]::ReadAllText($errFile.FullName)
                                $errData = ConvertFrom-SignedEnvelopeJson -Json $errContent -SigningKey $signingKey -RequireSignature:([bool]$signingKey) -Context ("detail error file {0}" -f $errFile.Name)
                                if ($null -eq $errData) {
                                    Write-ExportLog -Message ("  Warning: Empty/null detail error file {0}" -f $errFile.Name) -Level Warning -LogOnly
                                    Remove-Item -Path $errFile.FullName -Force -ErrorAction SilentlyContinue
                                    continue
                                }
                                $errLocation = if ($errData.Location) { $errData.Location } else { "" }
                                $errors += @{
                                    TagType      = $errData.TagType
                                    TagName      = $errData.TagName
                                    Workload     = $errData.Workload
                                    Location     = $errLocation
                                    ErrorMessage = $errData.Error
                                    Message      = "{0}/{1}{2}: {3}" -f $errData.TagName, $errData.Workload, $(if ($errLocation) { "/$errLocation" } else { "" }), $errData.Error
                                }
                                Remove-Item -Path $errFile.FullName -Force -ErrorAction SilentlyContinue
                            }
                            catch {
                                Write-ExportLog -Message ("  Warning: Could not parse error file {0}" -f $errFile.Name) -Level Warning -LogOnly
                            }
                        }

                        return @{ CompletedTasks = $completed; ErrorTasks = $errors }
                    }

                    # --- Shared CE Detail Callback: OnMatchTask ---
                    $ceDetailOnMatch = {
                        param($Data, $Tasks, $Context)
                        $loc = if ($Data.Location) { $Data.Location } else { "" }
                        $match = $Tasks | Where-Object {
                            $_.TagType -eq $Data.TagType -and
                            $_.TagName -eq $Data.TagName -and
                            $_.Workload -eq $Data.Workload -and
                            $(if ($_.Location) { $_.Location } else { "" }) -eq $loc -and
                            $_.Status -eq "InProgress"
                        } | Select-Object -First 1
                        if (-not $match -and -not $loc) {
                            $match = $Tasks | Where-Object {
                                $_.TagType -eq $Data.TagType -and
                                $_.TagName -eq $Data.TagName -and
                                $_.Workload -eq $Data.Workload -and
                                $_.Status -eq "InProgress"
                            } | Select-Object -First 1
                        }
                        return $match
                    }

                    # --- CE Resume Callback: OnDispatchTask ---
                    $ceResumeOnDispatch = {
                        param($Worker, $NextTask, $Context)
                        $taskData = @{
                            Phase         = "Detail"
                            TagType       = $NextTask.TagType
                            TagName       = $NextTask.TagName
                            Workload      = $NextTask.Workload
                            Location      = if ($NextTask.Location) { $NextTask.Location } else { "" }
                            LocationType  = if ($NextTask.LocationType) { $NextTask.LocationType } else { "" }
                            ExpectedCount = ($NextTask.ExpectedCount -as [int])
                            PageSize      = ($NextTask.PageSize -as [int])
                        }
                        $sent = Send-WorkerTask -WorkerDir $Worker.WorkerDir -TaskData $taskData -ExportDir $Context.ExportDir
                        if ($sent) {
                            $Context.DispatchTimes[$Worker.PID] = Get-Date
                        }
                        return $sent
                    }

                    # --- CE Resume Callback: OnShowDashboard ---
                    $ceResumeOnDashboard = {
                        param($LoopState, $Context)
                        $tasks = $Context.DetailTasks
                        $workerStatusList = @()
                        $completedItems = [long]0
                        foreach ($dt in $tasks) {
                            if ($dt.Status -eq "Completed") { $completedItems += ($dt.ExpectedCount -as [long]) }
                        }

                        foreach ($wp in $LoopState.WorkerProcesses) {
                            $wPID = $wp.PID
                            $wDir = $wp.WorkerDir
                            $wState = Get-WorkerState -WorkerDir $wDir -WorkerPID $wPID
                            $currentTaskName = "-"
                            $taskTimeStr = "-"
                            $expectedStr = "-"
                            $progressStr = "-"
                            $pageSizeStr = "-"
                            $lastPageTimeStr = "-"
                            $ctData = $null

                            try {
                                $ctPath = Join-Path $wDir "currenttask"
                                $ctData = [System.IO.File]::ReadAllText($ctPath) | ConvertFrom-Json
                                if ($null -ne $ctData) {
                                    if ($ctData.TagName -and $ctData.Workload) { $currentTaskName = "{0}/{1}" -f $ctData.TagName, $ctData.Workload }
                                    if ($ctData.ExpectedCount -and ($ctData.ExpectedCount -as [int]) -gt 0) { $expectedStr = ($ctData.ExpectedCount -as [int]).ToString('N0') }
                                    if ($ctData.PageSize -and ($ctData.PageSize -as [int]) -gt 0) { $pageSizeStr = ($ctData.PageSize -as [int]).ToString('N0') }
                                }
                            } catch { }

                            if ($wState -eq "Busy") {
                                if ($Context.DispatchTimes.ContainsKey($wPID)) {
                                    $taskTimeStr = Format-TimeSpan -Seconds ((Get-Date) - $Context.DispatchTimes[$wPID]).TotalSeconds
                                }
                                $wpExpected = if ($ctData) { $ctData.ExpectedCount -as [int] } else { 0 }
                                try {
                                    $wProgressLog = Join-Path $wDir "Progress.log"
                                    if (Test-Path $wProgressLog) {
                                        $lastLines = Get-Content -Path $wProgressLog -Tail 5 -ErrorAction Stop
                                        for ($li = $lastLines.Count - 1; $li -ge 0; $li--) {
                                            if ($lastLines[$li] -match 'Total:\s*([\d,]+)/.*\[(\d+)ms\]') {
                                                $currentCount = ($Matches[1] -replace ',', '') -as [int]
                                                if ($null -ne $currentCount) {
                                                    $progressStr = $currentCount.ToString('N0')
                                                    if ($wpExpected -and $wpExpected -gt 0) {
                                                        $pctDone = [Math]::Round(($currentCount / $wpExpected) * 100)
                                                        $progressStr += " ({0}%)" -f $pctDone
                                                    }
                                                }
                                                $pageTimeMs = $Matches[2] -as [int]
                                                if ($null -ne $pageTimeMs) {
                                                    if ($pageTimeMs -ge 60000) { $lastPageTimeStr = "{0:N1}m" -f ($pageTimeMs / 60000) }
                                                    elseif ($pageTimeMs -ge 1000) { $lastPageTimeStr = "{0:N1}s" -f ($pageTimeMs / 1000) }
                                                    else { $lastPageTimeStr = "${pageTimeMs}ms" }
                                                }
                                                break
                                            }
                                        }
                                    }
                                } catch { }
                            }

                            $workerStatusList += @{
                                PID = $wPID; State = $wState; CurrentTask = $currentTaskName
                                Expected = $expectedStr; Progress = $progressStr; PageSize = $pageSizeStr
                                LastPage = $lastPageTimeStr; TaskTime = $taskTimeStr
                            }
                        }

                        # Dynamic worker spawning via W hotkey
                        Test-AddWorkerKeypress -ExportRunDirectory $Context.ExportDir `
                            -WorkerProcesses $Context.WorkerProcessesRef -NextWorkerNumber $Context.NextWorkerNumberRef

                        try {
                            Show-OrchestratorDashboard -Phase "Detail" `
                                -Completed $LoopState.CompletedCount -Total $LoopState.TotalCount `
                                -Workers $workerStatusList -RecentErrors $LoopState.RecentErrors `
                                -RecentActivity $LoopState.RecentActivity -DispatchLog @() `
                                -ExportStartTime $Context.ExportStartTime -PhaseStartTime $Context.PhaseStartTime `
                                -CompletedItems $completedItems -TotalItems $Context.TotalDetailItems `
                                -DetailTasks $tasks `
                                -CompletedBaseline $Context.PreResumeCompleted -CompletedItemsBaseline $Context.PreResumeCompletedItems
                        } catch {
                            Write-ExportLog -Message ("  Dashboard render error (non-fatal): {0}" -f $_.Exception.Message) -Level Warning -LogOnly
                        }
                    }

                    # --- CE Resume Callback: OnAllWorkersDead ---
                    $ceResumeOnAllDead = {
                        param($Tasks, $PendingCount, $Context)
                        Write-ExportLog -Message "All CE resume workers dead - saving state" -Level Error
                        Write-TaskCsv -Path $Context.TaskCsvPath -Tasks $Context.DetailTasks
                    }

                    # --- CE Resume Callback: OnIterationComplete ---
                    $ceResumeOnIterComplete = {
                        param($Tasks, $LoopState, $Context)
                        # Batched CSV write: only when tasks have changed (engine modifies Status/AssignedPID)
                        Write-TaskCsv -Path $Context.TaskCsvPath -Tasks $Context.DetailTasks
                    }

                    # Build context hashtable
                    $resumeContext = @{
                        ExportDir             = $ExportDir
                        TaskCsvPath           = $detailTaskCsvPath
                        DetailTasks           = $detailTasks
                        TotalDetailItems      = $totalDetailItems
                        ExportStartTime       = $detailPhaseStartTime
                        PhaseStartTime        = $detailPhaseStartTime
                        PreResumeCompleted    = $preResumeCompleted
                        PreResumeCompletedItems = $preResumeCompletedItems
                        DispatchTimes         = @{}
                        WorkerProcessesRef    = ([ref]$workerProcesses)
                        NextWorkerNumberRef   = ([ref]$nextWorkerNumber)
                    }

                    $resumeLoopResult = Invoke-DispatchLoop `
                        -ExportDir $ExportDir `
                        -Tasks $dispatchTasks `
                        -WorkerProcesses $workerProcesses `
                        -Context $resumeContext `
                        -OnScanCompletions $ceDetailOnScan `
                        -OnMatchTask $ceDetailOnMatch `
                        -OnDispatchTask $ceResumeOnDispatch `
                        -OnShowDashboard $ceResumeOnDashboard `
                        -OnAllWorkersDead $ceResumeOnAllDead `
                        -OnIterationComplete $ceResumeOnIterComplete `
                        -SleepSeconds 2

                    # Save final task state
                    Write-TaskCsv -Path $detailTaskCsvPath -Tasks $detailTasks

                    $doneDetailCount = @($detailTasks | Where-Object { $_.Status -in @("Completed", "Error") }).Count
                    $errorDetailCount = @($detailTasks | Where-Object { $_.Status -eq "Error" }).Count
                    Write-ExportLog -Message ("  Resume detail export complete: {0}/{1} tasks done ({2} errors)" -f $doneDetailCount, $detailTasks.Count, $errorDetailCount) -Level Success
                }
            }
            else {
                # -- Single-terminal resume: process detail tasks directly --
                $completedTaskCounts = @{}
                foreach ($task in $script:AllWorkPlanTasks) {
                    $taskKey = "{0}|{1}|{2}" -f $task.TagType, $task.TagName, $task.Workload

                    if ($completedTaskKeys.ContainsKey($taskKey)) {
                        Write-ExportLog -Message ("    Skipping {0} / {1} - already completed" -f $task.TagName, $task.Workload) -Level Info
                        continue
                    }

                    if ((-not $task.ExpectedCount -or $task.ExpectedCount -eq 0) -and $task.Status -ne "Error") {
                        Write-ExportLog -Message ("    Skipping {0} / {1} - no data" -f $task.TagName, $task.Workload) -Level Info
                        continue
                    }

                    # Run export — output to Data/ContentExplorer/TagType/TagName/
                    $classifierDir = Get-CEClassifierDir $ExportDir $task.TagType $task.TagName

                    $telemetry = New-ContentExplorerTelemetry -TagType $task.TagType -TagName $task.TagName -Workload $task.Workload
                    $exportParams = @{
                        Task                 = $task
                        PageSize             = $cePageSize
                        ProgressLogPath      = $progressLogPath
                        Telemetry            = $telemetry
                        TelemetryDatabasePath = $telemetryDbPath
                        AdaptivePageSize     = $true
                        OutputDirectory      = $classifierDir
                    }

                    try {
                        Export-ContentExplorerWithProgress @exportParams | Out-Null
                        $exportedCount = if ($task.ExportedCount) { $task.ExportedCount } else { 0 }

                        if ($exportedCount -gt 0) {
                            Write-ExportLog -Message ("    Completed: {0} / {1} - {2} records" -f $task.TagName, $task.Workload, $exportedCount) -Level Success
                        }

                        $completedTaskCounts[$taskKey] = $exportedCount
                    }
                    catch {
                        Write-ExportLog -Message ("    FAILED: {0} / {1} - {2}" -f $task.TagName, $task.Workload, $_.Exception.Message) -Level Error
                        if ($script:ErrorLogPath) {
                            Write-ExportErrorLog -ErrorLogPath $script:ErrorLogPath -Context "Resume Detail Export" -TaskKey $taskKey -ErrorRecord $_
                        }
                    }

                    Save-ContentExplorerRunTracker -Tracker $tracker -TrackerPath $trackerPath
                }
            }

            # Retry bucket detection before completion
            $retryBucketTasks = @()
            $detailCsvForRetry = Join-Path (Get-CoordinationDir $ExportDir) "DetailTasks.csv"
            if (Test-Path $detailCsvForRetry) {
                $finalDetailTasks = Read-TaskCsv -Path $detailCsvForRetry
                $retryBucketTasks = @(Get-RetryBucketTasks -DetailTasks $finalDetailTasks)
                if ($retryBucketTasks.Count -gt 0) {
                    $retryTasksCsvPath = Join-Path (Get-CoordinationDir $ExportDir) "RetryTasks.csv"
                    Write-RetryTasksCsv -Path $retryTasksCsvPath -RetryTasks $retryBucketTasks
                    Write-ExportLog -Message ("  Wrote {0} retry tasks to RetryTasks.csv" -f $retryBucketTasks.Count) -Level Info
                }
            }

        }
    }

    # Display retry bucket summary
    if (-not $retryBucketTasks) { $retryBucketTasks = @() }
    Show-RetryBucketSummary -RetryTasks $retryBucketTasks -ExportDir $ExportDir

    # Write remaining (non-completed) tasks for follow-on runs
    $remainingCount = Write-RemainingTasksCsv -ExportDir $ExportDir
    if ($remainingCount -gt 0) {
        Write-ExportLog -Message ("  Remaining tasks: {0} (see RemainingTasks.csv)" -f $remainingCount) -Level Warning
        Write-ExportLog -Message ("  To re-run: .\Export-Compl8Configuration.ps1 -CETasksCsv ""{0}""" -f (Join-Path (Get-CoordinationDir $ExportDir) "RemainingTasks.csv")) -Level Info
    }

    # Write manifest and set final phase
    Write-CEManifest -ExportDir $ExportDir
    Write-ExportPhase -ExportDir $ExportDir -Phase "Completed"
    Write-ExportLog -Message "Resume complete" -Level Success
}

function Invoke-ContentExplorerRetry {
    <#
    .SYNOPSIS
        Re-exports Content Explorer tasks that had >2% discrepancy between expected and actual counts.
    .DESCRIPTION
        Reads RetryTasks.csv from a previous export, re-runs detail export for those specific tasks
        (skipping aggregation), then re-evaluates retry buckets.
    .PARAMETER ExportDir
        Path to the export run directory containing RetryTasks.csv.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$ExportDir
    )

    # Validate RetryTasks.csv exists
    $retryTasksCsvPath = Join-Path (Get-CoordinationDir $ExportDir) "RetryTasks.csv"
    if (-not (Test-Path $retryTasksCsvPath)) {
        Write-Host "  ERROR: No RetryTasks.csv found in $ExportDir" -ForegroundColor Red
        return
    }

    # Load retry tasks
    $retryTasks = @(Import-Csv -Path $retryTasksCsvPath -Encoding UTF8)
    if ($retryTasks.Count -eq 0) {
        Write-Host "  No retry tasks found in RetryTasks.csv" -ForegroundColor Yellow
        return
    }

    # Display retry summary
    $retryLines = @(
        ("Export: {0}" -f (Split-Path $ExportDir -Leaf)),
        ("Tasks to retry: {0}" -f $retryTasks.Count)
    )
    foreach ($rt in $retryTasks) {
        $sign = if (($rt.DiscrepancyPct -as [double]) -ge 0) { "+" } else { "" }
        $retryLines += ("  {0} / {1}: {2}{3}%%" -f $rt.TagName, $rt.Workload, $sign, $rt.DiscrepancyPct)
    }
    Write-Host ""
    Write-Banner -Title 'CONTENT EXPLORER - RETRY MODE' -Lines $retryLines -Color 'Cyan'

    # Confirm
    Write-Host ""
    $confirm = Read-Host "  Retry these tasks? [Y/N]"
    if ([string]::IsNullOrEmpty($confirm) -or $confirm.Trim().ToUpper() -ne "Y") {
        Write-Host "  Retry cancelled." -ForegroundColor Yellow
        return
    }

    # Initialize logging
    Initialize-ExportLog -LogDirectory (Get-LogsDir $ExportDir)
    $script:ExportRunDirectory = $ExportDir
    $script:SharedExportDirectory = $ExportDir
    $script:ErrorLogPath = Join-Path (Get-LogsDir $ExportDir) "ExportProject-Errors.log"

    Write-ExportLog -Message ("Retrying {0} discrepant tasks" -f $retryTasks.Count) -Level Info

    # Load Content Explorer page size (manifest overrides config file for consistency)
    $configPath = Join-Path $PSScriptRoot "ConfigFiles" "ContentExplorerClassifiers.json"
    $resolved = Resolve-CEPageSize -ExportRunDirectory $ExportDir -ConfigPath $configPath -FallbackPageSize $PageSize
    $cePageSize = $resolved.PageSize

    $progressLogPath = Join-Path (Get-LogsDir $ExportDir) "ContentExplorer-Progress.log"
    $aggregateCsvPath = Join-Path (Get-CoordinationDir $ExportDir) "ContentExplorer-Aggregates.csv"
    $trackerPath = Join-Path (Get-CoordinationDir $ExportDir) "RunTracker.json"
    $tracker = Get-ContentExplorerRunTracker -TrackerPath $trackerPath
    $telemetryDbPath = Join-Path $PSScriptRoot "TelemetryDB" "content-explorer-telemetry.jsonl"

    # Build work plan tasks from retry CSV (reuse aggregate data, skip aggregation)
    $script:AllWorkPlanTasks = @()
    $workloads = @($script:CEDefaultWorkloads)

    foreach ($rt in $retryTasks) {
        # Try to get location data from existing aggregate CSV for adaptive page sizing
        $taskLocations = @()
        if (Test-Path $aggregateCsvPath) {
            try {
                $cachedResult = Import-AggregateDataFromCsv -CsvPath $aggregateCsvPath -TagType $rt.TagType -TagNames @($rt.TagName) -Workloads @($rt.Workload)
                $taskKey = "{0}|{1}|{2}" -f $rt.TagType, $rt.TagName, $rt.Workload
                if ($cachedResult.TaskData -and $cachedResult.TaskData[$taskKey]) {
                    $taskLocations = $cachedResult.TaskData[$taskKey].Locations
                }
            }
            catch {
                Write-ExportLog -Message ("  Could not load aggregate data for {0}/{1}: {2}" -f $rt.TagName, $rt.Workload, $_.Exception.Message) -Level Warning
            }
        }

        $script:AllWorkPlanTasks += @{
            TagType       = $rt.TagType
            TagName       = $rt.TagName
            Workload      = $rt.Workload
            ExpectedCount = ($rt.OriginalExpectedCount -as [int])
            ExportedCount = 0
            Locations     = $taskLocations
            Status        = "Pending"
        }
    }

    Write-ExportLog -Message ("  Built {0} retry tasks" -f $script:AllWorkPlanTasks.Count) -Level Info

    # Set phase to Detail
    Write-ExportPhase -ExportDir $ExportDir -Phase "Detail"

    # Run detail export for retry tasks (single-terminal)
    $completedTaskCounts = @{}

    foreach ($task in @($script:AllWorkPlanTasks)) {
        $taskKey = "{0}|{1}|{2}" -f $task.TagType, $task.TagName, $task.Workload

        if (-not $task.ExpectedCount -or $task.ExpectedCount -eq 0) {
            Write-ExportLog -Message ("    Skipping {0} / {1} - no data" -f $task.TagName, $task.Workload) -Level Info
            continue
        }

        Write-ExportLog -Message ("    Retrying: {0} / {1} (expected {2})" -f $task.TagName, $task.Workload, $task.ExpectedCount) -Level Info

        # Export task with progress tracking — output to Data/ContentExplorer/TagType/TagName/
        $classifierDir = Get-CEClassifierDir $ExportDir $task.TagType $task.TagName

        $exportParams = @{
            Task                  = $task
            PageSize              = $cePageSize
            ProgressLogPath       = $progressLogPath
            AdaptivePageSize      = $true
            TelemetryDatabasePath = $telemetryDbPath
            OutputDirectory       = $classifierDir
        }

        try {
            Export-ContentExplorerWithProgress @exportParams | Out-Null
            $exportedCount = if ($task.ExportedCount) { $task.ExportedCount } else { 0 }
            $completedTaskCounts[$taskKey] = $exportedCount

            if ($exportedCount -gt 0) {
                Write-ExportLog -Message ("    Completed: {0} / {1} - {2} records" -f $task.TagName, $task.Workload, $exportedCount) -Level Success
            }
            else {
                Write-ExportLog -Message ("    Completed: {0} / {1} - 0 records" -f $task.TagName, $task.Workload) -Level Info
            }
        }
        catch {
            Write-ExportLog -Message ("    FAILED: {0} / {1} - {2}" -f $task.TagName, $task.Workload, $_.Exception.Message) -Level Error
            if ($script:ErrorLogPath) {
                Write-ExportErrorLog -ErrorLogPath $script:ErrorLogPath -Context "Retry Detail Export" -TaskKey $taskKey -ErrorRecord $_
            }
        }
    }

    # Re-evaluate retry bucket after re-export
    $detailCsvPath = Join-Path (Get-CoordinationDir $ExportDir) "DetailTasks.csv"
    $remainingRetryTasks = @()
    if (Test-Path $detailCsvPath) {
        # Update the DetailTasks.csv with new actual counts for retried tasks
        $detailTasks = Read-TaskCsv -Path $detailCsvPath
        foreach ($rt in $retryTasks) {
            $taskKey = "{0}|{1}|{2}" -f $rt.TagType, $rt.TagName, $rt.Workload
            if ($completedTaskCounts.ContainsKey($taskKey)) {
                $matchTask = $detailTasks | Where-Object {
                    $_.TagType -eq $rt.TagType -and $_.TagName -eq $rt.TagName -and $_.Workload -eq $rt.Workload
                } | Select-Object -First 1
                if ($matchTask) {
                    $matchTask.ExpectedCount = $completedTaskCounts[$taskKey]
                }
            }
        }
        Write-TaskCsv -Path $detailCsvPath -Tasks $detailTasks
        $remainingRetryTasks = @(Get-RetryBucketTasks -DetailTasks $detailTasks)
    }

    # Update RetryTasks.csv
    if ($remainingRetryTasks.Count -gt 0) {
        Write-RetryTasksCsv -Path $retryTasksCsvPath -RetryTasks $remainingRetryTasks
    }
    else {
        # All tasks now pass - remove retry CSV
        Remove-Item -Path $retryTasksCsvPath -Force -ErrorAction SilentlyContinue
    }

    # Display final retry summary
    Show-RetryBucketSummary -RetryTasks $remainingRetryTasks -ExportDir $ExportDir

    # Write remaining (non-completed) tasks for follow-on runs
    $remainingCount = Write-RemainingTasksCsv -ExportDir $ExportDir
    if ($remainingCount -gt 0) {
        Write-ExportLog -Message ("  Remaining tasks: {0} (see RemainingTasks.csv)" -f $remainingCount) -Level Warning
        Write-ExportLog -Message ("  To re-run: .\Export-Compl8Configuration.ps1 -CETasksCsv ""{0}""" -f (Join-Path (Get-CoordinationDir $ExportDir) "RemainingTasks.csv")) -Level Info
    }

    # Write manifest and set final phase
    Write-CEManifest -ExportDir $ExportDir
    Write-ExportPhase -ExportDir $ExportDir -Phase "Completed"
    Write-ExportLog -Message "Retry complete" -Level Success
}

function Invoke-ContentExplorerFromTasksCsv {
    <#
    .SYNOPSIS
        Runs Content Explorer detail export from an input task CSV file.
    .DESCRIPTION
        Reads a DetailTasks-format CSV (e.g. RemainingTasks.csv from a prior run), validates
        the schema, resets Error/InProgress tasks to Pending, then executes detail export.
        Supports both single-terminal and multi-terminal (via WorkerCount).
        Runs retry bucket detection and remaining tasks output at end.
    .PARAMETER TasksCsvPath
        Path to the input CSV file (same 11-column schema as DetailTasks.csv).
    .PARAMETER WorkerCount
        Number of worker terminals to spawn. 0 = single-terminal (default).
    #>
    param(
        [Parameter(Mandatory)]
        [string]$TasksCsvPath,

        [int]$WorkerCount = 0
    )

    # Validate file exists
    if (-not (Test-Path $TasksCsvPath)) {
        Write-Host ("  ERROR: Task CSV not found: {0}" -f $TasksCsvPath) -ForegroundColor Red
        return
    }

    # Read and validate CSV schema
    $inputTasks = @(Read-TaskCsv -Path $TasksCsvPath)
    if ($inputTasks.Count -eq 0) {
        Write-Host "  ERROR: No tasks found in CSV" -ForegroundColor Red
        return
    }

    $requiredColumns = @("TagType", "TagName", "Workload", "ExpectedCount", "PageSize", "Status")
    $csvColumns = $inputTasks[0].PSObject.Properties.Name
    $missingColumns = @($requiredColumns | Where-Object { $_ -notin $csvColumns })
    if ($missingColumns.Count -gt 0) {
        Write-Host ("  ERROR: CSV missing required columns: {0}" -f ($missingColumns -join ", ")) -ForegroundColor Red
        Write-Host "  Expected DetailTasks.csv format with columns: TagType, TagName, Workload, Location, LocationType, ExpectedCount, PageSize, AssignedPID, Status, ErrorMessage, OriginalExpectedCount" -ForegroundColor Yellow
        return
    }

    # Display summary
    $byStatus = @{}
    foreach ($t in $inputTasks) {
        $s = if ($t.Status) { $t.Status } else { "Unknown" }
        if (-not $byStatus.ContainsKey($s)) { $byStatus[$s] = 0 }
        $byStatus[$s]++
    }

    $taskCsvLines = @(
        ("Source: {0}" -f (Split-Path $TasksCsvPath -Leaf)),
        ("Total tasks: {0}" -f $inputTasks.Count)
    )
    foreach ($s in ($byStatus.Keys | Sort-Object)) {
        $taskCsvLines += ("  {0}: {1}" -f $s, $byStatus[$s])
    }
    $totalExpected = ($inputTasks | ForEach-Object { ($_.ExpectedCount -as [long]) } | Measure-Object -Sum).Sum
    if ($totalExpected) {
        $taskCsvLines += ("Total expected records: {0}" -f $totalExpected.ToString('N0'))
    }
    if ($WorkerCount -gt 0) {
        $taskCsvLines += ("Workers: {0} (multi-terminal)" -f $WorkerCount)
    } else {
        $taskCsvLines += "Workers: Single terminal"
    }
    Write-Host ""
    Write-Banner -Title 'CONTENT EXPLORER - RUN FROM TASK CSV' -Lines $taskCsvLines -Color 'Cyan'

    # Confirm
    Write-Host ""
    $confirm = Read-Host "  Run these tasks? [Y/N]"
    if ([string]::IsNullOrEmpty($confirm) -or $confirm.Trim().ToUpper() -ne "Y") {
        Write-Host "  Cancelled." -ForegroundColor Yellow
        return
    }

    # Create a new export directory for this run
    $exportDir = $script:ExportRunDirectory

    # Initialize logging
    Initialize-ExportLog -LogDirectory (Get-LogsDir $exportDir)
    $script:SharedExportDirectory = $exportDir
    $script:ErrorLogPath = Join-Path (Get-LogsDir $exportDir) "ExportProject-Errors.log"

    Write-ExportLog -Message ("Running {0} tasks from CSV: {1}" -f $inputTasks.Count, $TasksCsvPath) -Level Info

    # Reset Error/InProgress tasks to Pending, clear ErrorMessage/AssignedPID
    $resetCount = 0
    foreach ($t in $inputTasks) {
        if ($t.Status -in @("Error", "InProgress")) {
            $t.Status = "Pending"
            $t.AssignedPID = 0
            $t.ErrorMessage = ""
            $resetCount++
        }
    }
    if ($resetCount -gt 0) {
        Write-ExportLog -Message ("  Reset {0} Error/InProgress tasks to Pending" -f $resetCount) -Level Info
    }

    # Ensure OriginalExpectedCount is populated
    foreach ($t in $inputTasks) {
        if (-not $t.OriginalExpectedCount -or ($t.OriginalExpectedCount -as [int]) -eq 0) {
            $t.OriginalExpectedCount = ($t.ExpectedCount -as [int])
        }
    }

    # Write DetailTasks.csv into the new export directory
    $detailCsvPath = Join-Path (Get-CoordinationDir $exportDir) "DetailTasks.csv"
    Write-TaskCsv -Path $detailCsvPath -Tasks $inputTasks

    # Set phase to Detail
    Write-ExportPhase -ExportDir $exportDir -Phase "Detail"

    # Load Content Explorer configuration
    $configPath = Join-Path $PSScriptRoot "ConfigFiles" "ContentExplorerClassifiers.json"
    $ceConfig = Read-JsonConfig -Path $configPath
    $cePageSize = if ($ceConfig -and $ceConfig._PageSize) { ($ceConfig._PageSize -as [int]) } else { $PageSize }
    if (-not $cePageSize -or $cePageSize -lt 1) { $cePageSize = 100 }

    $progressLogPath = Join-Path (Get-LogsDir $exportDir) "ContentExplorer-Progress.log"
    $trackerPath = Join-Path (Get-CoordinationDir $exportDir) "RunTracker.json"
    $tracker = Get-ContentExplorerRunTracker -TrackerPath $trackerPath
    $telemetryDbPath = Join-Path $PSScriptRoot "TelemetryDB" "content-explorer-telemetry.jsonl"

    # Build work plan tasks from input CSV
    $script:AllWorkPlanTasks = @()
    foreach ($t in $inputTasks) {
        if ($t.Status -eq "Completed") { continue }
        $script:AllWorkPlanTasks += @{
            TagType       = $t.TagType
            TagName       = $t.TagName
            Workload      = $t.Workload
            ExpectedCount = ($t.ExpectedCount -as [int])
            ExportedCount = 0
            Locations     = @()
            Status        = "Pending"
        }
    }

    Write-ExportLog -Message ("  {0} tasks to export" -f $script:AllWorkPlanTasks.Count) -Level Info

    if ($WorkerCount -gt 0) {
        # -- Multi-terminal: spawn workers and dispatch tasks --
        $detailTasks = @($inputTasks | Where-Object { $_.Status -eq "Pending" })
        # Sort: largest first
        $detailTasks = @($detailTasks | Sort-Object { [int]$_.ExpectedCount } -Descending)
        Write-TaskCsv -Path $detailCsvPath -Tasks $inputTasks

        $workerProcesses = Start-WorkerTerminals -ExportRunDirectory $exportDir -Count $WorkerCount
        if ($workerProcesses.Count -eq 0) {
            Write-ExportLog -Message "  No workers spawned - aborting" -Level Error
            Write-Host "  ERROR: No workers were spawned." -ForegroundColor Red
            return
        }

        # Dispatch via Invoke-DispatchLoop engine (reuses shared CE detail callbacks)
        $totalDetailItems = ($inputTasks | ForEach-Object { ($_.ExpectedCount -as [long]) } | Measure-Object -Sum).Sum
        if (-not $totalDetailItems) { $totalDetailItems = 0 }
        Reset-OrchestratorDashboard
        $detailPhaseStartTime = Get-Date
        $nextWorkerNumber = $workerProcesses.Count + 1

        # Baseline: tasks already completed before this session (for ETA accuracy)
        $preBaselineCompleted = @($inputTasks | Where-Object { $_.Status -eq "Completed" }).Count
        $preBaselineCompletedItems = ($inputTasks | Where-Object { $_.Status -eq "Completed" } | ForEach-Object { ($_.ExpectedCount -as [long]) } | Measure-Object -Sum).Sum
        if (-not $preBaselineCompletedItems) { $preBaselineCompletedItems = 0 }

        Write-ExportLog -Message "  Entering dispatch loop..." -Level Info

        # Build tasks ArrayList for the engine
        $dispatchTasks = [System.Collections.ArrayList]::new()
        foreach ($t in $inputTasks) {
            [void]$dispatchTasks.Add($t)
        }

        # --- Shared CE Detail Callback: OnScanCompletions ---
        # (Same signal file scanning pattern as resume — defined here for standalone use)
        $ceDetailOnScan = {
            param($ExportDir, $WorkerDirs, $Context)
            $completed = @()
            $errors = @()
            $completionsDir = Get-CompletionsDir $ExportDir
            $signingKey = Get-ExportRunSigningKey -ExportDir $ExportDir

            $doneSignalFiles = @(Get-ChildItem -Path $completionsDir -Filter "detail-done-*.txt" -File -ErrorAction SilentlyContinue)
            $errorSignalFiles = @(Get-ChildItem -Path $completionsDir -Filter "error-detail-*.txt" -File -ErrorAction SilentlyContinue)

            foreach ($wDir in $WorkerDirs) {
                $doneSignalFiles += @(Get-ChildItem -Path $wDir -Filter "detail-done-*.txt" -File -ErrorAction SilentlyContinue)
                $doneSignalFiles += @(Get-ChildItem -Path $wDir -Filter "done-detail-*.txt" -File -ErrorAction SilentlyContinue)
                $errorSignalFiles += @(Get-ChildItem -Path $wDir -Filter "error-detail-*.txt" -File -ErrorAction SilentlyContinue)
            }

            foreach ($doneFile in $doneSignalFiles) {
                try {
                    $doneContent = [System.IO.File]::ReadAllText($doneFile.FullName)
                    $doneData = ConvertFrom-SignedEnvelopeJson -Json $doneContent -SigningKey $signingKey -RequireSignature:([bool]$signingKey) -Context ("detail completion file {0}" -f $doneFile.Name)
                    if ($null -eq $doneData) {
                        Write-ExportLog -Message ("  Warning: Empty/null detail done file {0}" -f $doneFile.Name) -Level Warning -LogOnly
                        Remove-Item -Path $doneFile.FullName -Force -ErrorAction SilentlyContinue
                        continue
                    }
                    $doneLocation = if ($doneData.Location) { $doneData.Location } else { "" }
                    $completed += @{
                        TagType     = $doneData.TagType
                        TagName     = $doneData.TagName
                        Workload    = $doneData.Workload
                        Location    = $doneLocation
                        RecordCount = $doneData.RecordCount
                        Message     = "{0}/{1}{2} -> {3} records" -f $doneData.TagName, $doneData.Workload, $(if ($doneLocation) { "/$doneLocation" } else { "" }), $doneData.RecordCount
                    }
                    Remove-Item -Path $doneFile.FullName -Force -ErrorAction SilentlyContinue
                }
                catch {
                    Write-ExportLog -Message ("  Warning: Could not parse detail done file {0}" -f $doneFile.Name) -Level Warning -LogOnly
                }
            }

            foreach ($errFile in $errorSignalFiles) {
                try {
                    $errContent = [System.IO.File]::ReadAllText($errFile.FullName)
                    $errData = ConvertFrom-SignedEnvelopeJson -Json $errContent -SigningKey $signingKey -RequireSignature:([bool]$signingKey) -Context ("detail error file {0}" -f $errFile.Name)
                    if ($null -eq $errData) {
                        Write-ExportLog -Message ("  Warning: Empty/null detail error file {0}" -f $errFile.Name) -Level Warning -LogOnly
                        Remove-Item -Path $errFile.FullName -Force -ErrorAction SilentlyContinue
                        continue
                    }
                    $errLocation = if ($errData.Location) { $errData.Location } else { "" }
                    $errors += @{
                        TagType      = $errData.TagType
                        TagName      = $errData.TagName
                        Workload     = $errData.Workload
                        Location     = $errLocation
                        ErrorMessage = $errData.Error
                        Message      = "{0}/{1}{2}: {3}" -f $errData.TagName, $errData.Workload, $(if ($errLocation) { "/$errLocation" } else { "" }), $errData.Error
                    }
                    Remove-Item -Path $errFile.FullName -Force -ErrorAction SilentlyContinue
                }
                catch {
                    Write-ExportLog -Message ("  Warning: Could not parse error file {0}" -f $errFile.Name) -Level Warning -LogOnly
                }
            }

            return @{ CompletedTasks = $completed; ErrorTasks = $errors }
        }

        # --- Shared CE Detail Callback: OnMatchTask ---
        $ceDetailOnMatch = {
            param($Data, $Tasks, $Context)
            $loc = if ($Data.Location) { $Data.Location } else { "" }
            $match = $Tasks | Where-Object {
                $_.TagType -eq $Data.TagType -and
                $_.TagName -eq $Data.TagName -and
                $_.Workload -eq $Data.Workload -and
                $(if ($_.Location) { $_.Location } else { "" }) -eq $loc -and
                $_.Status -eq "InProgress"
            } | Select-Object -First 1
            if (-not $match -and -not $loc) {
                $match = $Tasks | Where-Object {
                    $_.TagType -eq $Data.TagType -and
                    $_.TagName -eq $Data.TagName -and
                    $_.Workload -eq $Data.Workload -and
                    $_.Status -eq "InProgress"
                } | Select-Object -First 1
            }
            return $match
        }

        # --- CE TasksCsv Callback: OnDispatchTask ---
        $csvOnDispatch = {
            param($Worker, $NextTask, $Context)
            $taskData = @{
                Phase         = "Detail"
                TagType       = $NextTask.TagType
                TagName       = $NextTask.TagName
                Workload      = $NextTask.Workload
                Location      = if ($NextTask.Location) { $NextTask.Location } else { "" }
                LocationType  = if ($NextTask.LocationType) { $NextTask.LocationType } else { "" }
                ExpectedCount = ($NextTask.ExpectedCount -as [int])
                PageSize      = ($NextTask.PageSize -as [int])
            }
            return (Send-WorkerTask -WorkerDir $Worker.WorkerDir -TaskData $taskData -ExportDir $Context.ExportDir)
        }

        # --- CE TasksCsv Callback: OnShowDashboard ---
        $csvOnDashboard = {
            param($LoopState, $Context)
            $completedItems = [long]0
            foreach ($dt in $Context.InputTasks) {
                if ($dt.Status -eq "Completed") { $completedItems += ($dt.ExpectedCount -as [long]) }
            }

            $workerStatusList = @()
            foreach ($wp in $LoopState.WorkerProcesses) {
                $wState = Get-WorkerState -WorkerDir $wp.WorkerDir -WorkerPID $wp.PID
                $workerStatusList += @{
                    PID = $wp.PID; State = $wState; CurrentTask = "-"
                    Expected = "-"; Progress = "-"; PageSize = "-"
                    LastPage = "-"; TaskTime = "-"
                }
            }

            try {
                Show-OrchestratorDashboard -Phase "Detail" `
                    -Completed $LoopState.CompletedCount -Total $LoopState.TotalCount `
                    -Workers $workerStatusList -RecentErrors $LoopState.RecentErrors `
                    -RecentActivity $LoopState.RecentActivity -DispatchLog @() `
                    -ExportStartTime $Context.ExportStartTime -PhaseStartTime $Context.PhaseStartTime `
                    -CompletedItems $completedItems -TotalItems $Context.TotalDetailItems `
                    -CompletedBaseline $Context.PreBaselineCompleted `
                    -CompletedItemsBaseline $Context.PreBaselineCompletedItems
            } catch {
                Write-ExportLog -Message ("  Dashboard render error (non-fatal): {0}" -f $_.Exception.Message) -Level Warning -LogOnly
            }
        }

        # --- CE TasksCsv Callback: OnAllWorkersDead ---
        $csvOnAllDead = {
            param($Tasks, $PendingCount, $Context)
            Write-ExportLog -Message ("WARNING: All workers exited with {0} pending tasks" -f $PendingCount) -Level Warning
            Write-TaskCsv -Path $Context.TaskCsvPath -Tasks $Context.InputTasks
        }

        # --- CE TasksCsv Callback: OnIterationComplete ---
        $csvOnIterComplete = {
            param($Tasks, $LoopState, $Context)
            Write-TaskCsv -Path $Context.TaskCsvPath -Tasks $Context.InputTasks
        }

        # Build context hashtable
        $csvContext = @{
            ExportDir              = $exportDir
            TaskCsvPath            = $detailCsvPath
            InputTasks             = $inputTasks
            TotalDetailItems       = $totalDetailItems
            ExportStartTime        = $detailPhaseStartTime
            PhaseStartTime         = $detailPhaseStartTime
            PreBaselineCompleted   = $preBaselineCompleted
            PreBaselineCompletedItems = $preBaselineCompletedItems
        }

        $csvLoopResult = Invoke-DispatchLoop `
            -ExportDir $exportDir `
            -Tasks $dispatchTasks `
            -WorkerProcesses $workerProcesses `
            -Context $csvContext `
            -OnScanCompletions $ceDetailOnScan `
            -OnMatchTask $ceDetailOnMatch `
            -OnDispatchTask $csvOnDispatch `
            -OnShowDashboard $csvOnDashboard `
            -OnAllWorkersDead $csvOnAllDead `
            -OnIterationComplete $csvOnIterComplete `
            -SleepSeconds 2

        # Save final task state
        Write-TaskCsv -Path $detailCsvPath -Tasks $inputTasks

        # Shutdown workers
        if ($workerProcesses.Count -gt 0) {
            Write-ExportLog -Message ("Shutting down {0} worker(s)..." -f $workerProcesses.Count) -Level Info
            Start-Sleep -Seconds 5
            foreach ($wp in $workerProcesses) {
                try {
                    if (-not $wp.Process.HasExited) {
                        Stop-Process -Id $wp.PID -Force -ErrorAction SilentlyContinue
                    }
                }
                catch {
                    Write-Verbose "Could not stop worker PID $($wp.PID): $($_.Exception.Message)"
                }
            }
        }
    }
    else {
        # -- Single-terminal: process detail tasks directly --
        $completedTaskCounts = @{}
        foreach ($task in $script:AllWorkPlanTasks) {
            $taskKey = "{0}|{1}|{2}" -f $task.TagType, $task.TagName, $task.Workload

            if ((-not $task.ExpectedCount -or $task.ExpectedCount -eq 0) -and $task.Status -ne "Error") {
                Write-ExportLog -Message ("    Skipping {0} / {1} - no data" -f $task.TagName, $task.Workload) -Level Info
                continue
            }

            Write-ExportLog -Message ("    Exporting: {0} / {1} (expected {2})" -f $task.TagName, $task.Workload, $task.ExpectedCount) -Level Info

            # Output to Data/ContentExplorer/TagType/TagName/
            $classifierDir = Get-CEClassifierDir $exportDir $task.TagType $task.TagName

            $telemetry = New-ContentExplorerTelemetry -TagType $task.TagType -TagName $task.TagName -Workload $task.Workload
            $exportParams = @{
                Task                  = $task
                PageSize              = $cePageSize
                ProgressLogPath       = $progressLogPath
                Telemetry             = $telemetry
                TelemetryDatabasePath = $telemetryDbPath
                AdaptivePageSize      = $true
                OutputDirectory       = $classifierDir
            }

            try {
                Export-ContentExplorerWithProgress @exportParams | Out-Null
                $exportedCount = if ($task.ExportedCount) { $task.ExportedCount } else { 0 }

                if ($exportedCount -gt 0) {
                    Write-ExportLog -Message ("    Completed: {0} / {1} - {2} records" -f $task.TagName, $task.Workload, $exportedCount) -Level Success
                }

                $completedTaskCounts[$taskKey] = $exportedCount

                # Update task status in CSV
                $csvTask = $inputTasks | Where-Object {
                    $_.TagType -eq $task.TagType -and $_.TagName -eq $task.TagName -and $_.Workload -eq $task.Workload
                } | Select-Object -First 1
                if ($csvTask) {
                    $csvTask.Status = "Completed"
                    if (-not $csvTask.OriginalExpectedCount -or ($csvTask.OriginalExpectedCount -as [int]) -eq 0) {
                        $csvTask.OriginalExpectedCount = $csvTask.ExpectedCount
                    }
                    $csvTask.ExpectedCount = $exportedCount
                    Write-TaskCsv -Path $detailCsvPath -Tasks $inputTasks
                }
            }
            catch {
                Write-ExportLog -Message ("    FAILED: {0} / {1} - {2}" -f $task.TagName, $task.Workload, $_.Exception.Message) -Level Error
                if ($script:ErrorLogPath) {
                    Write-ExportErrorLog -ErrorLogPath $script:ErrorLogPath -Context "TasksCsv Detail Export" -TaskKey $taskKey -ErrorRecord $_
                }
                # Mark as Error in CSV
                $csvTask = $inputTasks | Where-Object {
                    $_.TagType -eq $task.TagType -and $_.TagName -eq $task.TagName -and $_.Workload -eq $task.Workload
                } | Select-Object -First 1
                if ($csvTask) {
                    $csvTask.Status = "Error"
                    $csvTask.ErrorMessage = $_.Exception.Message
                    Write-TaskCsv -Path $detailCsvPath -Tasks $inputTasks
                }
            }

            Save-ContentExplorerRunTracker -Tracker $tracker -TrackerPath $trackerPath
        }
    }

    # Retry bucket detection
    $retryBucketTasks = @()
    if (Test-Path $detailCsvPath) {
        $finalDetailTasks = Read-TaskCsv -Path $detailCsvPath
        $retryBucketTasks = @(Get-RetryBucketTasks -DetailTasks $finalDetailTasks)
        if ($retryBucketTasks.Count -gt 0) {
            $retryTasksCsvPath = Join-Path (Get-CoordinationDir $exportDir) "RetryTasks.csv"
            Write-RetryTasksCsv -Path $retryTasksCsvPath -RetryTasks $retryBucketTasks
            Write-ExportLog -Message ("  Wrote {0} retry tasks to RetryTasks.csv" -f $retryBucketTasks.Count) -Level Info
        }
    }

    # Display retry bucket summary
    Show-RetryBucketSummary -RetryTasks $retryBucketTasks -ExportDir $exportDir

    # Write remaining (non-completed) tasks for follow-on runs
    $remainingCount = Write-RemainingTasksCsv -ExportDir $exportDir
    if ($remainingCount -gt 0) {
        Write-ExportLog -Message ("  Remaining tasks: {0} (see RemainingTasks.csv)" -f $remainingCount) -Level Warning
        Write-ExportLog -Message ("  To re-run: .\Export-Compl8Configuration.ps1 -CETasksCsv ""{0}""" -f (Join-Path $exportDir "RemainingTasks.csv")) -Level Info
    }

    # Write manifest and set final phase
    Write-CEManifest -ExportDir $exportDir
    Write-ExportPhase -ExportDir $exportDir -Phase "Completed"
    Write-ExportLog -Message "Task CSV export complete" -Level Success
}

function Invoke-ContentExplorerExport {
    <#
    .SYNOPSIS
        Orchestrator for Content Explorer exports (single-terminal and multi-terminal).
    .DESCRIPTION
        Manages the full Content Explorer export lifecycle using file-drop coordination:
        - Discovery of tag names from tenant
        - Aggregate phase (parallel via workers or single-terminal)
        - Planning phase (build detail tasks from aggregate results)
        - Detail export phase (parallel via workers or single-terminal)
        Uses ExportPhase.txt, AggregateTasks.csv, DetailTasks.csv, and per-worker
        file-drop nexttask/currenttask files for coordination (no mutex, no ExportProject.json).
    #>
    Write-ExportLog -Message "`n========== Content Explorer Export ==========" -Level Info

    $baseOutputDir = Split-Path $script:ExportRunDirectory -Parent
    $script:UseExistingAggregates = $false
    $script:ExistingAggregatePath = $null
    $script:SharedExportDirectory = $script:ExportRunDirectory

    # ... [unchanged: aggregate caching check] ...
    # Check for recent aggregate CSVs to reuse (no more "join existing export" -- that concept
    # is removed because there is no ExportProject.json to join).
    # Instead, we only offer to reuse aggregate data from previous exports.

    # Get current tenant info for filtering and logging
    $currentTenant = Get-Compl8TenantInfo
    if ($currentTenant) {
        Write-ExportLog -Message ("Current tenant: {0} ({1})" -f $currentTenant.TenantDomain, $currentTenant.TenantId) -Level Info
    }

    # Find aggregates matching current tenant (or all if no tenant filter)
    $tenantFilter = if ($currentTenant) { $currentTenant.TenantId } else { $null }
    $recentAggregates = Find-RecentAggregateCsv -OutputDirectory $baseOutputDir -MaxAgeDays 30 -TenantId $tenantFilter

    if ($recentAggregates.Count -gt 0) {
        Write-ExportLog -Message "`n--- Recent Aggregate Data Found (matching tenant) ---" -Level Info
        Write-ExportLog -Message ("Found {0} aggregate file(s) from the last 30 days:" -f $recentAggregates.Count) -Level Info

        $displayCount = [Math]::Min($recentAggregates.Count, 5)
        for ($i = 0; $i -lt $displayCount; $i++) {
            $agg = $recentAggregates[$i]
            $ageStr = if ($agg.AgeHours -lt 24) { "$($agg.AgeHours) hours ago" } else { "$($agg.AgeDays) days ago" }
            $tenantStr = if ($agg.TenantDomain) { " [$($agg.TenantDomain)]" } else { "" }
            Write-ExportLog -Message ("  [{0}] {1}: {2} records ({3}){4}" -f ($i + 1), $agg.FolderName, $agg.RecordCount.ToString('N0'), $ageStr, $tenantStr) -Level Info
        }

        Write-Host ""
        Write-Host "Would you like to reuse existing aggregate data? (Saves time on large tenants)" -ForegroundColor Cyan
        Write-Host ("  [1-{0}] Use the aggregate file shown above" -f $displayCount)
        Write-Host "  [N] Generate fresh aggregate data (slower but current)"
        Write-Host ""
        $choice = Read-Host "Enter choice [N]"

        if ($choice -match '^[1-5]$') {
            $choiceIndex = [int]$choice - 1
            if ($choiceIndex -lt $recentAggregates.Count) {
                $selectedAggregate = $recentAggregates[$choiceIndex]
                $script:UseExistingAggregates = $true
                $script:ExistingAggregatePath = $selectedAggregate.Path

                # Copy the aggregate file to current export directory for reference
                Copy-Item -Path $selectedAggregate.Path -Destination (Join-Path (Get-CoordinationDir $script:ExportRunDirectory) "ContentExplorer-Aggregates.csv") -Force
                $sourceMetadata = Join-Path (Split-Path $selectedAggregate.Path) "AggregateMetadata.json"
                if (Test-Path $sourceMetadata) {
                    Copy-Item -Path $sourceMetadata -Destination (Join-Path (Get-CoordinationDir $script:ExportRunDirectory) "AggregateMetadata.json") -Force
                }
                Write-ExportLog -Message ("Using existing aggregate data from: {0}" -f $selectedAggregate.FolderName) -Level Success
            }
        }
        else {
            Write-ExportLog -Message "Generating fresh aggregate data..." -Level Info
        }
    }

    # Save tenant metadata for this export (for future aggregate reuse)
    if ($currentTenant -and -not $script:UseExistingAggregates) {
        Save-AggregateMetadata -ExportRunDirectory $script:ExportRunDirectory -TenantInfo $currentTenant
    }

    # ... [unchanged: load classifier configuration] ...
    # Load classifier configuration
    $configPath = Join-Path $scriptRoot "ConfigFiles\ContentExplorerClassifiers.json"
    $config = Read-JsonConfig -Path $configPath

    # Default settings
    $batchSize = $script:CEDefaultBatchSize
    $workloads = @($script:CEDefaultWorkloads)
    $cePageSize = 1000  # Default page size for Content Explorer

    if ($config -and $config.Settings) {
        if ($config.Settings.BatchSize) { $batchSize = $config.Settings.BatchSize }
        if ($config.Settings.Workloads) { $workloads = $config.Settings.Workloads }
        if ($config.Settings.PageSize) { $cePageSize = $config.Settings.PageSize }
    }

    # Apply menu/CLI workload selection (overrides config)
    if ($CEWorkloads) {
        $workloads = @($CEWorkloads)
    }

    # Save export settings manifest for resume consistency
    Save-ExportSettings -ExportRunDirectory $script:ExportRunDirectory -ExportType "ContentExplorer" -Settings @{
        Workloads = $workloads
        CEAllSITs = [bool]$CEAllSITs
        BatchSize = $batchSize
        PageSize  = $cePageSize
    }

    Write-ExportLog -Message ("Default page size: {0} (adaptive sizing selects optimal size per workload)" -f $cePageSize) -Level Info
    Write-ExportLog -Message ("Workloads: {0}" -f ($workloads -join ', ')) -Level Info
    if ($CEAllSITs) {
        Write-ExportLog -Message "MODE: All SITs - Auto-discovering all Sensitive Information Types" -Level Info
    }

    # Create progress log file for tailing
    $progressLogPath = Join-Path (Get-LogsDir $script:ExportRunDirectory) "ContentExplorer-Progress.log"
    Write-ExportLog -Message ("Progress log (tail -f): {0}" -f $progressLogPath) -Level Info

    # Telemetry database path for adaptive paging analysis
    $telemetryDbPath = Join-Path $scriptRoot "TelemetryDB\content-explorer-telemetry.jsonl"
    Write-ExportLog -Message ("Telemetry database: {0}" -f $telemetryDbPath) -Level Info

    # Aggregate results CSV for planning and progress tracking
    $aggregateCsvPath = Join-Path (Get-CoordinationDir $script:ExportRunDirectory) "ContentExplorer-Aggregates.csv"
    if (-not (Test-Path $aggregateCsvPath)) {
        "Timestamp,TagType,TagName,Workload,Location,Count,Error" | Set-Content -Path $aggregateCsvPath -Encoding UTF8
    }
    Write-ExportLog -Message ("Aggregate results CSV: {0}" -f $aggregateCsvPath) -Level Info

    # Initialize or load run tracker
    $trackerPath = Join-Path (Get-CoordinationDir $script:ExportRunDirectory) "ContentExplorer-RunTracker.json"
    $tracker = Get-ContentExplorerRunTracker -TrackerPath $trackerPath

    # Track exported counts per task for progress display
    $completedTaskCounts = @{}

    # ... [unchanged: SIT GUID mapping] ...
    # Build SIT GUID mapping for resolving GUIDs in results
    if (-not $tracker.SitMapping -or $tracker.SitMapping.Count -eq 0) {
        $tracker.SitMapping = Get-SitGuidMapping
        Save-ContentExplorerRunTracker -Tracker $tracker -TrackerPath $trackerPath
    }
    else {
        Write-ExportLog -Message ("Using cached SIT mapping ({0} SITs)" -f $tracker.SitMapping.Count) -Level Info
    }

    # Load SITs to skip
    $sitsToSkipPath = Join-Path $scriptRoot "ConfigFiles\SITstoSkip.json"
    $sitsToSkip = Get-SITsToSkip -ConfigPath $sitsToSkipPath

    $allRecords = @()
    $isMultiTerminal = ($WorkerCount -and $WorkerCount -ge 2)
    $exportStartTime = Get-Date

    #region -- Initialization: Write ExportPhase.txt and ExportType.txt --
    # Replace Initialize-ExportProject / Find-ExportProject with simple phase file
    Write-ExportPhase -ExportDir $script:SharedExportDirectory -Phase "Aggregate"
    Write-ExportType -ExportDir $script:SharedExportDirectory -Type "ContentExplorer"
    Write-ExportLog -Message ("Export phase initialized: Aggregate (dir: {0})" -f $script:SharedExportDirectory) -Level Info
    #endregion

    #region -- Phase 1: Discovery --
    # Discover all tag names BEFORE spawning workers or running aggregates.
    # This is unchanged -- we discover tag names from tenant, config, or cached aggregates.

    $tagTypeConfigs = @(
        @{ TagType = "Sensitivity"; ConfigSection = "SensitivityLabels"; DiscoverCmd = { Get-Label -ErrorAction Stop } }
        @{ TagType = "Retention"; ConfigSection = "RetentionLabels"; DiscoverCmd = { Get-ComplianceTag -ErrorAction Stop } }
        @{ TagType = "SensitiveInformationType"; ConfigSection = "SensitiveInformationTypes"; DiscoverCmd = { Get-DlpSensitiveInformationType -ErrorAction Stop } }
        @{ TagType = "TrainableClassifier"; ConfigSection = "TrainableClassifiers"; DiscoverCmd = $null }
    )

    # Collect discovered tag names per tag type for later aggregate/detail phases
    $discoveredTagsByType = @{}

    foreach ($ttConfig in $tagTypeConfigs) {
        $tagType = $ttConfig.TagType
        $configSection = $ttConfig.ConfigSection
        $tagNames = @()

        # CEAllSITs mode: only process SensitiveInformationType
        if ($CEAllSITs -and $tagType -ne "SensitiveInformationType") {
            Write-ExportLog -Message ("`n--- {0} --- (skipped in All SITs mode)" -f $tagType) -Level Info
            continue
        }

        Write-ExportLog -Message ("`n--- {0} ---" -f $tagType) -Level Info

        # ... [unchanged: config section parsing and auto-discover logic] ...
        # Check if this tag type is configured
        $sectionConfig = $null
        if ($config -and $config.$configSection) {
            $sectionConfig = $config.$configSection
        }

        # Determine tag names - either from config or auto-discover
        $autoDiscover = $true
        if ($CEAllSITs -and $tagType -eq "SensitiveInformationType") {
            $autoDiscover = $true
            Write-ExportLog -Message "  All SITs mode: forcing auto-discovery" -Level Info
        }
        elseif ($sectionConfig -and $sectionConfig._AutoDiscover -eq "False") {
            $autoDiscover = $false
            $tagNames = @($sectionConfig.PSObject.Properties |
                Where-Object { $_.Name -notlike "_*" -and $_.Value -eq "True" } |
                ForEach-Object { $_.Name })
            Write-ExportLog -Message ("  Using {0} classifiers from config" -f $tagNames.Count) -Level Info
        }

        if ($autoDiscover) {
            # Check if we can use cached aggregate data to skip tenant discovery
            if ($script:UseExistingAggregates -and $script:ExistingAggregatePath) {
                $cachedTagNames = Get-TagNamesFromAggregateCsv -CsvPath $script:ExistingAggregatePath -TagType $tagType
                if ($cachedTagNames.Count -gt 0) {
                    $tagNames = $cachedTagNames
                    Write-ExportLog -Message ("  Using {0} classifiers from cached aggregates (skipped tenant discovery)" -f $tagNames.Count) -Level Success
                }
                else {
                    Write-ExportLog -Message ("  No cached data for {0} - falling back to tenant discovery" -f $tagType) -Level Info
                }
            }

            # Normal tenant discovery (if not using cached data or fallback needed)
            if ($tagNames.Count -eq 0 -and $ttConfig.DiscoverCmd) {
                Write-ExportLog -Message "  Auto-discovering classifiers from tenant..." -Level Info
                try {
                    $discovered = & $ttConfig.DiscoverCmd
                    if ($tagType -eq "Sensitivity") {
                        foreach ($label in $discovered) {
                            if ($label.ParentLabelDisplayName) {
                                $tagNames += "{0}/{1}" -f $label.ParentLabelDisplayName, $label.DisplayName
                            }
                            else {
                                $tagNames += $label.DisplayName
                            }
                        }
                    }
                    else {
                        $tagNames = @($discovered.Name)
                    }
                    Write-ExportLog -Message ("  Discovered {0} classifiers" -f $tagNames.Count) -Level Info
                }
                catch {
                    Write-ExportLog -Message ("  Failed to discover: {0}" -f $_.Exception.Message) -Level Error
                    continue
                }
            }
            elseif ($tagNames.Count -eq 0 -and -not $ttConfig.DiscoverCmd) {
                Write-ExportLog -Message "  No classifiers configured and auto-discover not available" -Level Warning
                continue
            }
        }

        # Filter out SITs to skip
        if ($tagType -eq "SensitiveInformationType" -and $sitsToSkip.Count -gt 0) {
            $originalCount = $tagNames.Count
            $tagNames = @($tagNames | Where-Object { $_ -notin $sitsToSkip })
            $skippedCount = $originalCount - $tagNames.Count
            if ($skippedCount -gt 0) {
                Write-ExportLog -Message ("  Filtered out {0} SITs from skip list ({1} remaining)" -f $skippedCount, $tagNames.Count) -Level Info
            }
        }

        # Filter out empty/null tag names (can occur if tenant returns labels with blank names)
        $tagNames = @($tagNames | Where-Object { -not [string]::IsNullOrEmpty($_) })

        if ($tagNames.Count -eq 0) {
            Write-ExportLog -Message "  No classifiers to process" -Level Warning
            continue
        }

        # Store discovered tag names for this type
        $discoveredTagsByType[$tagType] = $tagNames
    }

    #endregion

    #region -- Phase 2: Build and Write Aggregate Task CSV --
    # Replace Update-ProjectAggregateTasks with Write-TaskCsv for AggregateTasks.csv

    $aggregateTaskList = @()
    foreach ($tagType in $discoveredTagsByType.Keys) {
        $tagNames = $discoveredTagsByType[$tagType]
        foreach ($tagName in $tagNames) {
            foreach ($workload in $workloads) {
                $aggregateTaskList += @{
                    TagType      = $tagType
                    TagName      = $tagName
                    Workload     = $workload
                    ExpectedCount = 0
                    PageSize     = 5000
                    AssignedPID  = 0
                    Status       = "Pending"
                    ErrorMessage = ""
                }
            }
        }
    }

    $aggTaskCsvPath = Join-Path (Get-CoordinationDir $script:SharedExportDirectory) "AggregateTasks.csv"
    if ($aggregateTaskList.Count -gt 0 -and -not $script:UseExistingAggregates) {
        Write-TaskCsv -Path $aggTaskCsvPath -Tasks $aggregateTaskList
        Write-ExportLog -Message ("Wrote {0} aggregate tasks to AggregateTasks.csv" -f $aggregateTaskList.Count) -Level Info
    }
    elseif ($script:UseExistingAggregates) {
        Write-ExportLog -Message "Skipping aggregate task CSV (using existing aggregates)" -Level Info
    }

    #endregion

    #region -- Phase 3: Spawn Workers (after discovery) --
    # Workers are spawned AFTER all tag types are discovered and aggregate tasks
    # are written to CSV. Workers receive ExportRunDirectory, not ProjectPath.

    $workerProcesses = [System.Collections.ArrayList]::new()
    if ($isMultiTerminal) {
        Write-ExportLog -Message ("Multi-terminal mode: spawning {0} worker(s)" -f $WorkerCount) -Level Info
        $workerProcesses = Start-WorkerTerminals -ExportRunDirectory $script:SharedExportDirectory -Count $WorkerCount
    }

    #endregion

    #region -- Phase 4: Worker folder setup --
    # In multi-terminal mode, the orchestrator does NOT create a worker folder --
    # it acts as coordinator only (no task processing).

    $workerDir = $null
    if (-not $isMultiTerminal) {
        # Single-terminal mode: no worker subfolder needed (output goes to root)
    }

    #endregion

    #region -- Phase 5: Aggregate Phase --

    $hasAggregateErrors = $false
    $aggregateErrorTasks = @()

    if ($isMultiTerminal -and -not $script:UseExistingAggregates) {
        # -- Multi-terminal: unified continuous pipeline via Invoke-DispatchLoop --
        # Replaces separate aggregate loop, planning phase, and detail loop with one
        # continuous pipeline. Detail tasks are generated incrementally as each
        # aggregate completes (no Planning pause between phases).
        Write-ExportLog -Message "Orchestrator starting unified dispatch pipeline..." -Level Info
        Write-ExportPhase -ExportDir $script:SharedExportDirectory -Phase "Aggregate"

        $aggTasks = Read-TaskCsv -Path $aggTaskCsvPath
        $detailTaskCsvPath = Join-Path (Get-CoordinationDir $script:SharedExportDirectory) "DetailTasks.csv"
        $completionsDir = Get-CompletionsDir $script:SharedExportDirectory
        Reset-OrchestratorDashboard
        $pipelineStartTime = Get-Date
        $lastSessionKeepalive = Get-Date
        $keepaliveInterval = New-TimeSpan -Minutes 10
        $nextWorkerNumber = $workerProcesses.Count + 1

        # Build unified task ArrayList starting with aggregate tasks
        $unifiedTasks = [System.Collections.ArrayList]::new()
        foreach ($at in $aggTasks) {
            [void]$unifiedTasks.Add(@{
                Phase         = "Aggregate"
                TagType       = $at.TagType
                TagName       = $at.TagName
                Workload      = $at.Workload
                ExpectedCount = ($at.ExpectedCount -as [int])
                PageSize      = ($at.PageSize -as [int])
                Status        = $at.Status
                AssignedPID   = ($at.AssignedPID -as [int])
                ErrorMessage  = if ($at.ErrorMessage) { $at.ErrorMessage } else { "" }
            })
        }

        # --- CE Callback: OnScanCompletions ---
        # Scans worker dirs for aggregate AND detail completion/error signals.
        $ceOnScan = {
            param($ExportDir, $WorkerDirs, $Context)
            $completed = @()
            $errors = @()
            $aggCsvPath = $Context.AggregateCsvPath
            $signingKey = Get-ExportRunSigningKey -ExportDir $ExportDir

            # --- Aggregate completions: agg-*.csv files (now in Data/ContentExplorer/Aggregates/) ---
            $aggDataDir = Join-Path (Get-CEDataDir $ExportDir) "Aggregates"
            if (Test-Path $aggDataDir) {
                $aggOutputFiles = @(Get-ChildItem -Path $aggDataDir -Filter "agg-*.csv" -File -ErrorAction SilentlyContinue)
                foreach ($aggFile in $aggOutputFiles) {
                    try {
                        $aggFileContent = @(Import-Csv -Path $aggFile.FullName -Encoding UTF8 -ErrorAction Stop)
                        if ($aggFileContent.Count -gt 0) {
                            $fileTagType = $aggFileContent[0].TagType
                            $fileTagName = $aggFileContent[0].TagName
                            $fileWorkload = $aggFileContent[0].Workload

                            # Check for error CSV (Location=ERROR)
                            $errorRow = $aggFileContent | Where-Object { $_.Location -eq 'ERROR' } | Select-Object -First 1
                            if ($errorRow) {
                                $errMsg = if ($errorRow.Error) { $errorRow.Error } else { "Aggregate failed (error CSV from worker)" }
                                # Write error to central aggregate CSV
                                $errorCsvLine = '{0},{1},{2},{3},ERROR,0,"{4}"' -f $errorRow.Timestamp, $fileTagType, $fileTagName, $fileWorkload, ($errMsg -replace '"', '""')
                                Add-Content -Path $aggCsvPath -Value $errorCsvLine -Encoding UTF8

                                $errors += @{
                                    TaskType     = "Aggregate"
                                    TagType      = $fileTagType
                                    TagName      = $fileTagName
                                    Workload     = $fileWorkload
                                    ErrorMessage = $errMsg
                                    Message      = "{0}/{1}: {2}" -f $fileTagName, $fileWorkload, $errMsg
                                }
                            }
                            else {
                                # Normal success: compute file count
                                $fileCountRow = $aggFileContent | Where-Object { $_.Location -eq '_FILECOUNT' } | Select-Object -First 1
                                if ($fileCountRow -and ($fileCountRow.Count -as [int]) -gt 0) {
                                    $totalCount = $fileCountRow.Count -as [int]
                                } else {
                                    $totalCount = ($aggFileContent | Where-Object { $_.Location -notin @('NONE', '_FILECOUNT') } | Measure-Object -Property Count -Sum -ErrorAction SilentlyContinue).Sum
                                    if (-not $totalCount) { $totalCount = 0 }
                                }

                                # Copy rows to central aggregate CSV (batched)
                                $csvBatch = [System.Text.StringBuilder]::new()
                                foreach ($row in $aggFileContent) {
                                    if ($row.Location -eq 'NONE') { continue }
                                    $locationName = $row.Location -replace '"', '""'
                                    if ($locationName -match '[,"]') { $locationName = '"{0}"' -f $locationName }
                                    $csvLine = "{0},{1},{2},{3},{4},{5}" -f $row.Timestamp, $row.TagType, $row.TagName, $row.Workload, $locationName, $row.Count
                                    [void]$csvBatch.AppendLine($csvLine)
                                }
                                if ($csvBatch.Length -gt 0) {
                                    Add-Content -Path $aggCsvPath -Value $csvBatch.ToString().TrimEnd() -Encoding UTF8
                                }

                                $completed += @{
                                    TaskType     = "Aggregate"
                                    TagType      = $fileTagType
                                    TagName      = $fileTagName
                                    Workload     = $fileWorkload
                                    ExpectedCount = $totalCount
                                    Message      = "{0}/{1} -> {2} files" -f $fileTagName, $fileWorkload, $totalCount
                                }
                            }
                        }
                        Rename-Item -Path $aggFile.FullName -NewName ($aggFile.Name -replace '\.csv$', '.done') -Force -ErrorAction SilentlyContinue
                    }
                    catch {
                        Write-ExportLog -Message ("  Warning: Could not read aggregate output {0}: {1}" -f $aggFile.Name, $_.Exception.Message) -Level Warning -LogOnly
                    }
                }

            }

            # --- Aggregate errors: error-agg-*.txt files (in worker coordination dirs) ---
            foreach ($wDir in $WorkerDirs) {
                $aggErrorFiles = @(Get-ChildItem -Path $wDir -Filter "error-agg-*.txt" -File -ErrorAction SilentlyContinue)
                foreach ($errFile in $aggErrorFiles) {
                    try {
                        $errContent = [System.IO.File]::ReadAllText($errFile.FullName)
                        $errData = ConvertFrom-SignedEnvelopeJson -Json $errContent -SigningKey $signingKey -RequireSignature:([bool]$signingKey) -Context ("aggregate error file {0}" -f $errFile.Name)
                        if ($null -ne $errData) {
                            $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                            $errorCsvLine = '{0},{1},{2},{3},ERROR,0,"{4}"' -f $timestamp, $errData.TagType, $errData.TagName, $errData.Workload, ($errData.Error -replace '"', '""')
                            Add-Content -Path $aggCsvPath -Value $errorCsvLine -Encoding UTF8

                            $errors += @{
                                TaskType     = "Aggregate"
                                TagType      = $errData.TagType
                                TagName      = $errData.TagName
                                Workload     = $errData.Workload
                                ErrorMessage = $errData.Error
                                Message      = "{0}/{1}: {2}" -f $errData.TagName, $errData.Workload, $errData.Error
                            }
                        }
                        Rename-Item -Path $errFile.FullName -NewName ($errFile.Name -replace '\.txt$', '.done') -Force -ErrorAction SilentlyContinue
                    }
                    catch {
                        Write-ExportLog -Message ("  Warning: Could not parse aggregate error file {0}" -f $errFile.Name) -Level Warning -LogOnly
                    }
                }
            }

            # --- Detail completions: detail-done-*.txt in Completions/ dir ---
            $cDir = Get-CompletionsDir $ExportDir
            $doneSignalFiles = @(Get-ChildItem -Path $cDir -Filter "detail-done-*.txt" -File -ErrorAction SilentlyContinue)
            foreach ($doneFile in $doneSignalFiles) {
                try {
                    $doneData = ConvertFrom-SignedEnvelopeJson -Json ([System.IO.File]::ReadAllText($doneFile.FullName)) -SigningKey $signingKey -RequireSignature:([bool]$signingKey) -Context ("detail completion file {0}" -f $doneFile.Name)
                    if ($null -ne $doneData) {
                        $doneLocation = if ($doneData.Location) { $doneData.Location } else { "" }
                        $completed += @{
                            TaskType    = "Detail"
                            TagType     = $doneData.TagType
                            TagName     = $doneData.TagName
                            Workload    = $doneData.Workload
                            Location    = $doneLocation
                            RecordCount = $doneData.RecordCount
                            Message     = "{0}/{1}{2} -> {3} records" -f $doneData.TagName, $doneData.Workload, $(if ($doneLocation) { "/$doneLocation" } else { "" }), $doneData.RecordCount
                        }
                    }
                    Remove-Item -Path $doneFile.FullName -Force -ErrorAction SilentlyContinue
                }
                catch {
                    Write-ExportLog -Message ("  Warning: Could not parse detail done file {0}" -f $doneFile.Name) -Level Warning -LogOnly
                }
            }

            # --- Detail errors: error-detail-*.txt in Completions/ dir ---
            $errorSignalFiles = @(Get-ChildItem -Path $cDir -Filter "error-detail-*.txt" -File -ErrorAction SilentlyContinue)
            foreach ($errFile in $errorSignalFiles) {
                try {
                    $errContent = [System.IO.File]::ReadAllText($errFile.FullName)
                    $errData = ConvertFrom-SignedEnvelopeJson -Json $errContent -SigningKey $signingKey -RequireSignature:([bool]$signingKey) -Context ("detail error file {0}" -f $errFile.Name)
                    if ($null -ne $errData) {
                        $errLocation = if ($errData.Location) { $errData.Location } else { "" }
                        $errors += @{
                            TaskType     = "Detail"
                            TagType      = $errData.TagType
                            TagName      = $errData.TagName
                            Workload     = $errData.Workload
                            Location     = $errLocation
                            ErrorMessage = $errData.Error
                            Message      = "{0}/{1}{2}: {3}" -f $errData.TagName, $errData.Workload, $(if ($errLocation) { "/$errLocation" } else { "" }), $errData.Error
                        }
                    }
                    Remove-Item -Path $errFile.FullName -Force -ErrorAction SilentlyContinue
                }
                catch {
                    Write-ExportLog -Message ("  Warning: Could not parse detail error file {0}" -f $errFile.Name) -Level Warning -LogOnly
                }
            }

            return @{ CompletedTasks = $completed; ErrorTasks = $errors }
        }

        # --- CE Callback: OnMatchTask ---
        # Matches completion/error data to the correct task in the unified ArrayList.
        $ceOnMatch = {
            param($Data, $Tasks, $Context)
            if ($Data.TaskType -eq "Aggregate") {
                $Tasks | Where-Object {
                    $_.Phase -eq "Aggregate" -and
                    $_.TagType -eq $Data.TagType -and
                    $_.TagName -eq $Data.TagName -and
                    $_.Workload -eq $Data.Workload -and
                    $_.Status -eq "InProgress"
                } | Select-Object -First 1
            }
            elseif ($Data.TaskType -eq "Detail") {
                $loc = if ($Data.Location) { $Data.Location } else { "" }
                # Try location-specific match first
                $match = $Tasks | Where-Object {
                    $_.Phase -eq "Detail" -and
                    $_.TagType -eq $Data.TagType -and
                    $_.TagName -eq $Data.TagName -and
                    $_.Workload -eq $Data.Workload -and
                    $(if ($_.Location) { $_.Location } else { "" }) -eq $loc -and
                    $_.Status -eq "InProgress"
                } | Select-Object -First 1
                if (-not $match -and -not $loc) {
                    # Fallback: match by type/name/workload only
                    $match = $Tasks | Where-Object {
                        $_.Phase -eq "Detail" -and
                        $_.TagType -eq $Data.TagType -and
                        $_.TagName -eq $Data.TagName -and
                        $_.Workload -eq $Data.Workload -and
                        $_.Status -eq "InProgress"
                    } | Select-Object -First 1
                }
                return $match
            }
        }

        # --- CE Callback: OnCompletionGeneratesTasks ---
        # Fires when an aggregate completes. Generates detail tasks from the
        # aggregate data (one per location, or a WorkloadFallback if no locations).
        $ceOnGenerate = {
            param($CompletedTask, $CompletionData, $Context)
            # Only generate detail tasks for aggregate completions
            if ($CompletedTask.Phase -ne "Aggregate") { return @() }

            $newTasks = @()
            $aggCsvPath = $Context.AggregateCsvPath
            $cePageSize = $Context.DefaultPageSize

            # Determine LocationType based on workload
            $locationType = switch ($CompletedTask.Workload) {
                'Exchange'   { 'UPN' }
                'Teams'      { 'UPN' }
                'SharePoint' { 'SiteUrl' }
                'OneDrive'   { 'SiteUrl' }
                default      { 'SiteUrl' }
            }

            $totalCount = $CompletedTask.ExpectedCount -as [int]

            # Skip zero-count aggregates (no data to export)
            if (-not $totalCount -or $totalCount -le 0) {
                Write-ExportLog -Message ("  Aggregate {0}/{1} completed with 0 files - skipping detail tasks" -f $CompletedTask.TagName, $CompletedTask.Workload) -Level Info -LogOnly
                return @()
            }

            try {
                $importResult = Import-AggregateDataFromCsv -CsvPath $aggCsvPath `
                    -TagType $CompletedTask.TagType -TagNames @($CompletedTask.TagName) -Workloads @($CompletedTask.Workload)

                foreach ($taskKey in $importResult.TaskData.Keys) {
                    $taskData = $importResult.TaskData[$taskKey]

                    if ((-not $taskData.TotalCount) -or ($taskData.TotalCount -as [int]) -eq 0) { continue }

                    if ($taskData.Locations -and @($taskData.Locations).Count -gt 0) {
                        # --- Location-level detail tasks ---
                        foreach ($loc in @($taskData.Locations)) {
                            $locExpected = $loc.ExpectedCount -as [int]
                            if ($locExpected -le 0) { continue }

                            # Page size tiers for location-level tasks
                            $locPageSize = if ($locExpected -ge 10000) { 5000 }
                                           elseif ($locExpected -ge 4000) { 2000 }
                                           elseif ($locExpected -ge 500) { 1000 }
                                           else { 500 }

                            $newTasks += @{
                                Phase                 = "Detail"
                                TagType               = $taskData.TagType
                                TagName               = $taskData.TagName
                                Workload              = $taskData.Workload
                                Location              = $loc.Name
                                LocationType          = $locationType
                                ExpectedCount         = $locExpected
                                OriginalExpectedCount = $locExpected
                                PageSize              = $locPageSize
                                Status                = "Pending"
                                AssignedPID           = 0
                                ErrorMessage          = ""
                            }
                        }
                    }
                    else {
                        # --- No location data: WorkloadFallback task ---
                        $taskPageSize = $cePageSize
                        $wpExpected = $taskData.TotalCount -as [int]
                        $psFloor = 100
                        if ($wpExpected -gt 0) {
                            $maxPs = [Math]::Max($psFloor, 2 * $wpExpected)
                            $taskPageSize = [Math]::Max($psFloor, [Math]::Min($taskPageSize, $maxPs))
                        } else {
                            $taskPageSize = [Math]::Max($psFloor, $taskPageSize)
                        }

                        $newTasks += @{
                            Phase                 = "Detail"
                            TagType               = $taskData.TagType
                            TagName               = $taskData.TagName
                            Workload              = $taskData.Workload
                            Location              = ""
                            LocationType          = "WorkloadFallback"
                            ExpectedCount         = $taskData.TotalCount
                            OriginalExpectedCount = $taskData.TotalCount
                            PageSize              = $taskPageSize
                            Status                = "Pending"
                            AssignedPID           = 0
                            ErrorMessage          = ""
                        }
                    }
                }
            }
            catch {
                Write-ExportLog -Message ("  Failed to generate detail tasks from aggregate {0}/{1}: {2}" -f $CompletedTask.TagName, $CompletedTask.Workload, $_.Exception.Message) -Level Warning -LogOnly
                # Fallback: single WorkloadFallback task
                $taskPageSize = [Math]::Max(500, $cePageSize)
                $newTasks += @{
                    Phase                 = "Detail"
                    TagType               = $CompletedTask.TagType
                    TagName               = $CompletedTask.TagName
                    Workload              = $CompletedTask.Workload
                    Location              = ""
                    LocationType          = "WorkloadFallback"
                    ExpectedCount         = $totalCount
                    OriginalExpectedCount = $totalCount
                    PageSize              = $taskPageSize
                    Status                = "Pending"
                    AssignedPID           = 0
                    ErrorMessage          = "Detail planning failed"
                }
            }

            # Sort new detail tasks largest-first for optimal scheduling
            if ($newTasks.Count -gt 1) {
                $newTasks = @($newTasks | Sort-Object { [int]$_.ExpectedCount } -Descending)
            }

            if ($newTasks.Count -gt 0) {
                Write-ExportLog -Message ("  Generated {0} detail tasks from aggregate {1}/{2}" -f $newTasks.Count, $CompletedTask.TagName, $CompletedTask.Workload) -Level Info -LogOnly
            }

            return $newTasks
        }

        # --- CE Callback: OnCompletionGeneratesTasks for ERROR aggregates ---
        # The engine only calls OnCompletionGeneratesTasks for completed tasks, not errored ones.
        # We handle error-based detail generation in OnIterationComplete by checking for newly
        # errored aggregate tasks and injecting WorkloadFallback detail tasks.
        # (This is tracked via $Context.ProcessedAggErrorKeys)

        # --- CE Callback: OnDispatchTask ---
        $ceOnDispatch = {
            param($Worker, $NextTask, $Context)
            if ($NextTask.Phase -eq "Aggregate") {
                $taskData = @{
                    Phase    = "Aggregate"
                    TagType  = $NextTask.TagType
                    TagName  = $NextTask.TagName
                    Workload = $NextTask.Workload
                    PageSize = ($NextTask.PageSize -as [int])
                }
            }
            else {
                $taskData = @{
                    Phase         = "Detail"
                    TagType       = $NextTask.TagType
                    TagName       = $NextTask.TagName
                    Workload      = $NextTask.Workload
                    Location      = if ($NextTask.Location) { $NextTask.Location } else { "" }
                    LocationType  = if ($NextTask.LocationType) { $NextTask.LocationType } else { "" }
                    ExpectedCount = ($NextTask.ExpectedCount -as [int])
                    PageSize      = ($NextTask.PageSize -as [int])
                }
            }
            return (Send-WorkerTask -WorkerDir $Worker.WorkerDir -TaskData $taskData -ExportDir $Context.ExportDir)
        }

        # --- CE Callback: OnShowDashboard ---
        $ceOnDashboard = {
            param($LoopState, $Context)

            $tasks = $Context.UnifiedTasks
            $aggTasks = @($tasks | Where-Object { $_.Phase -eq "Aggregate" })
            $detTasks = @($tasks | Where-Object { $_.Phase -eq "Detail" })

            $aggDone = @($aggTasks | Where-Object { $_.Status -in @("Completed", "Error") }).Count
            $aggTotal = $aggTasks.Count
            $detDone = @($detTasks | Where-Object { $_.Status -in @("Completed", "Error") }).Count
            $detTotal = $detTasks.Count
            $detErrors = @($detTasks | Where-Object { $_.Status -eq "Error" }).Count
            $detActive = @($detTasks | Where-Object { $_.Status -eq "InProgress" }).Count

            # Determine current phase for display
            $displayPhase = if ($aggDone -lt $aggTotal) { "Aggregate" } else { "Detail" }

            # Build total progress (weighted: aggregate tasks are discovery, detail tasks are the real work)
            $displayCompleted = $aggDone + $detDone
            $displayTotal = $aggTotal + $detTotal

            # Build worker status
            $workerStatusList = @()
            $completedDetailItems = [long]0
            $totalDetailItems = [long]0
            $inProgressItemTotal = 0

            # Compute detail totals for ETA
            foreach ($dt in $detTasks) {
                $dtExp = ($dt.ExpectedCount -as [long])
                if ($dtExp -gt 0) { $totalDetailItems += $dtExp }
                if ($dt.Status -eq "Completed") {
                    $dtCount = ($dt.ExpectedCount -as [long])
                    if ($dtCount -gt 0) { $completedDetailItems += $dtCount }
                }
            }

            # Build PID index for worker task display
            $tasksByPID = @{}
            foreach ($t in $tasks) {
                if ($t.Status -eq "InProgress") {
                    $taskPid = $t.AssignedPID -as [int]
                    if ($taskPid -gt 0) {
                        if (-not $tasksByPID.ContainsKey($taskPid)) { $tasksByPID[$taskPid] = @() }
                        $tasksByPID[$taskPid] += $t
                    }
                }
            }

            foreach ($wp in $LoopState.WorkerProcesses) {
                $wDir = $wp.WorkerDir
                $wPID = $wp.PID
                $wState = Get-WorkerState -WorkerDir $wDir -WorkerPID $wPID
                $currentTaskName = "-"
                $taskTimeStr = "-"
                $expectedStr = "-"
                $progressStr = "-"
                $pageSizeStr = "-"
                $lastPageTimeStr = "-"
                try {
                    $ctPath = Join-Path $wDir "currenttask"
                    $ctData = [System.IO.File]::ReadAllText($ctPath) | ConvertFrom-Json
                    if ($null -ne $ctData -and $ctData.TagName -and $ctData.Workload) {
                        $currentTaskName = "{0}/{1}" -f $ctData.TagName, $ctData.Workload
                    }
                } catch { }

                # Get the active task for this worker
                $workerTask = if ($tasksByPID.ContainsKey($wPID)) { $tasksByPID[$wPID] | Select-Object -First 1 } else { $null }

                if ($wState -eq "Busy" -and $Context.DispatchTimes.ContainsKey($wPID)) {
                    $taskTimeStr = Format-TimeSpan -Seconds ((Get-Date) - $Context.DispatchTimes[$wPID]).TotalSeconds
                }

                if ($workerTask -and $workerTask.Phase -eq "Detail") {
                    $wpExpected = $workerTask.OriginalExpectedCount -as [int]
                    if (-not $wpExpected) { $wpExpected = $workerTask.ExpectedCount -as [int] }
                    if ($wpExpected -gt 0) { $expectedStr = $wpExpected.ToString('N0') } else { $expectedStr = "N/A" }
                    if ($workerTask.PageSize) { $pageSizeStr = ($workerTask.PageSize -as [int]).ToString('N0') }

                    # Read Progress.log for live record count
                    try {
                        $progressLogPath = Join-Path $wDir "Progress.log"
                        if (Test-Path $progressLogPath) {
                            $lastLines = Get-Content -Path $progressLogPath -Tail 5 -ErrorAction Stop
                            for ($li = $lastLines.Count - 1; $li -ge 0; $li--) {
                                if ($lastLines[$li] -match 'Total:\s*([\d,]+)/.*\[(\d+)ms\]') {
                                    $currentCount = ($Matches[1] -replace ',', '') -as [int]
                                    if ($null -ne $currentCount) {
                                        $progressStr = $currentCount.ToString('N0')
                                        if ($wpExpected -and $wpExpected -gt 0) {
                                            $pctDone = [Math]::Round(($currentCount / $wpExpected) * 100)
                                            $progressStr += " ({0}%)" -f $pctDone
                                        }
                                        $inProgressItemTotal += $currentCount
                                    }
                                    $pageTimeMs = $Matches[2] -as [int]
                                    if ($null -ne $pageTimeMs) {
                                        if ($pageTimeMs -ge 60000) { $lastPageTimeStr = "{0:N1}m" -f ($pageTimeMs / 60000) }
                                        elseif ($pageTimeMs -ge 1000) { $lastPageTimeStr = "{0:N1}s" -f ($pageTimeMs / 1000) }
                                        else { $lastPageTimeStr = "${pageTimeMs}ms" }
                                    }
                                    break
                                }
                            }
                        }
                    } catch { }
                }
                elseif ($workerTask -and $workerTask.Phase -eq "Aggregate") {
                    $expectedStr = "N/A"
                    $progressStr = "-"
                }

                $workerStatusList += @{ PID = $wPID; State = $wState; CurrentTask = $currentTaskName; TaskTime = $taskTimeStr; Expected = $expectedStr; Progress = $progressStr; PageSize = $pageSizeStr; LastPage = $lastPageTimeStr }
            }

            # Build classifier groups for dashboard
            $classifierGroups = @{}
            foreach ($dt in $detTasks) {
                $groupKey = "{0} / {1}" -f $dt.TagName, $dt.Workload
                if (-not $classifierGroups.ContainsKey($groupKey)) {
                    $classifierGroups[$groupKey] = @{
                        TagName = $dt.TagName; Workload = $dt.Workload
                        Completed = 0; InProgress = 0; Pending = 0; Error = 0; Total = 0
                        TotalFiles = [long]0; CompletedFiles = [long]0; IsFallback = $false
                    }
                }
                $classifierGroups[$groupKey].Total++
                $taskFiles = ($dt.ExpectedCount -as [long])
                if ($taskFiles -gt 0) { $classifierGroups[$groupKey].TotalFiles += $taskFiles }
                switch ($dt.Status) {
                    "Completed"  { $classifierGroups[$groupKey].Completed++; if ($taskFiles -gt 0) { $classifierGroups[$groupKey].CompletedFiles += $taskFiles } }
                    "InProgress" { $classifierGroups[$groupKey].InProgress++ }
                    "Pending"    { $classifierGroups[$groupKey].Pending++ }
                    "Error"      { $classifierGroups[$groupKey].Error++ }
                }
                if ($dt.LocationType -eq "WorkloadFallback") { $classifierGroups[$groupKey].IsFallback = $true }
            }

            $completedDetailItemsWithProgress = $completedDetailItems + $inProgressItemTotal

            # Session keepalive: run lightweight command every 10 minutes
            if (((Get-Date) - $Context.LastKeepalive) -gt $Context.KeepaliveInterval) {
                try {
                    Get-Label -ResultSize 1 -WarningAction SilentlyContinue -ErrorAction Stop | Out-Null
                    $Context.LastKeepalive = Get-Date
                }
                catch {
                    Write-ExportLog -Message ("  Session keepalive failed: {0}" -f $_.Exception.Message) -Level Warning -LogOnly
                    try {
                        Disconnect-Compl8Compliance -LogOnly
                        if ($Context.AuthParams -and $Context.AuthParams.Count -gt 0) {
                            Connect-Compl8Compliance @($Context.AuthParams) -LogOnly
                            $Context.LastKeepalive = Get-Date
                        }
                    } catch {
                        Write-ExportLog -Message ("  Session reconnection failed (non-fatal): {0}" -f $_.Exception.Message) -Level Warning -LogOnly
                        $Context.LastKeepalive = Get-Date
                    }
                }
            }

            # Dynamic worker spawning via W hotkey
            Test-AddWorkerKeypress -ExportRunDirectory $Context.ExportRunDirectory `
                -WorkerProcesses ([ref]$Context.WorkerProcessesRef.Value) -NextWorkerNumber ([ref]$Context.NextWorkerNumberRef.Value)
            # WorkerProcesses is now an ArrayList (reference type) — new workers added via
            # Test-AddWorkerKeypress are visible to the dispatch engine immediately.

            try {
                Show-OrchestratorDashboard `
                    -Phase $displayPhase `
                    -Completed $displayCompleted `
                    -Total $displayTotal `
                    -Workers $workerStatusList `
                    -RecentErrors $LoopState.RecentErrors `
                    -RecentActivity $LoopState.RecentActivity `
                    -DispatchLog @() `
                    -ExportStartTime $Context.ExportStartTime `
                    -PhaseStartTime $Context.PipelineStartTime `
                    -CompletedItems $completedDetailItemsWithProgress `
                    -TotalItems $totalDetailItems `
                    -RemainingAggregates ($aggTotal - $aggDone) `
                    -ClassifierGroups $classifierGroups `
                    -TotalLocations $detTotal `
                    -TotalCompleted @($detTasks | Where-Object { $_.Status -eq "Completed" }).Count `
                    -TotalErrors $detErrors `
                    -TotalActive $detActive
            } catch {
                Write-ExportLog -Message ("  Dashboard render error (non-fatal): {0}" -f $_.Exception.Message) -Level Warning -LogOnly
            }
        }

        # --- CE Callback: OnCheckComplete ---
        # The loop is complete when ALL tasks (both aggregate and detail) are done.
        $ceOnCheckComplete = {
            param($Tasks, $LoopState, $Context)
            $pending = @($Tasks | Where-Object { $_.Status -in @("Pending", "InProgress") })
            return ($pending.Count -eq 0 -and $Tasks.Count -gt 0)
        }

        # --- CE Callback: OnAllWorkersDead ---
        $ceOnAllDead = {
            param($Tasks, $PendingCount, $Context)
            Write-ExportLog -Message "All CE workers dead - saving state for resume" -Level Error
            # Write both CSVs for resume compatibility
            $aggOnly = @($Tasks | Where-Object { $_.Phase -eq "Aggregate" })
            $detOnly = @($Tasks | Where-Object { $_.Phase -eq "Detail" })
            Write-TaskCsv -Path $Context.AggTaskCsvPath -Tasks $aggOnly
            if ($detOnly.Count -gt 0) {
                Write-TaskCsv -Path $Context.DetailTaskCsvPath -Tasks $detOnly
            }
        }

        # --- CE Callback: OnIterationComplete ---
        # Writes task CSVs incrementally, updates ExportPhase.txt, and generates
        # WorkloadFallback detail tasks for errored aggregates.
        $ceOnIterComplete = {
            param($Tasks, $LoopState, $Context)
            $aggOnly = @($Tasks | Where-Object { $_.Phase -eq "Aggregate" })
            $detOnly = @($Tasks | Where-Object { $_.Phase -eq "Detail" })

            # Write AggregateTasks.csv
            Write-TaskCsv -Path $Context.AggTaskCsvPath -Tasks $aggOnly

            # Write DetailTasks.csv if any detail tasks exist
            if ($detOnly.Count -gt 0) {
                Write-TaskCsv -Path $Context.DetailTaskCsvPath -Tasks $detOnly
            }

            # Generate WorkloadFallback detail tasks for newly errored aggregates
            $cePageSize = $Context.DefaultPageSize
            foreach ($errAgg in @($aggOnly | Where-Object { $_.Status -eq "Error" })) {
                $errKey = "{0}|{1}|{2}" -f $errAgg.TagType, $errAgg.TagName, $errAgg.Workload
                if ($Context.ProcessedAggErrorKeys.ContainsKey($errKey)) { continue }
                $Context.ProcessedAggErrorKeys[$errKey] = $true

                # Track for summary
                $Context.HasAggregateErrors = $true
                if ($Context.AggregateErrorTasks -notcontains ("{0}|{1}" -f $errAgg.TagName, $errAgg.Workload)) {
                    $Context.AggregateErrorTasks += "{0}|{1}" -f $errAgg.TagName, $errAgg.Workload
                }

                # Generate WorkloadFallback detail task
                $taskPageSize = [Math]::Max(500, $cePageSize)
                $errExpected = $errAgg.ExpectedCount -as [int]
                if (-not $errExpected -or $errExpected -le 0) { $errExpected = 0 }

                $fallbackTask = @{
                    Phase                 = "Detail"
                    TagType               = $errAgg.TagType
                    TagName               = $errAgg.TagName
                    Workload              = $errAgg.Workload
                    Location              = ""
                    LocationType          = "WorkloadFallback"
                    ExpectedCount         = $errExpected
                    OriginalExpectedCount = $errExpected
                    PageSize              = $taskPageSize
                    Status                = "Pending"
                    AssignedPID           = 0
                    ErrorMessage          = "Aggregate failed"
                }
                [void]$Tasks.Add($fallbackTask)
                Write-ExportLog -Message ("  WorkloadFallback detail task added for errored aggregate: {0}/{1}" -f $errAgg.TagName, $errAgg.Workload) -Level Info -LogOnly
            }

            # Update ExportPhase.txt
            $aggPending = @($aggOnly | Where-Object { $_.Status -in @("Pending", "InProgress") }).Count
            if ($aggPending -eq 0 -and -not $Context.PhaseTransitioned) {
                Write-ExportPhase -ExportDir $Context.ExportDir -Phase "Detail"
                $Context.PhaseTransitioned = $true
                Write-ExportLog -Message "Phase transitioned to Detail (all aggregates complete)" -Level Success -LogOnly
            }
        }

        # Build context with all state the callbacks need
        $ceContext = @{
            AggregateCsvPath      = $aggregateCsvPath
            AggTaskCsvPath        = $aggTaskCsvPath
            DetailTaskCsvPath     = $detailTaskCsvPath
            ExportDir             = $script:SharedExportDirectory
            ExportRunDirectory    = $script:SharedExportDirectory
            DefaultPageSize       = $cePageSize
            ExportStartTime       = $exportStartTime
            PipelineStartTime     = $pipelineStartTime
            DispatchTimes         = @{}
            LastKeepalive         = $lastSessionKeepalive
            KeepaliveInterval     = $keepaliveInterval
            AuthParams            = $script:AuthParams
            UnifiedTasks          = $unifiedTasks
            WorkerProcessesRef    = [ref]$workerProcesses
            NextWorkerNumberRef   = [ref]$nextWorkerNumber
            HasAggregateErrors    = $false
            AggregateErrorTasks   = @()
            ProcessedAggErrorKeys = @{}
            PhaseTransitioned     = $false
        }

        $ceLoopResult = Invoke-DispatchLoop `
            -ExportDir $script:SharedExportDirectory `
            -Tasks $unifiedTasks `
            -WorkerProcesses $workerProcesses `
            -Context $ceContext `
            -OnScanCompletions $ceOnScan `
            -OnMatchTask $ceOnMatch `
            -OnDispatchTask $ceOnDispatch `
            -OnShowDashboard $ceOnDashboard `
            -OnCompletionGeneratesTasks $ceOnGenerate `
            -OnCheckComplete $ceOnCheckComplete `
            -OnAllWorkersDead $ceOnAllDead `
            -OnIterationComplete $ceOnIterComplete `
            -SleepSeconds 2

        # Save final task CSVs
        $aggOnly = @($unifiedTasks | Where-Object { $_.Phase -eq "Aggregate" })
        $detOnly = @($unifiedTasks | Where-Object { $_.Phase -eq "Detail" })
        Write-TaskCsv -Path $aggTaskCsvPath -Tasks $aggOnly
        if ($detOnly.Count -gt 0) {
            Write-TaskCsv -Path $detailTaskCsvPath -Tasks $detOnly
        }

        # Extract aggregate error info from context
        $hasAggregateErrors = $ceContext.HasAggregateErrors
        $aggregateErrorTasks = $ceContext.AggregateErrorTasks

        # Also check central aggregate CSV for any ERROR rows
        if ($aggregateCsvPath -and (Test-Path $aggregateCsvPath)) {
            $errorLines = @(Get-Content -Path $aggregateCsvPath -Encoding UTF8 | Where-Object { $_ -match ',ERROR,' })
            if ($errorLines.Count -gt 0) {
                $hasAggregateErrors = $true
                foreach ($eLine in $errorLines) {
                    $parts = $eLine -split ','
                    if ($parts.Count -ge 4) {
                        $errEntry = "{0}|{1}" -f $parts[2], $parts[3]
                        if ($aggregateErrorTasks -notcontains $errEntry) {
                            $aggregateErrorTasks += $errEntry
                        }
                    }
                }
            }
        }

        # Summary logging
        $aggCompleted = @($aggOnly | Where-Object { $_.Status -eq "Completed" }).Count
        $aggErrors = @($aggOnly | Where-Object { $_.Status -eq "Error" }).Count
        $detCompleted = @($detOnly | Where-Object { $_.Status -eq "Completed" }).Count
        $detErrors = @($detOnly | Where-Object { $_.Status -eq "Error" }).Count
        Write-ExportLog -Message ("  Unified pipeline complete: Agg={0}/{1} ({2} err), Detail={3}/{4} ({5} err)" -f $aggCompleted, $aggOnly.Count, $aggErrors, $detCompleted, $detOnly.Count, $detErrors) -Level Success
    }
    elseif (-not $script:UseExistingAggregates) {
        # -- Single-terminal: orchestrator does the aggregate work itself --
        foreach ($tagType in $discoveredTagsByType.Keys) {
            $tagNames = $discoveredTagsByType[$tagType]

            # Process in batches
            $batches = [System.Collections.ArrayList]::new()
            for ($i = 0; $i -lt $tagNames.Count; $i += $batchSize) {
                $batchEnd = [Math]::Min($i + $batchSize - 1, $tagNames.Count - 1)
                [void]$batches.Add($tagNames[$i..$batchEnd])
            }

            Write-ExportLog -Message ("  Processing {0} classifiers in {1} batches" -f $tagNames.Count, $batches.Count) -Level Info

            $batchNum = 0
            foreach ($batch in $batches) {
                $batchNum++
                Write-ExportLog -Message ("`n  Batch {0}/{1}: {2} classifiers" -f $batchNum, $batches.Count, $batch.Count) -Level Info

                # Query fresh aggregate data
                $workPlan = New-ContentExplorerWorkPlan -TagType $tagType -TagNames $batch -Workloads $workloads `
                    -AggregateCsvPath $aggregateCsvPath -ExportRunDirectory $script:SharedExportDirectory

                if ($workPlan.HasErrors) {
                    $hasAggregateErrors = $true
                    $aggregateErrorTasks += @($workPlan.ErrorTasks)
                }

                # Store work plan tasks for later detail export
                if (-not $script:AllWorkPlanTasks) { $script:AllWorkPlanTasks = @() }
                $script:AllWorkPlanTasks += @($workPlan.Tasks)
            }
        }
    }

    #endregion

    #region -- Phase 6: Planning Phase --
    # In multi-terminal mode, detail tasks are generated by the pipeline's OnCompletionGeneratesTasks callback.
    # Single-terminal modes still need explicit planning here.

    if ($script:UseExistingAggregates -and $script:ExistingAggregatePath) {
        # Single-terminal with existing aggregates: build work plan from cached data
        if (-not $script:AllWorkPlanTasks) { $script:AllWorkPlanTasks = @() }
        foreach ($tagType in $discoveredTagsByType.Keys) {
            $tagNames = $discoveredTagsByType[$tagType]

            $batches = [System.Collections.ArrayList]::new()
            for ($i = 0; $i -lt $tagNames.Count; $i += $batchSize) {
                $batchEnd = [Math]::Min($i + $batchSize - 1, $tagNames.Count - 1)
                [void]$batches.Add($tagNames[$i..$batchEnd])
            }

            foreach ($batch in $batches) {
                $cachedResult = Import-AggregateDataFromCsv -CsvPath $script:ExistingAggregatePath -TagType $tagType -TagNames $batch -Workloads $workloads
                $cachedData = $cachedResult.TaskData

                foreach ($taskKey in $cachedData.Keys) {
                    $taskData = $cachedData[$taskKey]
                    $script:AllWorkPlanTasks += @{
                        TagType       = $taskData.TagType
                        TagName       = $taskData.TagName
                        Workload      = $taskData.Workload
                        ExpectedCount = $taskData.TotalCount
                        ExportedCount = 0
                        Locations     = $taskData.Locations
                        Status        = if ($taskData.HasError) { "Error" } else { "Pending" }
                        PageMetrics   = @()
                        ResponseTimes = @()
                        AggregateError = $taskData.HasError
                    }
                }

                if ($cachedResult.HasErrors) {
                    $hasAggregateErrors = $true
                    $aggregateErrorTasks += @($cachedResult.ErrorTasks)
                }
            }
        }
        Write-ExportLog -Message ("  Loaded {0} tasks from cached aggregates" -f $script:AllWorkPlanTasks.Count) -Level Info
    }

    # Warn about aggregate errors
    if ($hasAggregateErrors) {
        Write-Host ""
        Write-Banner -Title 'WARNING: AGGREGATE PHASE HAD ERRORS - RESULTS MAY BE INCOMPLETE' -Color 'Yellow' -Single
        Write-ExportLog -Message ("WARNING: {0} aggregate queries failed:" -f $aggregateErrorTasks.Count) -Level Warning
        foreach ($errorTask in $aggregateErrorTasks) {
            Write-ExportLog -Message ("  - {0}" -f $errorTask) -Level Warning
        }
        Write-ExportLog -Message "Export will continue but may be missing data for failed SITs/workloads" -Level Warning
        Write-Host ""
    }

    #endregion

    #region -- Phase 7: Detail Export --

    if (-not $isMultiTerminal) {
        # -- Single-terminal: orchestrator does the detail export work itself --
        Write-ExportPhase -ExportDir $script:SharedExportDirectory -Phase "Detail"

        foreach ($task in @($script:AllWorkPlanTasks)) {
            $taskKey = "{0}|{1}|{2}" -f $task.TagType, $task.TagName, $task.Workload

            # Skip tasks with no expected records
            if ($task.ExpectedCount -eq 0 -and $task.Status -ne "Error") {
                Write-ExportLog -Message ("    Skipping {0} / {1} - no data" -f $task.TagName, $task.Workload) -Level Info
                $task.Status = "Skipped"
                $completedTaskCounts[$taskKey] = 0
                continue
            }

            # Check local tracker for tasks this terminal already completed
            if (@($tracker.CompletedTasks) -contains $taskKey) {
                Write-ExportLog -Message ("    Skipping {0} / {1} - already completed locally" -f $task.TagName, $task.Workload) -Level Info
                if (-not $completedTaskCounts.ContainsKey($taskKey)) {
                    $completedTaskCounts[$taskKey] = $task.ExpectedCount
                }
                continue
            }

            # Show current progress
            $currentProgress = Get-ContentExplorerAggregateProgress -AggregateCsvPath $aggregateCsvPath -CompletedTasks $completedTaskCounts
            Write-ContentExplorerProgress -Progress $currentProgress -CurrentTaskKey $taskKey

            # Export task with progress tracking and adaptive paging
            # Aggregate-error tasks use a higher page size floor (no location data for adaptive sizing)
            $taskPageSize = if ($task.Status -eq "Error") { [Math]::Max(500, $cePageSize) } else { $cePageSize }
            # Output to Data/ContentExplorer/TagType/TagName/
            $classifierDir = Get-CEClassifierDir $exportDir $task.TagType $task.TagName

            $exportParams = @{
                Task                  = $task
                PageSize              = $taskPageSize
                ProgressLogPath       = $progressLogPath
                AdaptivePageSize      = $true
                TelemetryDatabasePath = $telemetryDbPath
                OutputDirectory       = $classifierDir
            }

            Export-ContentExplorerWithProgress @exportParams | Out-Null

            # Track exported count for this task
            $exportedCount = if ($task.ExportedCount) { $task.ExportedCount } else { 0 }
            $completedTaskCounts[$taskKey] = $exportedCount

            if ($exportedCount -gt 0) {
                $tracker.TotalExported += $exportedCount
                Write-ExportLog -Message ("    Completed: {0} / {1} - {2} records" -f $task.TagName, $task.Workload, $exportedCount) -Level Success
            }
            else {
                Write-ExportLog -Message ("    Completed: {0} / {1} - 0 records" -f $task.TagName, $task.Workload) -Level Info
            }

            # Update tracker with task metrics and output file mapping
            if ($task.PageMetrics) {
                $tracker.TaskMetrics = ($tracker.TaskMetrics ?? @()) + @{
                    TaskKey             = $taskKey
                    TotalPages          = $task.TotalPages
                    TotalTimeMs         = $task.TotalTimeMs
                    FinalDegradationPct = $task.FinalDegradationPct
                    AvgDegradationPct   = $task.AvgDegradationPct
                }
            }

            # Track output files in RunTracker
            if (-not $tracker.OutputFiles) { $tracker.OutputFiles = @() }
            if ($exportedCount -gt 0) {
                $tracker.OutputFiles += @{
                    TaskKey         = $taskKey
                    OutputDirectory = $classifierDir
                    RecordCount     = $exportedCount
                    Pages           = $task.TotalPages
                    CompletedTime   = (Get-Date).ToString("o")
                }
            }

            $tracker.CompletedTasks += $taskKey
            Save-ContentExplorerRunTracker -Tracker $tracker -TrackerPath $trackerPath
        }
    }

    #endregion

    #region -- Phase 7.5: Retry Bucket Detection --
    # Detect tasks with >2% discrepancy between expected and actual counts

    $retryBucketTasks = @()
    $detailCsvForRetry = Join-Path (Get-CoordinationDir $script:SharedExportDirectory) "DetailTasks.csv"
    if (Test-Path $detailCsvForRetry) {
        $finalDetailTasks = Read-TaskCsv -Path $detailCsvForRetry
        $retryBucketTasks = @(Get-RetryBucketTasks -DetailTasks $finalDetailTasks)
        if ($retryBucketTasks.Count -gt 0) {
            $retryTasksCsvPath = Join-Path (Get-CoordinationDir $script:SharedExportDirectory) "RetryTasks.csv"
            Write-RetryTasksCsv -Path $retryTasksCsvPath -RetryTasks $retryBucketTasks
            Write-ExportLog -Message ("  Wrote {0} retry tasks to RetryTasks.csv" -f $retryBucketTasks.Count) -Level Info
        }
    }

    #endregion

    #region -- Phase 8: Summary --

    $exportDir = $script:SharedExportDirectory

    # Summary stats
    Write-ExportLog -Message "`n--- Content Explorer Summary ---" -Level Info
    if (Test-Path $aggregateCsvPath) {
        $finalProgress = Get-ContentExplorerAggregateProgress -AggregateCsvPath $aggregateCsvPath -CompletedTasks $completedTaskCounts
        $finalPct = if ($finalProgress.TotalExpected -gt 0) {
            [Math]::Round(($finalProgress.TotalExported / $finalProgress.TotalExpected) * 100, 1)
        } else { 100 }

        Write-ExportLog -Message ("  Tasks completed: {0}/{1}" -f $finalProgress.CompletedTasks, $finalProgress.TotalTasks) -Level Info
        Write-ExportLog -Message ("  Records exported: {0}/{1} ({2}%)" -f $finalProgress.TotalExported.ToString('N0'), $finalProgress.TotalExpected.ToString('N0'), $finalPct) -Level Info
        Write-ExportLog -Message ("  Data output: {0}" -f (Get-CEDataDir $exportDir)) -Level Info
        Write-ExportLog -Message ("  Aggregate CSV: {0}" -f $aggregateCsvPath) -Level Info
    }
    else {
        Write-ExportLog -Message "  No aggregate data found" -Level Warning
    }

    # Display retry bucket summary
    Show-RetryBucketSummary -RetryTasks $retryBucketTasks -ExportDir $exportDir

    # Write remaining (non-completed) tasks for follow-on runs
    $remainingCount = Write-RemainingTasksCsv -ExportDir $exportDir
    if ($remainingCount -gt 0) {
        Write-ExportLog -Message ("  Remaining tasks: {0} (see RemainingTasks.csv)" -f $remainingCount) -Level Warning
        Write-ExportLog -Message ("  To re-run: .\Export-Compl8Configuration.ps1 -CETasksCsv ""{0}""" -f (Join-Path $exportDir "RemainingTasks.csv")) -Level Info
    }

    #endregion

    #region -- Phase 9: Cleanup --

    # Final tracker save
    $tracker.Status = "Completed"
    Save-ContentExplorerRunTracker -Tracker $tracker -TrackerPath $trackerPath

    # Shut down worker processes
    # Workers are spawned with -NoExit so their terminals stay open even after the export
    # script finishes. The Completed phase signal tells workers to exit their main loop,
    # but the pwsh process remains. We give a grace period, then close remaining terminals.
    if ($workerProcesses.Count -gt 0) {
        Write-ExportLog -Message ("Shutting down {0} worker process(es)..." -f $workerProcesses.Count) -Level Info

        # Grace period: let workers detect the Completed phase and exit their main loop
        $graceWaitMs = 15000  # 15 seconds
        $cleanlyExited = 0
        $closedByOrchestrator = 0
        $alreadyExited = 0

        # First pass: count already-exited workers and wait briefly for the rest
        $stillRunning = @()
        foreach ($worker in $workerProcesses) {
            if (-not $worker.Process -or $worker.Process.HasExited) {
                $alreadyExited++
                Write-ExportLog -Message ("  Worker PID {0}: already exited" -f $worker.PID) -Level Info -LogOnly
            }
            else {
                $stillRunning += $worker
            }
        }

        if ($stillRunning.Count -gt 0) {
            Write-ExportLog -Message ("  {0} worker(s) still running - waiting {1}s for graceful exit..." -f $stillRunning.Count, ($graceWaitMs / 1000)) -Level Info
            Start-Sleep -Milliseconds $graceWaitMs

            # Second pass: check who exited during grace period, close the rest
            foreach ($worker in $stillRunning) {
                if ($worker.Process.HasExited) {
                    $cleanlyExited++
                    Write-ExportLog -Message ("  Worker PID {0}: exited cleanly" -f $worker.PID) -Level Info -LogOnly
                }
                else {
                    # Worker still running (likely due to -NoExit keeping terminal open)
                    try {
                        Stop-Process -Id $worker.PID -Force -ErrorAction Stop
                        $closedByOrchestrator++
                        Write-ExportLog -Message ("  Worker PID {0}: closed by orchestrator" -f $worker.PID) -Level Info -LogOnly
                    }
                    catch {
                        Write-ExportLog -Message ("  Worker PID {0}: failed to close ({1})" -f $worker.PID, $_.Exception.Message) -Level Warning -LogOnly
                    }
                }
            }
        }

        Write-ExportLog -Message ("Worker shutdown complete: {0} already exited, {1} exited cleanly, {2} closed by orchestrator" -f $alreadyExited, $cleanlyExited, $closedByOrchestrator) -Level Info
    }

    # Write manifest and set final phase
    Write-CEManifest -ExportDir $script:SharedExportDirectory
    Write-ExportPhase -ExportDir $script:SharedExportDirectory -Phase "Completed"
    Write-ExportLog -Message "Export phase set to Completed" -Level Success

    #endregion
}


#region Activity Explorer Multi-Terminal Functions

function Invoke-ActivityExplorerWorker {
    <#
    .SYNOPSIS
        Activity Explorer worker using file-drop coordination.
    .DESCRIPTION
        Runs in a spawned worker terminal. Receives per-day tasks from the orchestrator
        via file-drop (nexttask/currenttask files) and writes output to Data/ActivityExplorer/YYYY-MM-DD/.
    .PARAMETER WorkerExportDir
        The export run directory.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$WorkerExportDir
    )

    $exportDir = $WorkerExportDir

    Write-Host "`n  AE Worker starting in file-drop mode..." -ForegroundColor Yellow
    Write-Host ("  Export directory: {0}" -f $exportDir) -ForegroundColor Gray

    # Create Worker coordination subfolder
    $workerDir = Get-WorkerCoordDir $exportDir $PID
    if (-not (Test-Path $workerDir)) {
        New-Item -ItemType Directory -Force -Path $workerDir | Out-Null
    }

    # Initialize logging
    Initialize-ExportLog -LogDirectory (Get-LogsDir $exportDir)

    $script:ErrorLogPath = Join-Path (Get-LogsDir $exportDir) "ExportProject-Errors.log"
    $script:WorkerErrorLogPath = Join-Path (Get-LogsDir $exportDir) "ExportErrors-$PID.log"
    $progressLogPath = Join-Path $workerDir "Progress.log"
    $signalSigningKey = Get-ExportRunSigningKey -ExportDir $exportDir -CreateIfMissing

    Write-ExportLog -Message "AE Worker PID $PID started (file-drop), output folder: $workerDir" -Level Info
    Write-ProgressEntry -Path $progressLogPath -Message "AE Worker PID $PID started"

    # Load Activity Explorer filters (prefer saved manifest for consistency)
    $aeConfigPath = Join-Path $PSScriptRoot "ConfigFiles\ActivityExplorerSelector.json"
    $filters = Resolve-AEFilters -ExportRunDirectory $exportDir -ConfigPath $aeConfigPath

    $lastWorkerActivity = Get-Date
    $workerInactivityLimit = New-TimeSpan -Minutes $script:CEWorkerInactivityMinutes

    Write-ExportLog -Message "AE Worker entering file-drop task loop" -Level Info

    while ($true) {
        $task = Receive-WorkerTask -WorkerDir $workerDir -ExportDir $exportDir

        if (-not $task) {
            $phase = Read-ExportPhase -ExportDir $exportDir
            if ($phase -eq "AECompleted") {
                Write-ExportLog -Message "Project phase is $phase - AE worker exiting cleanly" -Level Info
                Write-ProgressEntry -Path $progressLogPath -Message "Phase is $phase - exiting"
                break
            }

            # Inactivity timeout
            if (((Get-Date) - $lastWorkerActivity) -gt $workerInactivityLimit) {
                $finalCheck = Receive-WorkerTask -WorkerDir $workerDir -ExportDir $exportDir
                if ($finalCheck) {
                    $lastWorkerActivity = Get-Date
                    $task = $finalCheck
                } else {
                    Write-ExportLog -Message "No activity for 35 minutes - AE worker exiting" -Level Warning
                    Write-ProgressEntry -Path $progressLogPath -Message "Inactivity timeout - exiting"
                    break
                }
            }

            if (-not $task) {
                Start-Sleep -Seconds 2
                continue
            }
        }

        # Reset activity timer
        $lastWorkerActivity = Get-Date

        $taskDay = $task.Day
        # Handle both DateTime (ConvertFrom-Json auto-converts ISO 8601) and string (from CSV)
        $taskStartTime = if ($task.StartTime -is [datetime]) {
            $task.StartTime.ToUniversalTime()
        } else {
            [datetime]::Parse($task.StartTime, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::RoundtripKind).ToUniversalTime()
        }
        $taskEndTime = if ($task.EndTime -is [datetime]) {
            $task.EndTime.ToUniversalTime()
        } else {
            [datetime]::Parse($task.EndTime, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::RoundtripKind).ToUniversalTime()
        }
        $taskPageSize = if ($task.PageSize) { $task.PageSize -as [int] } else { 5000 }

        Write-ExportLog -Message ("  AE Worker received task: Day {0} ({1} to {2})" -f $taskDay, $taskStartTime.ToString("yyyy-MM-dd HH:mm"), $taskEndTime.ToString("yyyy-MM-dd HH:mm")) -Level Info
        Write-ProgressEntry -Path $progressLogPath -Message ("Received task: Day {0}" -f $taskDay)

        # Create per-day output directory
        $dayDir = Get-AEDayDir $exportDir $taskDay
        if (-not (Test-Path $dayDir)) {
            New-Item -ItemType Directory -Force -Path $dayDir | Out-Null
        }

        # Initialize per-day tracker
        $trackerPath = Join-Path $dayDir "RunTracker.json"
        $tracker = Get-ActivityExplorerRunTracker -TrackerPath $trackerPath

        $taskStart = Get-Date
        try {
            $exportParams = @{
                StartTime       = $taskStartTime
                EndTime         = $taskEndTime
                PageSize        = $taskPageSize
                Filters         = $filters
                OutputDirectory = $dayDir
                Tracker         = $tracker
                TrackerPath     = $trackerPath
                ProgressLogPath = $progressLogPath
            }

            # If tracker has progress from a previous attempt, resume
            if ($tracker.CompletedPages -and $tracker.CompletedPages.Count -gt 0) {
                $exportParams['Resume'] = $true
                Write-ExportLog -Message ("  Resuming day {0} from page {1}" -f $taskDay, $tracker.CompletedPages.Count) -Level Info
            }

            $exportResult = Export-ActivityExplorerWithProgress @exportParams
            $taskElapsed = (Get-Date) - $taskStart

            Write-ExportLog -Message ("  Day {0} complete: {1} records, {2} pages in {3}" -f $taskDay, $exportResult.TotalRecords, $exportResult.PageCount, (Format-TimeSpan -Seconds $taskElapsed.TotalSeconds)) -Level Success
            Write-ProgressEntry -Path $progressLogPath -Message ("Day {0} complete: {1} records, {2} pages" -f $taskDay, $exportResult.TotalRecords, $exportResult.PageCount)

            # Write completion signal
            $completionsDir = Get-CompletionsDir $exportDir
            if (-not (Test-Path $completionsDir)) {
                New-Item -ItemType Directory -Force -Path $completionsDir | Out-Null
            }
            $doneFile = Join-Path $completionsDir ("ae-done-{0}-{1}.txt" -f $taskDay, $PID)
            $donePayload = @{
                Day            = $taskDay
                RecordCount    = $exportResult.TotalRecords
                PageCount      = $exportResult.PageCount
                ElapsedSeconds = [int]$taskElapsed.TotalSeconds
            }
            (ConvertTo-SignedEnvelopeJson -Payload $donePayload -SigningKey $signalSigningKey) | Set-Content -Path $doneFile -Encoding UTF8

        }
        catch {
            $taskElapsed = (Get-Date) - $taskStart
            Write-ExportLog -Message ("  Day {0} FAILED: {1}" -f $taskDay, $_.Exception.Message) -Level Error
            Write-ProgressEntry -Path $progressLogPath -Message ("Day {0} FAILED: {1}" -f $taskDay, $_.Exception.Message)

            # Write error signal
            $completionsDir = Get-CompletionsDir $exportDir
            if (-not (Test-Path $completionsDir)) {
                New-Item -ItemType Directory -Force -Path $completionsDir | Out-Null
            }
            $errorFile = Join-Path $completionsDir ("error-ae-{0}-{1}.txt" -f $taskDay, $PID)
            $errorPayload = @{
                Day          = $taskDay
                ErrorMessage = $_.Exception.Message
                ErrorType    = $_.Exception.GetType().Name
            }
            (ConvertTo-SignedEnvelopeJson -Payload $errorPayload -SigningKey $signalSigningKey) | Set-Content -Path $errorFile -Encoding UTF8

            # Error logging
            if ($script:ErrorLogPath) {
                Write-ExportErrorLog -ErrorLogPath $script:ErrorLogPath -Context "AE Worker Day Export" -TaskKey $taskDay -ErrorRecord $_ -AdditionalData @{ Day = $taskDay }
            }
            if ($script:WorkerErrorLogPath) {
                Write-ExportErrorLog -ErrorLogPath $script:WorkerErrorLogPath -Context "AE Worker Day Export" -TaskKey $taskDay -ErrorRecord $_ -AdditionalData @{ Day = $taskDay }
            }

            # Auth errors: attempt reconnect
            $errorInfo = Get-HttpErrorExplanation -ErrorMessage $_.Exception.Message -ErrorRecord $_
            if ($errorInfo.Category -eq "AuthError" -or $_.Exception -is [System.Management.Automation.CommandNotFoundException]) {
                Write-ExportLog -Message "    AUTH/CONNECTION error - attempting recovery..." -Level Warning
                try {
                    Disconnect-Compl8Compliance
                    if ($script:AuthParams -and $script:AuthParams.Count -gt 0) {
                        $reAuthResult = Connect-Compl8Compliance @script:AuthParams
                        if ($reAuthResult) {
                            Write-ExportLog -Message "    Recovery successful" -Level Success
                        }
                        else {
                            Write-ExportLog -Message "    Recovery failed - worker exiting" -Level Error
                            Complete-WorkerTask -WorkerDir $workerDir
                            break
                        }
                    }
                    else {
                        Write-ExportLog -Message "    No auth params - worker exiting" -Level Error
                        Complete-WorkerTask -WorkerDir $workerDir
                        break
                    }
                }
                catch {
                    Write-ExportLog -Message ("    Recovery exception: {0} - worker exiting" -f $_.Exception.Message) -Level Error
                    Complete-WorkerTask -WorkerDir $workerDir
                    break
                }
            }
        }

        Complete-WorkerTask -WorkerDir $workerDir
        $lastWorkerActivity = Get-Date
    }

    Write-ExportLog -Message "AE Worker PID $PID finished" -Level Info
    Write-ProgressEntry -Path $progressLogPath -Message "Worker finished"
}

function Invoke-AEMultiExport {
    <#
    .SYNOPSIS
        Activity Explorer multi-terminal orchestrator (fresh export or resume).
    .DESCRIPTION
        Splits the date range into per-day tasks (or reloads from CSV on resume),
        spawns worker terminals, dispatches tasks via file-drop, monitors progress,
        reclaims stale tasks, and completes. When -IsResume is specified,
        reads existing AEDayTasks.csv and spawns new workers for incomplete tasks.
    .PARAMETER IsResume
        When set, resumes a previous export instead of starting fresh.
    .PARAMETER ResumeDir
        Path to the export directory to resume (required when -IsResume is set).
    .PARAMETER ResumeWorkerCount
        Number of workers to spawn for resume. 0 = single-terminal sequential resume.
    #>
    param(
        [switch]$IsResume,
        [string]$ResumeDir = "",
        [int]$ResumeWorkerCount = 0
    )

    if ($IsResume) {
        # ---- Resume path: reload state from existing export ----
        Write-ExportLog -Message "`n========== Activity Explorer Resume ==========" -Level Info
        Write-ExportLog -Message ("Resuming from: {0}" -f $ResumeDir) -Level Info

        $script:SharedExportDirectory = $ResumeDir
        $script:ExportRunDirectory = $ResumeDir
        $exportDir = $ResumeDir

        # Read existing task CSV
        $aeTaskCsvPath = Join-Path (Get-CoordinationDir $exportDir) "AEDayTasks.csv"
        if (-not (Test-Path $aeTaskCsvPath)) {
            Write-ExportLog -Message "No AEDayTasks.csv found - cannot resume" -Level Error
            return
        }

        $dayTasks = [System.Collections.ArrayList]::new()
        $csvTasks = Read-AETaskCsv -Path $aeTaskCsvPath
        foreach ($ct in $csvTasks) {
            [void]$dayTasks.Add(@{
                Day          = $ct.Day
                StartTime    = $ct.StartTime
                EndTime      = $ct.EndTime
                AssignedPID  = 0
                Status       = if ($ct.Status -eq "Completed") { "Completed" } else { "Pending" }
                PageCount    = if ($ct.Status -eq "Completed") { $ct.PageCount -as [int] } else { 0 }
                RecordCount  = if ($ct.Status -eq "Completed") { $ct.RecordCount -as [int] } else { 0 }
                ErrorMessage = ""
            })
        }

        $taskCount = $dayTasks.Count
        $pendingCount = @($dayTasks | Where-Object { $_.Status -eq "Pending" }).Count
        $completedCount = @($dayTasks | Where-Object { $_.Status -eq "Completed" }).Count
        Write-ExportLog -Message ("Tasks: {0} pending, {1} already completed, {2} total" -f $pendingCount, $completedCount, $taskCount) -Level Info

        if ($pendingCount -eq 0) {
            Write-ExportLog -Message "All tasks already completed - nothing to resume" -Level Info
            Write-AEManifest -ExportDir $exportDir
            Write-ExportPhase -ExportDir $exportDir -Phase "AECompleted"
            return
        }

        # Write updated tasks and phase
        Write-AETaskCsv -Path $aeTaskCsvPath -Tasks $dayTasks
        Write-ExportPhase -ExportDir $exportDir -Phase "AEExport"

        # Single-terminal resume: process tasks sequentially without spawning workers
        if ($ResumeWorkerCount -lt 2) {
            Write-ExportLog -Message "Single-terminal resume mode" -Level Info

            # Load filters (prefer saved manifest for consistency)
            $aeConfigPath = Join-Path $PSScriptRoot "ConfigFiles\ActivityExplorerSelector.json"
            $filters = Resolve-AEFilters -ExportRunDirectory $exportDir -ConfigPath $aeConfigPath

            # Process each pending task using a "virtual worker" coordination directory
            $workerDir = Get-WorkerCoordDir $exportDir $PID
            if (-not (Test-Path $workerDir)) {
                New-Item -ItemType Directory -Force -Path $workerDir | Out-Null
            }

            foreach ($task in $dayTasks) {
                if ($task.Status -ne "Pending") { continue }

                $taskDay = $task.Day
                $taskStartTime = if ($task.StartTime -is [datetime]) {
                    $task.StartTime.ToUniversalTime()
                } else {
                    [datetime]::Parse($task.StartTime, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::RoundtripKind).ToUniversalTime()
                }
                $taskEndTime = if ($task.EndTime -is [datetime]) {
                    $task.EndTime.ToUniversalTime()
                } else {
                    [datetime]::Parse($task.EndTime, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::RoundtripKind).ToUniversalTime()
                }

                Write-ExportLog -Message ("  Processing day {0}..." -f $taskDay) -Level Info

                $dayDir = Get-AEDayDir $exportDir $taskDay
                if (-not (Test-Path $dayDir)) {
                    New-Item -ItemType Directory -Force -Path $dayDir | Out-Null
                }

                $trackerPath = Join-Path $dayDir "RunTracker.json"
                $tracker = Get-ActivityExplorerRunTracker -TrackerPath $trackerPath
                $progressLogPath = Join-Path $workerDir "Progress.log"

                try {
                    $exportParams = @{
                        StartTime       = $taskStartTime
                        EndTime         = $taskEndTime
                        PageSize        = $PageSize
                        Filters         = $filters
                        OutputDirectory = $dayDir
                        Tracker         = $tracker
                        TrackerPath     = $trackerPath
                        ProgressLogPath = $progressLogPath
                    }
                    if ($tracker.CompletedPages -and $tracker.CompletedPages.Count -gt 0) {
                        $exportParams['Resume'] = $true
                    }

                    $exportResult = Export-ActivityExplorerWithProgress @exportParams

                    $task.Status = "Completed"
                    $task.RecordCount = $exportResult.TotalRecords
                    $task.PageCount = $exportResult.PageCount
                    Write-ExportLog -Message ("  Day {0} complete: {1} records, {2} pages" -f $taskDay, $exportResult.TotalRecords, $exportResult.PageCount) -Level Success
                }
                catch {
                    $task.Status = "Error"
                    $task.ErrorMessage = $_.Exception.Message
                    Write-ExportLog -Message ("  Day {0} failed: {1}" -f $taskDay, $_.Exception.Message) -Level Error
                }

                Write-AETaskCsv -Path $aeTaskCsvPath -Tasks $dayTasks
            }

            Write-AEManifest -ExportDir $exportDir
            Write-ExportPhase -ExportDir $exportDir -Phase "AECompleted"
            Write-ExportLog -Message "Activity Explorer resume complete" -Level Success
            return
        }

        # Multi-terminal resume: cap workers to pending count, then fall through to shared dispatch
        $actualWorkers = [Math]::Min($ResumeWorkerCount, $pendingCount)
    }
    else {
        # ---- Fresh export path: generate tasks and spawn workers ----
        Write-ExportLog -Message "`n========== Activity Explorer Multi-Terminal Export ==========" -Level Info

        # 1. Calculate date range (same as single-terminal)
        $endTime = [DateTime]::UtcNow.AddMinutes(-5)
        $startTime = $endTime.AddDays(-$PastDays)

        Write-ExportLog -Message "Time range: $($startTime.ToString('yyyy-MM-dd HH:mm')) to $($endTime.ToString('yyyy-MM-dd HH:mm')) UTC" -Level Info
        Write-ExportLog -Message "Past days: $PastDays | Workers: $AEWorkerCount" -Level Info

        # 2. Generate per-day tasks
        # All dates must be explicitly UTC to avoid timezone drift when serialized/parsed by workers
        $dayTasks = [System.Collections.ArrayList]::new()
        $currentDay = [DateTime]::SpecifyKind($startTime.Date, [DateTimeKind]::Utc)
        while ($currentDay -lt $endTime) {
            $dayStart = if ($currentDay -lt $startTime) { $startTime } else { $currentDay }
            $nextDay = $currentDay.AddDays(1)
            $dayEnd = if ($nextDay -gt $endTime) { $endTime } else { $nextDay }

            if ($dayStart -lt $dayEnd) {
                [void]$dayTasks.Add(@{
                    Day          = $currentDay.ToString("yyyy-MM-dd")
                    StartTime    = $dayStart.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
                    EndTime      = $dayEnd.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
                    AssignedPID  = 0
                    Status       = "Pending"
                    PageCount    = 0
                    RecordCount  = 0
                    ErrorMessage = ""
                })
            }
            $currentDay = $nextDay
        }

        $taskCount = $dayTasks.Count
        Write-ExportLog -Message ("Generated {0} day task(s)" -f $taskCount) -Level Info

        if ($taskCount -eq 0) {
            Write-ExportLog -Message "No tasks to export (date range empty)" -Level Warning
            return
        }

        # Cap workers to task count
        $actualWorkers = [Math]::Min($AEWorkerCount, $taskCount)
        if ($actualWorkers -lt $AEWorkerCount) {
            Write-ExportLog -Message ("Capped workers from {0} to {1} (only {2} day tasks)" -f $AEWorkerCount, $actualWorkers, $taskCount) -Level Info
        }

        # 3. Write coordination files
        $script:SharedExportDirectory = $script:ExportRunDirectory
        $exportDir = $script:ExportRunDirectory
        Write-ExportType -ExportDir $exportDir -Type "ActivityExplorer"
        Write-ExportPhase -ExportDir $exportDir -Phase "AEExport"

        $aeTaskCsvPath = Join-Path (Get-CoordinationDir $exportDir) "AEDayTasks.csv"
        Write-AETaskCsv -Path $aeTaskCsvPath -Tasks $dayTasks

        # Save export settings manifest for resume consistency
        $configPath = Join-Path $PSScriptRoot "ConfigFiles\ActivityExplorerSelector.json"
        $selectorConfig = Read-JsonConfig -Path $configPath
        Save-ExportSettings -ExportRunDirectory $exportDir -ExportType "ActivityExplorer" -Settings @{
            PastDays       = $PastDays
            PageSize       = $PageSize
            SelectorConfig = $selectorConfig
        }
    }

    # ==== Shared multi-terminal dispatch (both fresh and resume paths) ====

    # Ensure Completions directory exists
    $completionsDir = Get-CompletionsDir $exportDir
    if (-not (Test-Path $completionsDir)) {
        New-Item -ItemType Directory -Force -Path $completionsDir | Out-Null
    }

    Write-ExportLog -Message ("Task CSV: {0}" -f $aeTaskCsvPath) -Level Info

    # Spawn workers
    $workerProcesses = Start-WorkerTerminals -ExportRunDirectory $exportDir -Count $actualWorkers
    if (-not $workerProcesses -or $workerProcesses.Count -eq 0) {
        Write-ExportLog -Message "No workers spawned - aborting" -Level Error
        return
    }
    Write-ExportLog -Message ("{0} worker(s) spawned" -f $workerProcesses.Count) -Level Info

    # Dispatch loop via Invoke-DispatchLoop engine
    $exportStartTime = Get-Date
    Reset-AEDashboard

    # --- AE Callbacks ---
    $aeOnScan = {
        param($ExportDir, $WorkerDirs, $Context)
        $completed = @()
        $errors = @()
        $completionsDir = Get-CompletionsDir $ExportDir
        $signingKey = Get-ExportRunSigningKey -ExportDir $ExportDir

        # Scan ae-done-*.txt
        Get-ChildItem -Path $completionsDir -Filter "ae-done-*.txt" -ErrorAction SilentlyContinue |
            ForEach-Object {
                try {
                    $data = ConvertFrom-SignedEnvelopeJson -Json (Get-Content -Raw -Path $_.FullName) -SigningKey $signingKey -RequireSignature:([bool]$signingKey) -Context ("AE completion file {0}" -f $_.Name)
                    if ($null -ne $data) {
                        $completed += @{
                            Day         = $data.Day
                            RecordCount = $data.RecordCount
                            PageCount   = $data.PageCount
                            Message     = "Day $($data.Day) completed: $([int]$data.RecordCount) records, $([int]$data.PageCount) pages"
                        }
                    }
                    Rename-Item -Path $_.FullName -NewName ($_.Name + ".done") -Force -ErrorAction SilentlyContinue
                } catch { Write-Verbose "Failed to parse AE completion file: $($_.Exception.Message)" }
            }

        # Scan error-ae-*.txt
        Get-ChildItem -Path $completionsDir -Filter "error-ae-*.txt" -ErrorAction SilentlyContinue |
            ForEach-Object {
                try {
                    $data = ConvertFrom-SignedEnvelopeJson -Json (Get-Content -Raw -Path $_.FullName) -SigningKey $signingKey -RequireSignature:([bool]$signingKey) -Context ("AE error file {0}" -f $_.Name)
                    if ($null -ne $data) {
                        $errors += @{
                            Day          = $data.Day
                            ErrorMessage = $data.ErrorMessage
                            Message      = "Day $($data.Day): $($data.ErrorMessage)"
                        }
                    }
                    Rename-Item -Path $_.FullName -NewName ($_.Name + ".done") -Force -ErrorAction SilentlyContinue
                } catch { Write-Verbose "Failed to parse AE error file: $($_.Exception.Message)" }
            }

        return @{ CompletedTasks = $completed; ErrorTasks = $errors }
    }

    $aeOnMatch = {
        param($Data, $Tasks, $Context)
        $Tasks | Where-Object { $_.Day -eq $Data.Day -and $_.Status -eq "InProgress" } | Select-Object -First 1
    }

    $aeOnDispatch = {
        param($Worker, $NextTask, $Context)
        $taskData = @{
            Phase     = "AEExport"
            Day       = $NextTask.Day
            StartTime = $NextTask.StartTime
            EndTime   = $NextTask.EndTime
            PageSize  = $Context.PageSize
        }
        return (Send-WorkerTask -WorkerDir $Worker.WorkerDir -TaskData $taskData -ExportDir $Context.ExportDir)
    }

    $aeOnDashboard = {
        param($LoopState, $Context)
        $workerInfoList = @()
        $totalRecords = [long]0
        # Track per-day progress fraction for in-progress days (day string -> 0.0-1.0)
        $dayProgressMap = @{}
        foreach ($wp in $LoopState.WorkerProcesses) {
            $wState = Get-WorkerState -WorkerDir $wp.WorkerDir -WorkerPID $wp.PID
            $currentDay = $null
            $currentPages = $null
            $currentRecords = $null
            $recordPct = $null
            $activeTask = $Context.DayTasks | Where-Object { ($_.AssignedPID -as [int]) -eq $wp.PID -and $_.Status -eq "InProgress" } | Select-Object -First 1
            if ($activeTask) {
                $currentDay = $activeTask.Day
                $dayPageDir = Get-AEDayDir $Context.ExportDir $activeTask.Day
                if (Test-Path $dayPageDir) {
                    $currentPages = @(Get-ChildItem -Path $dayPageDir -Filter "Page-*.json" -ErrorAction SilentlyContinue).Count
                    $dayTrackerPath = Join-Path $dayPageDir "RunTracker.json"
                    if (Test-Path $dayTrackerPath) {
                        try {
                            $dayTracker = Get-Content -Raw -Path $dayTrackerPath | ConvertFrom-Json
                            if ($null -ne $dayTracker) {
                                $currentRecords = $dayTracker.TotalRecords
                                if ($currentRecords) { $totalRecords += ($currentRecords -as [long]) }
                                if ($dayTracker.TotalAvailable -and $dayTracker.TotalAvailable -gt 0) {
                                    $dayFraction = [Math]::Min(1.0, ($dayTracker.TotalRecords / $dayTracker.TotalAvailable))
                                    $dayProgressMap[$activeTask.Day] = $dayFraction
                                    $recordPct = [Math]::Round($dayFraction * 100, 1)
                                }
                            }
                        }
                        catch {
                            Write-Verbose "Could not read day tracker: $($_.Exception.Message)"
                        }
                    }
                }
            }
            $workerInfoList += @{ PID = $wp.PID; State = $wState; CurrentDay = $currentDay; Pages = $currentPages; Records = $currentRecords; RecordPct = $recordPct }
        }

        # Sum records from completed days
        foreach ($t in $Context.DayTasks) {
            if ($t.Status -eq "Completed" -and $t.RecordCount) {
                $totalRecords += ($t.RecordCount -as [long])
            }
        }

        # Weighted percentage: each day = equal share, in-progress days contribute proportionally
        $totalDays = $LoopState.TotalCount
        $weightedPct = [double]0
        if ($totalDays -gt 0) {
            $dayShare = 100.0 / $totalDays
            foreach ($t in $Context.DayTasks) {
                if ($t.Status -eq "Completed") {
                    $weightedPct += $dayShare
                }
                elseif ($t.Status -eq "InProgress" -and $dayProgressMap.ContainsKey($t.Day)) {
                    $weightedPct += $dayShare * $dayProgressMap[$t.Day]
                }
            }
        }

        Show-AEDashboard `
            -Phase "AEExport" `
            -Completed $LoopState.CompletedCount `
            -Total $LoopState.TotalCount `
            -Workers $workerInfoList `
            -DayTasks $Context.DayTasks `
            -RecentActivity $LoopState.RecentActivity `
            -RecentErrors $LoopState.RecentErrors `
            -ExportStartTime $Context.ExportStartTime `
            -TotalRecords $totalRecords `
            -WeightedPct $weightedPct
    }

    $aeOnAllDead = {
        param($Tasks, $PendingCount, $Context)
        Write-ExportLog -Message "All AE workers dead - saving state for resume" -Level Error
        Write-AETaskCsv -Path $Context.TaskCsvPath -Tasks $Context.DayTasks
    }

    $aeOnIterComplete = {
        param($Tasks, $LoopState, $Context)
        Write-AETaskCsv -Path $Context.TaskCsvPath -Tasks $Context.DayTasks
    }

    # Build context with all state the callbacks need
    $aeContext = @{
        PageSize        = $PageSize
        DayTasks        = $dayTasks
        TaskCsvPath     = $aeTaskCsvPath
        ExportStartTime = $exportStartTime
        ExportDir       = $exportDir
    }

    $loopResult = Invoke-DispatchLoop `
        -ExportDir $exportDir `
        -Tasks $dayTasks `
        -WorkerProcesses $workerProcesses `
        -Context $aeContext `
        -OnScanCompletions $aeOnScan `
        -OnMatchTask $aeOnMatch `
        -OnDispatchTask $aeOnDispatch `
        -OnShowDashboard $aeOnDashboard `
        -OnAllWorkersDead $aeOnAllDead `
        -OnIterationComplete $aeOnIterComplete `
        -SleepSeconds 2

    # Save final task state
    Write-AETaskCsv -Path $aeTaskCsvPath -Tasks $dayTasks

    # Summary
    $completedTasks = @($dayTasks | Where-Object { $_.Status -eq "Completed" })
    $errorTasks = @($dayTasks | Where-Object { $_.Status -eq "Error" })

    if ($errorTasks.Count -gt 0) {
        Write-ExportLog -Message ("`n--- {0} day task(s) had errors ---" -f $errorTasks.Count) -Level Warning
        foreach ($et in $errorTasks) {
            Write-ExportLog -Message ("  Day {0}: {1}" -f $et.Day, $et.ErrorMessage) -Level Warning
        }
    }

    # Write manifest and complete
    Write-AEManifest -ExportDir $exportDir
    Write-ExportPhase -ExportDir $exportDir -Phase "AECompleted"

    $totalRecords = [long]0
    foreach ($t in $completedTasks) { $totalRecords += ($t.RecordCount -as [long]) }
    $exportElapsed = (Get-Date) - $exportStartTime

    Write-ExportLog -Message "`n========== Activity Explorer Multi-Terminal Summary ==========" -Level Info
    Write-ExportLog -Message ("Days: {0} completed, {1} errors, {2} total" -f $completedTasks.Count, $errorTasks.Count, $taskCount) -Level Info
    Write-ExportLog -Message ("Total records: {0:N0}" -f $totalRecords) -Level Info
    Write-ExportLog -Message ("Workers: {0}" -f $workerProcesses.Count) -Level Info
    Write-ExportLog -Message ("Duration: {0}" -f (Format-TimeSpan -Seconds $exportElapsed.TotalSeconds)) -Level Info
    Write-ExportLog -Message ("Output: per-day page files in Data/ActivityExplorer/") -Level Info
}

#endregion

function Invoke-ActivityExplorerExport {
    Write-ExportLog -Message "`n========== Activity Explorer Export ==========" -Level Info

    # Use UTC time and subtract a small buffer to avoid "future date" errors
    # This handles timezone edge cases where local time is ahead of UTC
    $endTime = [DateTime]::UtcNow.AddMinutes(-5)  # 5 minute buffer for safety
    $startTime = $endTime.AddDays(-$PastDays)

    Write-ExportLog -Message "Time range: $($startTime.ToString('yyyy-MM-dd HH:mm')) to $($endTime.ToString('yyyy-MM-dd HH:mm')) UTC" -Level Info
    Write-ExportLog -Message "Past days: $PastDays" -Level Info
    Write-ExportLog -Message "Local timezone: $([System.TimeZoneInfo]::Local.DisplayName)" -Level Info

    # Create ActivityExplorer subfolder for resilient export
    $aeOutputDir = Get-AEDataDir $script:ExportRunDirectory
    if (-not (Test-Path $aeOutputDir)) {
        New-Item -ItemType Directory -Force -Path $aeOutputDir | Out-Null
    }

    # Initialize run tracker
    $trackerPath = Join-Path (Get-CoordinationDir $script:ExportRunDirectory) "AE-RunTracker.json"
    $tracker = Get-ActivityExplorerRunTracker -TrackerPath $trackerPath

    # Progress log for tailing
    $progressLogPath = Join-Path (Get-LogsDir $script:ExportRunDirectory) "ActivityExplorer-Progress.log"
    Write-ExportLog -Message "Progress log (tail -f): $progressLogPath" -Level Info

    # Load configuration for filters
    $configPath = Join-Path $scriptRoot "ConfigFiles\ActivityExplorerSelector.json"

    if ($AEResume) {
        # On resume, prefer saved manifest for filter consistency
        $filters = Resolve-AEFilters -ExportRunDirectory $script:ExportRunDirectory -ConfigPath $configPath -LogDetails
    }
    else {
        # Read config once for both filter loading and manifest save
        $selectorConfig = Read-JsonConfig -Path $configPath
        $filters = Get-ActivityExplorerFilters -ConfigObject $selectorConfig -LogDetails

        # Save export settings manifest for resume consistency
        Save-ExportSettings -ExportRunDirectory $script:ExportRunDirectory -ExportType "ActivityExplorer" -Settings @{
            PastDays       = $PastDays
            PageSize       = $PageSize
            SelectorConfig = $selectorConfig
        }
    }

    try {
        # Use the new resilient export with per-page saving
        $exportParams = @{
            StartTime = $startTime
            EndTime = $endTime
            PageSize = $PageSize
            Filters = $filters
            OutputDirectory = $aeOutputDir
            Tracker = $tracker
            TrackerPath = $trackerPath
            ProgressLogPath = $progressLogPath
        }

        # Add Resume flag if specified
        if ($AEResume) {
            $exportParams['Resume'] = $true
            Write-ExportLog -Message "  Resume mode enabled - will continue from last successful page" -Level Info
        }

        $exportResult = Export-ActivityExplorerWithProgress @exportParams

        if ($exportResult.TotalRecords -eq 0) {
            Write-ExportLog -Message "  No activity records returned" -Level Warning
        }
        else {
            if ($exportResult.ResumedFrom) {
                Write-ExportLog -Message "  RESUMED from page $($exportResult.ResumedFrom.PageNumber) ($($exportResult.ResumedFrom.RecordCount) records)" -Level Info
            }
            Write-ExportLog -Message "  Total records exported: $($exportResult.TotalRecords) in $($exportResult.PageCount) pages" -Level Info
            Write-ExportLog -Message "  Data saved to: $(Get-AEDataDir $script:ExportRunDirectory)" -Level Success
        }
    }
    catch {
        Write-ExportLog -Message "  Activity Explorer export failed: $($_.Exception.Message)" -Level Error
        Write-ExportLog -Message "  Stack: $($_.ScriptStackTrace)" -Level Error
    }
}

function Invoke-eDiscoveryExport {
    Write-ExportLog -Message "`n========== eDiscovery Export ==========" -Level Info

    $exportResult = @{
        ExportTimestamp = Get-Date -Format "yyyy-MM-ddTHH:mm:ssZ"
        Cases = @()
        Searches = @()
    }

    try {
        $ediscovery = Export-eDiscoveryCases
        $exportResult.Cases = $ediscovery.Cases
        $exportResult.Searches = $ediscovery.Searches
    }
    catch {
        Write-ExportLog -Message "eDiscovery export failed: $($_.Exception.Message)" -Level Error
    }

    if ($OutputFormat -eq "JSON") {
        Save-ExportData -Data $exportResult -Name "eDiscovery-Config" -Format $OutputFormat -Directory $script:ExportRunDirectory
    }
    else {
        if (@($exportResult.Cases).Count -gt 0) {
            Save-ExportData -Data $exportResult.Cases -Name "eDiscovery-Cases" -Format $OutputFormat -Directory $script:ExportRunDirectory
        }
        if (@($exportResult.Searches).Count -gt 0) {
            Save-ExportData -Data $exportResult.Searches -Name "eDiscovery-Searches" -Format $OutputFormat -Directory $script:ExportRunDirectory
        }
    }
}

function Invoke-RbacExport {
    Write-ExportLog -Message "`n========== RBAC Export ==========" -Level Info

    $exportResult = @{
        ExportTimestamp = Get-Date -Format "yyyy-MM-ddTHH:mm:ssZ"
        RoleGroups = @()
        Members = @()
    }

    try {
        $rbac = Export-RbacConfiguration
        $exportResult.RoleGroups = $rbac.RoleGroups
        $exportResult.Members = $rbac.Members
    }
    catch {
        Write-ExportLog -Message "RBAC export failed: $($_.Exception.Message)" -Level Error
    }

    if ($OutputFormat -eq "JSON") {
        Save-ExportData -Data $exportResult -Name "RBAC-Config" -Format $OutputFormat -Directory $script:ExportRunDirectory
    }
    else {
        if (@($exportResult.RoleGroups).Count -gt 0) {
            Save-ExportData -Data $exportResult.RoleGroups -Name "RBAC-RoleGroups" -Format $OutputFormat -Directory $script:ExportRunDirectory
        }
        if (@($exportResult.Members).Count -gt 0) {
            Save-ExportData -Data $exportResult.Members -Name "RBAC-Members" -Format $OutputFormat -Directory $script:ExportRunDirectory
        }
    }
}

#endregion

#region Main Execution

# Worker mode: skip menu, confirmation, and go straight to work
if ($WorkerExportDir) {
    if (-not (Test-Path $WorkerExportDir)) {
        Write-Host ("`n  ERROR: Export directory not found: {0}" -f $WorkerExportDir) -ForegroundColor Red
        exit 1
    }

    # Detect export type (CE or AE) from ExportType.txt
    $workerExportType = Read-ExportType -ExportDir $WorkerExportDir
    if (-not $workerExportType) { $workerExportType = "ContentExplorer" }  # Backward compat

    # Display worker mode banner
    $folderName = Split-Path $WorkerExportDir -Leaf
    $bannerLabel = if ($workerExportType -eq 'ActivityExplorer') { "ACTIVITY EXPLORER - WORKER MODE" } else { "CONTENT EXPLORER - WORKER MODE" }
    Write-Host "`n  ╔════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host ("  ║  {0,-42}║" -f $bannerLabel) -ForegroundColor Cyan
    Write-Host "  ╠════════════════════════════════════════════╣" -ForegroundColor Cyan
    Write-Host ("  ║  Export: {0}" -f $folderName) -ForegroundColor Cyan
    Write-Host ("  ║  PID: {0}" -f $PID) -ForegroundColor Cyan
    Write-Host "  ╚════════════════════════════════════════════╝" -ForegroundColor Cyan

    # Check prerequisites
    if (-not (Test-ExportPrerequisites)) { exit 1 }

    # Authenticate
    $connectParams = Build-AuthParameters
    $script:AuthParams = $connectParams.Clone()
    $script:IsWorkerMode = $true
    if (-not (Connect-Compl8Compliance @connectParams)) {
        Write-Host "`n  ERROR: Authentication failed" -ForegroundColor Red
        exit 1
    }

    # Run appropriate worker based on export type
    try {
        if ($workerExportType -eq 'ActivityExplorer') {
            Invoke-ActivityExplorerWorker -WorkerExportDir $WorkerExportDir
        }
        else {
            Invoke-ContentExplorerWorker -WorkerExportDir $WorkerExportDir
        }
    }
    finally {
        Disconnect-Compl8Compliance
    }
    exit 0
}

# Resume mode: skip menu, authenticate, and resume export
if ($CEResumeDir) {
    if (-not (Test-Path $CEResumeDir)) {
        Write-Host ("`n  ERROR: Export directory not found: {0}" -f $CEResumeDir) -ForegroundColor Red
        exit 1
    }

    # Check for ExportPhase.txt to verify this is a valid export directory
    $resumePhase = Read-ExportPhase -ExportDir $CEResumeDir
    if (-not $resumePhase) {
        Write-Host "`n  ERROR: No ExportPhase.txt found - not a valid export directory" -ForegroundColor Red
        exit 1
    }

    if (-not (Test-ExportPrerequisites)) { exit 1 }

    $connectParams = Build-AuthParameters
    $script:AuthParams = $connectParams.Clone()
    $script:IsWorkerMode = $false
    if (-not (Connect-Compl8Compliance @connectParams)) {
        Write-Host "`n  ERROR: Authentication failed" -ForegroundColor Red
        exit 1
    }

    try {
        Invoke-ContentExplorerResume -ExportDir $CEResumeDir -WorkerCount $WorkerCount
    }
    finally {
        Disconnect-Compl8Compliance
    }
    exit 0
}

# Retry mode: skip menu, authenticate, and retry discrepant tasks
if ($CERetryDir) {
    if (-not (Test-Path $CERetryDir)) {
        Write-Host ("`n  ERROR: Export directory not found: {0}" -f $CERetryDir) -ForegroundColor Red
        exit 1
    }

    $retryTasksPath = Join-Path (Get-CoordinationDir $CERetryDir) "RetryTasks.csv"
    if (-not (Test-Path $retryTasksPath)) {
        Write-Host "`n  ERROR: No RetryTasks.csv found - no tasks to retry" -ForegroundColor Red
        exit 1
    }

    if (-not (Test-ExportPrerequisites)) { exit 1 }

    $connectParams = Build-AuthParameters
    $script:AuthParams = $connectParams.Clone()
    $script:IsWorkerMode = $false
    if (-not (Connect-Compl8Compliance @connectParams)) {
        Write-Host "`n  ERROR: Authentication failed" -ForegroundColor Red
        exit 1
    }

    try {
        Invoke-ContentExplorerRetry -ExportDir $CERetryDir
    }
    finally {
        Disconnect-Compl8Compliance
    }
    exit 0
}

# Task CSV mode: skip menu, authenticate, and run from task CSV
if ($CETasksCsv) {
    if (-not (Test-Path $CETasksCsv)) {
        Write-Host ("`n  ERROR: Task CSV not found: {0}" -f $CETasksCsv) -ForegroundColor Red
        exit 1
    }

    # Quick schema check
    $testTasks = @(Read-TaskCsv -Path $CETasksCsv)
    if ($testTasks.Count -eq 0) {
        Write-Host "`n  ERROR: No tasks found in CSV" -ForegroundColor Red
        exit 1
    }

    if (-not (Test-ExportPrerequisites)) { exit 1 }

    $connectParams = Build-AuthParameters
    $script:AuthParams = $connectParams.Clone()
    $script:IsWorkerMode = $false
    if (-not (Connect-Compl8Compliance @connectParams)) {
        Write-Host "`n  ERROR: Authentication failed" -ForegroundColor Red
        exit 1
    }

    try {
        Invoke-ContentExplorerFromTasksCsv -TasksCsvPath $CETasksCsv -WorkerCount $WorkerCount
    }
    finally {
        Disconnect-Compl8Compliance
    }
    exit 0
}

# Check if we should show the interactive menu
$noParamsResult = Test-NoParametersProvided
$showMenu = $Menu.IsPresent -or $noParamsResult

Write-Verbose "Menu check: Menu.IsPresent=$($Menu.IsPresent), NoParamsResult=$noParamsResult, showMenu=$showMenu"

# Variables to track menu selections (may override parameters)
$script:SelectedMode = $null
$script:MenuNoActivity = $false
$script:MenuNoContent = $false

if ($showMenu) {
    $menuResult = Show-ExportMenu

    if ($menuResult.Quit) {
        Write-Host "`nExport cancelled." -ForegroundColor Yellow
        exit 0
    }

    # Apply menu selections
    $script:MenuNoActivity = $menuResult.NoActivity
    $script:MenuNoContent = $menuResult.NoContent
    if ($menuResult.PastDays -and $menuResult.PastDays -ne 7) { $PastDays = $menuResult.PastDays }
    $CEAllSITs = $menuResult.CEAllSITs
    if ($menuResult.CEWorkloads) { $CEWorkloads = $menuResult.CEWorkloads }
    if ($menuResult.OutputFormat -ne "JSON") { $OutputFormat = $menuResult.OutputFormat }

    # Route new Content Explorer modes
    if ($menuResult.CEMultiTerminal) {
        $script:SelectedMode = "ContentExplorerMulti"
        $WorkerCount = $menuResult.CEWorkerCount
        $CEAllSITs = $menuResult.CEAllSITs
    }
    elseif ($menuResult.CEResumePath) {
        $script:SelectedMode = "ContentExplorerResume"
        $CEResumeDir = $menuResult.CEResumePath
        if ($menuResult.CEWorkerCount -and ($menuResult.CEWorkerCount -as [int]) -ge 2) {
            $WorkerCount = $menuResult.CEWorkerCount
        }
    }
    elseif ($menuResult.CERetryPath) {
        $script:SelectedMode = "ContentExplorerRetry"
        $CERetryDir = $menuResult.CERetryPath
    }
    elseif ($menuResult.CETasksCsvPath) {
        $script:SelectedMode = "ContentExplorerTasksCsv"
        $CETasksCsv = $menuResult.CETasksCsvPath
        if ($menuResult.CEWorkerCount -and ($menuResult.CEWorkerCount -as [int]) -ge 2) {
            $WorkerCount = $menuResult.CEWorkerCount
        }
    }
    elseif ($menuResult.AEMultiTerminal) {
        $script:SelectedMode = "ActivityExplorerMulti"
        $AEWorkerCount = $menuResult.AEWorkerCount
    }
    elseif ($menuResult.AEResumePath) {
        $script:SelectedMode = "ActivityExplorerResume"
        $AEResumeDir = $menuResult.AEResumePath
        if ($menuResult.AEWorkerCount -and ($menuResult.AEWorkerCount -as [int]) -ge 2) {
            $AEWorkerCount = $menuResult.AEWorkerCount
        }
    }
    else {
        $script:SelectedMode = $menuResult.Mode
    }

    Write-Host ""
}

# Merge NoActivity/NoContent from menu and command line
if ($script:MenuNoActivity) { $NoActivity = $true }
if ($script:MenuNoContent) { $NoContent = $true }

# Route CLI parameters to SelectedMode (menu already sets it; this handles CLI-only invocation)
# Note: -CEResumeDir, -CERetryDir, -CETasksCsv exit early (lines ~5602-5696) so never reach here
if (-not $script:SelectedMode) {
    if ($PSCmdlet.ParameterSetName -eq "ContentExplorer" -and $WorkerCount -ge 2) {
        $script:SelectedMode = "ContentExplorerMulti"
    }
    elseif ($PSCmdlet.ParameterSetName -eq "ActivityExplorer" -and $AEResumeDir) {
        $script:SelectedMode = "ActivityExplorerResume"
    }
    elseif ($PSCmdlet.ParameterSetName -eq "ActivityExplorer" -and $AEWorkerCount -ge 2) {
        $script:SelectedMode = "ActivityExplorerMulti"
    }
    else {
        # Base CLI modes: Full, DLP, Labels, ContentExplorer, ActivityExplorer, eDiscovery, RBAC
        $script:SelectedMode = $PSCmdlet.ParameterSetName
    }
}

# Show export plan BEFORE connecting (using Write-Host since logging not yet initialized)
Write-Banner -Title 'Compl8 Cmdlet Export Tool'

# Determine export mode display string
$exportMode = switch ($script:SelectedMode) {
    "Full" { "Full Export" }
    "DLP" { "DLP Only" }
    "Labels" { "Labels Only" }
    "ContentExplorer" { "Content Explorer" }
    "ContentExplorerMulti" { "Content Explorer (Multi-Terminal)" }
    "ContentExplorerResume" { "Content Explorer (Resume)" }
    "ContentExplorerRetry" { "Content Explorer (Retry)" }
    "ContentExplorerTasksCsv" { "Content Explorer (Task CSV)" }
    "ActivityExplorer" { "Activity Explorer" }
    "ActivityExplorerMulti" { "Activity Explorer (Multi-Terminal)" }
    "ActivityExplorerResume" { "Activity Explorer (Resume)" }
    "eDiscovery" { "eDiscovery Only" }
    "RBAC" { "RBAC Only" }
    default { "Full Export" }
}

Write-Host "`nExport Plan:" -ForegroundColor Yellow
Write-Host "  Mode:             $exportMode"
Write-Host "  Format:           $OutputFormat"
Write-Host "  Output Directory: $script:ExportRunDirectory"

# Show what will be exported based on mode and config files
Write-Host "`nData to be exported:" -ForegroundColor Yellow

switch ($script:SelectedMode) {
    "DLP" {
        Write-Host "  - DLP Policies and Rules"
        Write-Host "  - Sensitive Information Types"
    }
    "Labels" {
        Write-Host "  - Sensitivity Labels and Policies"
        Write-Host "  - Retention Labels and Policies"
    }
    "ContentExplorer" {
        # Read config to show what's enabled
        $ceConfigPath = Join-Path $scriptRoot "ConfigFiles\ContentExplorerClassifiers.json"
        if (Test-Path $ceConfigPath) {
            $ceConfig = Get-Content -Raw $ceConfigPath | ConvertFrom-Json
            $enabledTagTypes = @()
            $enabledWorkloads = @()
            if ($ceConfig.TagTypes) {
                foreach ($prop in $ceConfig.TagTypes.PSObject.Properties) {
                    if ($prop.Value -eq "True") { $enabledTagTypes += $prop.Name }
                }
            }
            if ($ceConfig.Workloads) {
                foreach ($prop in $ceConfig.Workloads.PSObject.Properties) {
                    if ($prop.Value -eq "True") { $enabledWorkloads += $prop.Name }
                }
            }
            Write-Host "  Tag Types: $($enabledTagTypes -join ', ')"
            Write-Host "  Workloads: $($enabledWorkloads -join ', ')"
        }
        else {
            Write-Host "  Tag Types: Sensitivity, Retention, SensitiveInformationType (defaults)"
            Write-Host "  Workloads: Exchange, SharePoint, OneDrive, Teams (defaults)"
        }
        Write-Host "  (Multiple files will be created per tag/workload combination)"
    }
    "ContentExplorerMulti" {
        Write-Host "  - Content Explorer data (multi-terminal parallel export)"
        Write-Host "  Workers: $WorkerCount"
        if ($CEAllSITs) {
            Write-Host "  Scope: ALL Sensitive Information Types (full tenant scan)"
        } else {
            Write-Host "  Scope: From config file"
        }
        Write-Host "  (Workers coordinate via file-drop protocol)"
    }
    "ContentExplorerResume" {
        Write-Host "  - Content Explorer data (resuming previous export)"
        Write-Host ("  Export Dir: {0}" -f $CEResumeDir)
        if ($WorkerCount -gt 0) {
            Write-Host ("  Workers: {0} (multi-terminal resume)" -f $WorkerCount)
        }
        else {
            Write-Host "  Workers: Single terminal"
        }
    }
    "ContentExplorerRetry" {
        Write-Host "  - Content Explorer data (retrying discrepant tasks)"
        Write-Host ("  Export Dir: {0}" -f $CERetryDir)
    }
    "ContentExplorerTasksCsv" {
        Write-Host "  - Content Explorer data (from task CSV)"
        Write-Host ("  Task CSV: {0}" -f $CETasksCsv)
        if ($WorkerCount -gt 0) {
            Write-Host ("  Workers: {0} (multi-terminal)" -f $WorkerCount)
        }
        else {
            Write-Host "  Workers: Single terminal"
        }
    }
    "ActivityExplorer" {
        # Read config to show what's enabled
        $aeConfigPath = Join-Path $scriptRoot "ConfigFiles\ActivityExplorerSelector.json"
        if (Test-Path $aeConfigPath) {
            $aeConfig = Get-Content -Raw $aeConfigPath | ConvertFrom-Json
            $enabledActivities = @()
            $enabledWorkloads = @()
            if ($aeConfig.Activities) {
                foreach ($prop in $aeConfig.Activities.PSObject.Properties) {
                    if ($prop.Value -eq "True") { $enabledActivities += $prop.Name }
                }
            }
            if ($aeConfig.Workloads) {
                foreach ($prop in $aeConfig.Workloads.PSObject.Properties) {
                    if ($prop.Value -eq "True") { $enabledWorkloads += $prop.Name }
                }
            }
            Write-Host "  Activities: $($enabledActivities -join ', ')"
            Write-Host "  Workloads:  $($enabledWorkloads -join ', ')"
        }
        else {
            Write-Host "  Activities: All (no filter)"
            Write-Host "  Workloads:  All (no filter)"
        }
        Write-Host "  Time Range: Last $PastDays days"
    }
    "ActivityExplorerMulti" {
        Write-Host "  - Activity Explorer data (multi-terminal parallel export)"
        Write-Host "  Workers: $AEWorkerCount"
        Write-Host "  Time Range: Last $PastDays days (split into per-day tasks)"
        Write-Host "  (Workers coordinate via file-drop protocol)"
    }
    "ActivityExplorerResume" {
        Write-Host "  - Activity Explorer data (resuming previous export)"
        Write-Host ("  Export Dir: {0}" -f $AEResumeDir)
        if ($AEWorkerCount -gt 0) {
            Write-Host ("  Workers: {0} (multi-terminal resume)" -f $AEWorkerCount)
        }
        else {
            Write-Host "  Workers: Single terminal"
        }
    }
    "eDiscovery" {
        Write-Host "  - Compliance Cases"
        Write-Host "  - Compliance Searches"
    }
    "RBAC" {
        Write-Host "  - Role Groups"
        Write-Host "  - Role Group Members"
    }
    default {
        Write-Host "  - DLP Policies, Rules, and Sensitive Information Types"
        Write-Host "  - Sensitivity Labels and Policies"
        Write-Host "  - Retention Labels and Policies"
        Write-Host "  - eDiscovery Cases and Searches"
        Write-Host "  - RBAC Role Groups and Members"
        if ($NoContent) {
            Write-Host "  - Content Explorer data " -NoNewline
            Write-Host "(SKIPPED - -NoContent)" -ForegroundColor Yellow
        }
        else {
            Write-Host "  - Content Explorer data (per config file)"
        }
        if ($NoActivity) {
            Write-Host "  - Activity Explorer data " -NoNewline
            Write-Host "(SKIPPED - -NoActivity)" -ForegroundColor Yellow
        }
        else {
            Write-Host "  - Activity Explorer data (last $PastDays days)"
        }
    }
}

# Show output files (timestamp is in the folder name, not file names)
$ext = if ($OutputFormat -eq "JSON") { "json" } else { "csv" }
Write-Host "`nOutput files (in $script:ExportRunDirectory):" -ForegroundColor Yellow
switch ($script:SelectedMode) {
    "DLP" {
        if ($OutputFormat -eq "JSON") {
            Write-Host "  - DLP-Config.$ext"
        }
        else {
            Write-Host "  - DLP-Policies.$ext"
            Write-Host "  - DLP-Rules.$ext"
            Write-Host "  - SensitiveInfoTypes.$ext"
        }
    }
    "Labels" {
        if ($OutputFormat -eq "JSON") {
            Write-Host "  - Labels-Config.$ext"
        }
        else {
            Write-Host "  - SensitivityLabels.$ext"
            Write-Host "  - LabelPolicies.$ext"
            Write-Host "  - RetentionLabels.$ext"
            Write-Host "  - RetentionPolicies.$ext"
        }
    }
    "ContentExplorer" {
        Write-Host "  - Data/ContentExplorer/            (per-classifier page files)"
        Write-Host "  - _Coordination/                   (phase, task tracking)"
    }
    "ContentExplorerMulti" {
        Write-Host "  - Data/ContentExplorer/            (per-classifier page files)"
        Write-Host "  - _Coordination/                   (phase, task tracking, worker coordination)"
    }
    "ContentExplorerResume" {
        Write-Host "  - (Resuming into existing export directory)"
        Write-Host "  - Data/ContentExplorer/            (per-classifier page files)"
        Write-Host "  - _Coordination/                   (phase, task tracking)"
    }
    "ContentExplorerTasksCsv" {
        Write-Host "  - Data/ContentExplorer/            (per-task page files)"
        Write-Host "  - _Coordination/DetailTasks.csv    (task tracking)"
    }
    "ActivityExplorer" {
        Write-Host "  - Data/ActivityExplorer/           (per-day page files)"
        Write-Host "    - YYYY-MM-DD/Page-001.json, ... (per-page data)"
        Write-Host "    - RunTracker.json               (state/progress tracking)"
        Write-Host "    - Progress.log                  (tailable progress log)"
    }
    "ActivityExplorerMulti" {
        Write-Host "  - Data/ActivityExplorer/           (per-day page files)"
        Write-Host "    - YYYY-MM-DD/Page-001.json, ... (per-page data)"
        Write-Host "  - _Coordination/AEDayTasks.csv    (day task tracking)"
        Write-Host "  - _Coordination/                  (phase, type, worker coordination)"
    }
    "ActivityExplorerResume" {
        Write-Host "  - (Resuming into existing export directory)"
        Write-Host "  - Data/ActivityExplorer/           (per-day page files)"
    }
    "eDiscovery" {
        if ($OutputFormat -eq "JSON") {
            Write-Host "  - eDiscovery-Config.$ext"
        }
        else {
            Write-Host "  - eDiscovery-Cases.$ext"
            Write-Host "  - eDiscovery-Searches.$ext"
        }
    }
    "RBAC" {
        if ($OutputFormat -eq "JSON") {
            Write-Host "  - RBAC-Config.$ext"
        }
        else {
            Write-Host "  - RBAC-RoleGroups.$ext"
            Write-Host "  - RBAC-Members.$ext"
        }
    }
    default {
        Write-Host "  - DLP-Config.$ext"
        Write-Host "  - SensitivityLabels-Config.$ext"
        Write-Host "  - RetentionLabels-Config.$ext"
        Write-Host "  - eDiscovery-Config.$ext"
        Write-Host "  - RBAC-Config.$ext"
        if (-not $NoContent) {
            Write-Host "  - Data/ContentExplorer/            (per-classifier page files)"
        }
        if (-not $NoActivity) {
            Write-Host "  - Data/ActivityExplorer/           (per-day page files)"
        }
    }
}

Write-Host ""

# Confirmation prompt (defaults to Yes on Enter)
$confirmation = Read-Host "Proceed with export? [Y]/N"
if ($confirmation -match '^[Nn]') {
    Write-Host "Export cancelled by user." -ForegroundColor Yellow
    exit 0
}

Write-Host ""

# Now start the actual export process
try {
    Write-ExportLog -Message "=================================================================================" -Level Info
    Write-ExportLog -Message "Starting Export..." -Level Info
    Write-ExportLog -Message "=================================================================================" -Level Info
    Write-ExportLog -Message "Mode: $exportMode" -Level Info
    Write-ExportLog -Message "Format: $OutputFormat" -Level Info
    Write-ExportLog -Message "Output Directory: $script:ExportRunDirectory" -Level Info
    Write-ExportLog -Message "Log File: $logFile" -Level Info

    # Check prerequisites
    Write-ExportLog -Message "`nChecking prerequisites..." -Level Info
    if (-not (Test-ExportPrerequisites)) {
        Write-ExportLog -Message "Prerequisites not met. Please install required modules." -Level Error
        exit 1
    }

    # Validate configuration files BEFORE login
    # Map extended modes to base modes for config validation
    $validationMode = switch ($script:SelectedMode) {
        "ContentExplorerMulti"     { "ContentExplorer" }
        "ContentExplorerResume"    { $null }  # Skip validation for resume (project already exists)
        "ContentExplorerRetry"     { $null }  # Skip validation for retry (project already exists)
        "ContentExplorerTasksCsv"  { $null }  # Skip validation for task CSV (tasks already defined)
        "ActivityExplorerMulti"    { "ActivityExplorer" }
        "ActivityExplorerResume"   { $null }  # Skip validation for resume (project already exists)
        default { $script:SelectedMode }
    }

    if ($validationMode) {
        Write-ExportLog -Message "`nValidating configuration files..." -Level Info
        $configParams = @{
            ExportMode = $validationMode
            ScriptRoot = $scriptRoot
        }
        if ($NoContent) { $configParams['NoContent'] = $true }
        if ($NoActivity) { $configParams['NoActivity'] = $true }

        if (-not (Test-ExportConfiguration @configParams)) {
            Write-ExportLog -Message "Configuration validation failed. Fix errors above before connecting." -Level Error
            exit 1
        }
    } else {
        Write-ExportLog -Message "`nSkipping config validation (resume mode - project already exists)" -Level Info
    }

    # Connect to Security & Compliance PowerShell
    Write-ExportLog -Message "`nConnecting to Security & Compliance PowerShell..." -Level Info
    $connectParams = Build-AuthParameters
    $script:AuthParams = $connectParams.Clone()
    $script:IsWorkerMode = [bool]$WorkerExportDir

    if (-not (Connect-Compl8Compliance @connectParams)) {
        Write-ExportLog -Message "Failed to connect. Exiting." -Level Error
        exit 1
    }

    # Determine export type and execute
    switch ($script:SelectedMode) {
        "DLP" { Invoke-DlpExport }
        "Labels" { Invoke-LabelsExport }
        "ContentExplorer" { Invoke-ContentExplorerExport }
        "ContentExplorerMulti" { Invoke-ContentExplorerExport }
        "ContentExplorerResume" { Invoke-ContentExplorerResume -ExportDir $CEResumeDir -WorkerCount $WorkerCount }
        "ContentExplorerRetry" { Invoke-ContentExplorerRetry -ExportDir $CERetryDir }
        "ContentExplorerTasksCsv" { Invoke-ContentExplorerFromTasksCsv -TasksCsvPath $CETasksCsv -WorkerCount $WorkerCount }
        "ActivityExplorer" { Invoke-ActivityExplorerExport }
        "ActivityExplorerMulti" { Invoke-AEMultiExport }
        "ActivityExplorerResume" { Invoke-AEMultiExport -IsResume -ResumeDir $AEResumeDir -ResumeWorkerCount $AEWorkerCount }
        "eDiscovery" { Invoke-eDiscoveryExport }
        "RBAC" { Invoke-RbacExport }
        default { Invoke-FullExport }
    }

    # Show summary
    $stats = Get-ExportStatistics
    Write-ExportLog -Message "`n=================================================================================" -Level Info
    Write-ExportLog -Message "Export Complete" -Level Success
    Write-ExportLog -Message "=================================================================================" -Level Info
    Write-ExportLog -Message "Duration: $($stats.Duration)" -Level Info
    Write-ExportLog -Message "Items exported:" -Level Info

    foreach ($key in $stats.ItemsExported.Keys) {
        Write-ExportLog -Message "  - ${key}: $($stats.ItemsExported[$key])" -Level Info
    }

    if ($stats.ErrorCount -gt 0) {
        Write-ExportLog -Message "Errors: $($stats.ErrorCount)" -Level Warning
    }
    if ($stats.WarningCount -gt 0) {
        Write-ExportLog -Message "Warnings: $($stats.WarningCount)" -Level Warning
    }

    Write-ExportLog -Message "`nOutput directory: $script:ExportRunDirectory" -Level Info
    Write-ExportLog -Message "Log file: $logFile" -Level Info

    # Post-export: unified Parquet conversion
    if ($UnifiedParquet) {
        $parquetScript = Join-Path $scriptRoot "build_unified_parquet.py"
        if (-not (Test-Path $parquetScript)) {
            Write-ExportLog -Message "Unified Parquet script not found: $parquetScript" -Level Error
        }
        else {
            $parquetOutputDir = if ($UnifiedParquetDir) { $UnifiedParquetDir } else { "C:\PurviewData" }
            Write-ExportLog -Message "`nConverting to unified Parquet format..." -Level Info
            Write-ExportLog -Message "  Output: $parquetOutputDir" -Level Info

            try {
                $pyArgs = @("--input-dir", $script:ExportRunDirectory, "--output-dir", $parquetOutputDir)
                foreach ($csvPath in $UsersCsv) {
                    $pyArgs += @("--users-csv", $csvPath)
                }
                $pyResult = & python $parquetScript @pyArgs 2>&1
                $pyExitCode = $LASTEXITCODE
                foreach ($line in $pyResult) { Write-ExportLog -Message "  [parquet] $line" -Level Info }

                if ($pyExitCode -eq 0) {
                    Write-ExportLog -Message "Unified Parquet export complete." -Level Success
                }
                else {
                    Write-ExportLog -Message "Unified Parquet export failed (exit code $pyExitCode)" -Level Error
                }
            }
            catch {
                Write-ExportLog -Message "Failed to run Parquet converter: $($_.Exception.Message)" -Level Error
            }
        }
    }
}
catch {
    Write-ExportLog -Message "Fatal error: $($_.Exception.Message)" -Level Error
    Write-ExportLog -Message $_.ScriptStackTrace -Level Error
    exit 1
}
finally {
    Disconnect-Compl8Compliance
}

#endregion
