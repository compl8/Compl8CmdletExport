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
    # Fallback H: unattended mode with no recognised export mode → fail fast.
    if ($script:Unattended) {
        Write-ExportLog -Message "Unattended mode requires an explicit export mode (-FullExport, -DlpOnly, -ContentExplorer, etc.). No mode was provided; cannot show interactive menu." -Level Error
        if ($script:ExportRunDirectory -and (Test-Path $script:ExportRunDirectory)) {
            Write-RunSummary -ExportDir $script:ExportRunDirectory -Result @{
                Mode   = 'None'
                Status = 'ConfigError'
                Errors = @("No export mode specified for unattended run.")
            }
        }
        exit (Get-ExportExitCode -Status 'ConfigError')
    }

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
if ($script:TenantPrefix) {
    $prefixSource = if ($prefixResult) { $prefixResult.Source } else { "previous" }
    Write-Host ("  Tenant:           {0} (from {1})" -f $script:TenantPrefix, $prefixSource)
}
Write-Host "  Output Directory: $script:ExportRunDirectory"
if ($UnifiedParquet) {
    $plannedUnifiedParquetDir = Resolve-UnifiedParquetOutputDir -ConfiguredPath $UnifiedParquetDir -ExportRunDirectory $script:ExportRunDirectory
    Write-Host "  C8 Tuning Input:  $plannedUnifiedParquetDir"
}
if ($PowerBIParquet) {
    $plannedPowerBIParquetDir = Resolve-PowerBIParquetOutputDir -ConfiguredPath $PowerBIParquetDir -ExportRunDirectory $script:ExportRunDirectory
    Write-Host "  PBI Star Parquet: $plannedPowerBIParquetDir"
}

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
        } elseif ($CEAllTCs) {
            Write-Host "  Scope: ALL Trainable Classifiers (full tenant scan)"
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
if (-not $script:Unattended) {
    $confirmation = Read-Host "Proceed with export? [Y]/N"
    if ($confirmation -match '^[Nn]') {
        Write-Host "Export cancelled by user." -ForegroundColor Yellow
        exit 0
    }
} else {
    Write-ExportLog -Message "Unattended: proceeding without confirmation (prompt A skipped)." -Level Info
}

Write-Host ""

function Rename-ExportRunDirectory {
    <#
    .SYNOPSIS
        Renames the active export run directory to carry a new tenant-prefix label and
        updates the script-scope path variables that depend on it.
    .DESCRIPTION
        Factored out of Confirm-ConnectedTenant so the rename logic can be unit-tested
        without a live connection. The new directory name is
        "Export-{NewLabel}-{$script:ExportTimestamp}". On success, updates
        $script:ExportRunDirectory, $script:TenantPrefix, $script:LogFile and
        $script:ErrorLogPath to point at the renamed directory.

        The log is written via Add-Content (per-write open/close, no persistent handle),
        so renaming the directory between writes is safe.
    .OUTPUTS
        $true if the directory was renamed and the script vars updated; $false on failure
        (in which case the old directory and label are left untouched).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$NewLabel
    )

    $oldDir = $script:ExportRunDirectory
    $parent = Split-Path $oldDir -Parent
    $newName = "Export-{0}-{1}" -f $NewLabel, $script:ExportTimestamp
    $newPath = Join-Path $parent $newName

    if ($newPath -eq $oldDir) {
        # Same label as current; nothing to rename, but treat as success.
        return $true
    }

    try {
        Rename-Item -Path $oldDir -NewName $newName -ErrorAction Stop
    }
    catch {
        Write-Host ("  Could not rename export directory: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
        Write-Host "  Keeping the existing directory and label." -ForegroundColor Yellow
        return $false
    }

    # Repoint all script-scope path vars at the renamed directory.
    $logLeaf = if ($script:LogFile) { Split-Path $script:LogFile -Leaf } else { $null }
    $script:ExportRunDirectory = $newPath
    $script:TenantPrefix = $NewLabel
    if ($logLeaf) {
        $script:LogFile = Join-Path (Get-LogsDir $newPath) $logLeaf
    }
    $script:ErrorLogPath = Join-Path (Get-LogsDir $newPath) "ExportProject-Errors.log"

    return $true
}

function Confirm-ConnectedTenant {
    <#
    .SYNOPSIS
        Post-connect interactive guard: confirms the operator is on the intended tenant
        before any export work begins.
    .DESCRIPTION
        WAM/MSAL caches credentials, so Connect-IPPSSession can silently reuse a cached
        account - the export label (TenantPrefix) might say one tenant while the live
        session is actually connected to another. This shows BOTH the label and the
        actual connected domain (from Get-Compl8TenantInfo) and lets the operator
        Proceed, Relabel (rename the run dir), Re-authenticate (different account), or
        Abort.

        Non-interactive (input redirected / no console) or -Unattended
        ($script:Unattended): logs the label + domain and proceeds without prompting -
        keeps scheduled / cert-auth silent runs from hanging at this guard.
    .PARAMETER ConnectParams
        The splatted connect parameters (from Build-AuthParameters), reused for the
        Re-authenticate path.
    .OUTPUTS
        $true to proceed with the export; $false to abort.
    #>
    param(
        [Parameter(Mandatory)][hashtable]$ConnectParams
    )

    $reAuthed = $false

    while ($true) {
        # --- 1. Determine the actual connected domain (never hard-fail) ---
        $connectedDomain = $null
        try {
            $tenantInfo = Get-Compl8TenantInfo
            if ($tenantInfo -and $tenantInfo.TenantDomain) {
                $connectedDomain = [string]$tenantInfo.TenantDomain
            }
            elseif ($tenantInfo -and $tenantInfo.TenantId) {
                $connectedDomain = "TenantId: " + [string]$tenantInfo.TenantId
            }
        }
        catch {
            $connectedDomain = $null
        }
        $domainDisplay = if ($connectedDomain) { $connectedDomain } else { "(could not determine - Get-Compl8TenantInfo failed)" }

        $label = if ($script:TenantPrefix) { [string]$script:TenantPrefix } else { "(none)" }

        # --- 2. Display ---
        Write-Host ""
        Write-Host "--- Tenant confirmation ---" -ForegroundColor Cyan
        Write-Host ("  Export label (prefix): {0}" -f $label)
        Write-Host ("  Connected domain:      {0}" -f $domainDisplay)

        $labelLower = if ($script:TenantPrefix) { ([string]$script:TenantPrefix).ToLower() } else { "" }
        $domainLower = if ($connectedDomain) { $connectedDomain.ToLower() } else { "" }
        if ($labelLower -and ($domainLower -notlike "*$labelLower*")) {
            Write-Host "  WARNING: Label does not obviously match the connected domain - if you're on cached credentials you may be signed into the wrong tenant." -ForegroundColor Yellow
        }
        if ($reAuthed -and $connectedDomain -and $connectedDomain -eq $script:LastConfirmedDomain) {
            Write-Host "  (domain unchanged - if you intended a different account, you may need to sign out of the cached account in the browser)" -ForegroundColor Yellow
        }

        # --- 3. Non-interactive / unattended guard ---
        # -Unattended must skip this prompt even on an interactive console (a scheduled
        # task can launch with IsInputRedirected = $false), otherwise it would hang here.
        if ([Console]::IsInputRedirected -or $script:Unattended) {
            $skipReason = if ($script:Unattended) { "unattended/non-interactive" } else { "non-interactive" }
            Write-ExportLog -Message ("Tenant confirmation skipped ({0}): label='{1}', connected domain='{2}'. Proceeding." -f $skipReason, $label, $domainDisplay) -Level Info
            return $true
        }

        # --- 4. Interactive prompt ---
        Write-Host ""
        Write-Host "  [Y] Proceed   [L] Relabel   [R] Re-authenticate (different account)   [N] Abort"
        $choice = $null
        try {
            $choice = Read-Host "  Choice [Y]"
        }
        catch {
            # No usable console after all - proceed rather than block.
            Write-ExportLog -Message ("Tenant confirmation prompt unavailable (console read failed): label='{0}', connected domain='{1}'. Proceeding." -f $label, $domainDisplay) -Level Info
            return $true
        }

        $choice = ([string]$choice).Trim()
        $script:LastConfirmedDomain = $connectedDomain

        if ($choice -eq '' -or $choice -match '^[Yy]$') {
            Write-ExportLog -Message ("Tenant confirmed by operator: label='{0}', connected domain='{1}'." -f $label, $domainDisplay) -Level Info
            return $true
        }
        elseif ($choice -match '^[Ll]$') {
            # Relabel: sanitize the new label, rename the run dir, repoint vars, re-loop.
            $rawLabel = $null
            try { $rawLabel = Read-Host "  New label" } catch { $rawLabel = "" }
            $newLabel = ConvertTo-SafeTenantPrefix -Value ([string]$rawLabel)
            if (-not $newLabel) {
                Write-Host "  Label sanitized to empty (allowed: a-z 0-9 _ -). Please try again." -ForegroundColor Yellow
                continue
            }
            $oldDir = $script:ExportRunDirectory
            if (Rename-ExportRunDirectory -NewLabel $newLabel) {
                Write-ExportLog -Message ("Export relabeled '{0}' -> directory renamed: '{1}' -> '{2}'." -f $newLabel, $oldDir, $script:ExportRunDirectory) -Level Info
            }
            continue
        }
        elseif ($choice -match '^[Rr]$') {
            # Re-authenticate as a (potentially) different account.
            Write-Host "  Re-authenticating - disconnecting current session..." -ForegroundColor Cyan
            try { Disconnect-Compl8Compliance } catch { }
            if (-not (Connect-Compl8Compliance @ConnectParams)) {
                Write-Host "  Re-authentication failed. You can try again or Abort." -ForegroundColor Yellow
                continue
            }
            $reAuthed = $true
            continue
        }
        elseif ($choice -match '^[Nn]$') {
            return $false
        }
        else {
            Write-Host "  Unrecognized choice. Enter Y, L, R, or N." -ForegroundColor Yellow
            continue
        }
    }
}

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

    # Post-connect tenant confirmation (interactive only; non-interactive proceeds).
    # WAM/MSAL can silently reuse a cached account, so confirm label vs. connected
    # domain before any export work spawns workers / writes data.
    if (-not (Confirm-ConnectedTenant -ConnectParams $connectParams)) {
        Write-ExportLog -Message "Export cancelled at tenant confirmation." -Level Warning
        Disconnect-Compl8Compliance
        exit 0
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
            $parquetOutputDir = Resolve-UnifiedParquetOutputDir -ConfiguredPath $UnifiedParquetDir -ExportRunDirectory $script:ExportRunDirectory
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
                    Write-ExportLog -Message "C8 tuning input root: $parquetOutputDir" -Level Info
                    Write-ExportLog -Message "  content_files:  $(Join-Path $parquetOutputDir 'content\content_files')" -Level Info
                    Write-ExportLog -Message "  sit_detections: $(Join-Path $parquetOutputDir 'content\sit_detections')" -Level Info
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

    # Post-export: Power BI star-schema (v6) Parquet conversion — Activity Explorer data only
    if ($PowerBIParquet) {
        $aeDataDir = Join-Path $script:ExportRunDirectory "Data\ActivityExplorer"
        if (-not (Test-Path $aeDataDir)) {
            Write-ExportLog -Message "Skipping Power BI star-schema conversion: no Activity Explorer data at $aeDataDir (-PowerBIParquet applies to AE exports only)" -Level Warning
        }
        else {
            $starOutputDir = Resolve-PowerBIParquetOutputDir -ConfiguredPath $PowerBIParquetDir -ExportRunDirectory $script:ExportRunDirectory
            Write-ExportLog -Message "`nConverting to Power BI star-schema (v6) Parquet model..." -Level Info
            Write-ExportLog -Message "  Output: $starOutputDir" -Level Info

            $starArgs = @("-m", "parquet_builder.star.convert", "--input-dir", $script:ExportRunDirectory)
            if (-not [string]::IsNullOrWhiteSpace($PowerBIParquetDir)) {
                $starArgs += @("--output-dir", $PowerBIParquetDir)
            }

            # Enrichment inputs: ConfigFiles/AEStarEnrichment.local.json (gitignored).
            # Present -> pass the configured paths (hard-fail enrichment policy applies).
            # Absent  -> build an unenriched model with a prominent warning.
            # Malformed -> skip the conversion entirely rather than silently degrading.
            $skipStarConversion = $false
            $enrichmentConfigPath = Join-Path $scriptRoot "ConfigFiles\AEStarEnrichment.local.json"
            if (Test-Path $enrichmentConfigPath) {
                try {
                    $enrichmentConfig = Get-Content -Path $enrichmentConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
                    if (-not [string]::IsNullOrWhiteSpace($enrichmentConfig.RiskWorkbookPath)) {
                        $starArgs += @("--risk-workbook", $enrichmentConfig.RiskWorkbookPath)
                    }
                    if (-not [string]::IsNullOrWhiteSpace($enrichmentConfig.DepartmentCsvPath)) {
                        $starArgs += @("--department-csv", $enrichmentConfig.DepartmentCsvPath)
                    }
                    Write-ExportLog -Message "  Enrichment config: $enrichmentConfigPath" -Level Info
                }
                catch {
                    Write-ExportLog -Message "Failed to read enrichment config ${enrichmentConfigPath}: $($_.Exception.Message)" -Level Error
                    Write-ExportLog -Message "Skipping star-schema conversion. Fix or remove the config, then run: py -m parquet_builder.star.convert --input-dir `"$script:ExportRunDirectory`"" -Level Error
                    $skipStarConversion = $true
                }
            }
            else {
                $starArgs += "--allow-unenriched"
                Write-ExportLog -Message "  WARNING: No enrichment config found ($enrichmentConfigPath)" -Level Warning
                Write-ExportLog -Message "  WARNING: Building an UNENRICHED model - risk scores will be 0 and departments unmapped." -Level Warning
                Write-ExportLog -Message "  WARNING: Copy ConfigFiles\AEStarEnrichment.example.json to AEStarEnrichment.local.json to enable enrichment." -Level Warning
            }

            if (-not $skipStarConversion) {
                try {
                    # py launcher preferred (matches Build-PowerBI.ps1); python fallback
                    $pythonLauncher = if (Get-Command py -ErrorAction SilentlyContinue) { "py" } else { "python" }
                    # -m needs the repo root on sys.path
                    Push-Location $scriptRoot
                    try {
                        $starResult = & $pythonLauncher @starArgs 2>&1
                        $starExitCode = $LASTEXITCODE
                    }
                    finally {
                        Pop-Location
                    }
                    foreach ($line in $starResult) { Write-ExportLog -Message "  [star] $line" -Level Info }

                    if ($starExitCode -eq 0) {
                        Write-ExportLog -Message "Power BI star-schema Parquet export complete." -Level Success
                        Write-ExportLog -Message "Star model root: $starOutputDir" -Level Info
                        Write-ExportLog -Message "  Build the report: .\PowerBI\Build-PowerBI.ps1 -Project ActivityExplorer -ParquetRoot `"$starOutputDir`"" -Level Info
                    }
                    else {
                        Write-ExportLog -Message "Power BI star-schema Parquet export failed (exit code $starExitCode)" -Level Error
                    }
                }
                catch {
                    Write-ExportLog -Message "Failed to run star-schema converter: $($_.Exception.Message)" -Level Error
                }
            }
        }
    }

    # ── Compute final status and write RunSummary.json ─────────────────────
    # Status determination:
    #   Partial  — errors were logged, or remaining tasks > 0 (Detail phase incomplete)
    #   Completed — no errors, no remaining tasks
    # Exit code: 0 for Completed, 2 for Partial (so interactive users are not
    # surprised on a clean run, and schedulers can distinguish partial vs. clean).
    $finalStats          = Get-ExportStatistics
    $remainingTasksCount  = 0
    $remainingCheckFailed = $false
    if ($script:ExportRunDirectory -and (Test-Path $script:ExportRunDirectory)) {
        try {
            $remainingTasksCount = Write-RemainingTasksCsv -ExportDir $script:ExportRunDirectory
        }
        catch {
            # Can't trust the remaining-task count — never report a false Completed.
            # Force Partial so the exit code (2) reflects the uncertainty.
            $remainingCheckFailed = $true
            Write-ExportLog -Message ("Could not determine remaining-task count ({0}); reporting status as Partial." -f $_.Exception.Message) -Level Warning
        }
    }

    $finalStatus = if ($finalStats.ErrorCount -gt 0 -or $remainingTasksCount -gt 0 -or $remainingCheckFailed) {
        'Partial'
    } else {
        'Completed'
    }

    # Build sections array from ExportStats.ItemsExported categories.
    $finalSections = @(
        $finalStats.ItemsExported.Keys | ForEach-Object {
            @{
                Name        = $_
                Status      = 'Completed'
                RecordCount = [int]$finalStats.ItemsExported[$_]
                ErrorCount  = 0
            }
        }
    )

    if ($script:ExportRunDirectory) {
        Write-RunSummary -ExportDir $script:ExportRunDirectory -Result @{
            Mode           = $exportMode
            Status         = $finalStatus
            StartedUtc     = if ($finalStats.StartTime) { $finalStats.StartTime } else { [datetime]::UtcNow }
            Sections       = $finalSections
            RemainingTasks = $remainingTasksCount
            Errors         = @($finalStats.Errors)
        }
    }

    exit (Get-ExportExitCode -Status $finalStatus)
}
catch {
    Write-ExportLog -Message "Fatal error: $($_.Exception.Message)" -Level Error
    Write-ExportLog -Message $_.ScriptStackTrace -Level Error
    # Write a machine-readable RunSummary before exiting — non-fatal if it fails.
    if ($script:ExportRunDirectory -and (Test-Path $script:ExportRunDirectory)) {
        $fatalStats = Get-ExportStatistics
        # Report whatever completed before the crash, same shape as the happy path.
        $fatalSections = @(
            $fatalStats.ItemsExported.Keys | ForEach-Object {
                @{
                    Name        = $_
                    Status      = 'Completed'
                    RecordCount = [int]$fatalStats.ItemsExported[$_]
                    ErrorCount  = 0
                }
            }
        )
        Write-RunSummary -ExportDir $script:ExportRunDirectory -Result @{
            Mode           = if ($exportMode) { $exportMode } else { $script:SelectedMode }
            Status         = 'Failed'
            StartedUtc     = if ($fatalStats.StartTime) { $fatalStats.StartTime } else { [datetime]::UtcNow }
            Sections       = $fatalSections
            RemainingTasks = 0
            Errors         = @($fatalStats.Errors)
        }
    }
    exit 1
}
finally {
    Disconnect-Compl8Compliance
}

#endregion
