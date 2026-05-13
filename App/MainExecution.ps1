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
if ($script:TenantPrefix) {
    $prefixSource = if ($prefixResult) { $prefixResult.Source } else { "previous" }
    Write-Host ("  Tenant:           {0} (from {1})" -f $script:TenantPrefix, $prefixSource)
}
Write-Host "  Output Directory: $script:ExportRunDirectory"
if ($UnifiedParquet) {
    $plannedUnifiedParquetDir = Resolve-UnifiedParquetOutputDir -ConfiguredPath $UnifiedParquetDir -ExportRunDirectory $script:ExportRunDirectory
    Write-Host "  C8 Tuning Input:  $plannedUnifiedParquetDir"
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
