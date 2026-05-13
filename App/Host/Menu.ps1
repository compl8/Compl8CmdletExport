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
        Write-BoxLine -Text "CONTENT EXPLORER" -InnerWidth $w -Color Cyan -Single
        Write-BoxLine -Text "  [1]  Content Explorer (configurable)" -InnerWidth $w -Color Cyan -Single
        Write-BoxLine -Text "  [2]  Content Explorer - Resume Previous Export" -InnerWidth $w -Color Cyan -Single
        Write-BoxLine -Text "  [3]  Content Explorer - Retry Discrepant Tasks" -InnerWidth $w -Color Cyan -Single
        Write-BoxLine -Text "  [4]  Content Explorer - Run from Task CSV" -InnerWidth $w -Color Cyan -Single
        Write-BoxLine -Text "       * Filtered by SITstoSkip.json (all CE modes)" -InnerWidth $w -Color DarkCyan -Single
        Write-BoxSeparator -InnerWidth $w -Color Cyan -Single
        Write-BoxLine -Text "ACTIVITY EXPLORER" -InnerWidth $w -Color Cyan -Single
        Write-BoxLine -Text "  [5]  Activity Explorer (configurable)" -InnerWidth $w -Color Cyan -Single
        Write-BoxLine -Text "  [6]  Activity Explorer - Resume Previous Export" -InnerWidth $w -Color Cyan -Single
        Write-BoxSeparator -InnerWidth $w -Color Cyan -Single
        Write-BoxLine -Text "OTHER EXPORTS" -InnerWidth $w -Color Cyan -Single
        Write-BoxLine -Text "  [7]  DLP Policies & Rules" -InnerWidth $w -Color Cyan -Single
        Write-BoxLine -Text "  [8]  Sensitivity & Retention Labels" -InnerWidth $w -Color Cyan -Single
        Write-BoxLine -Text "  [9]  eDiscovery Cases" -InnerWidth $w -Color Cyan -Single
        Write-BoxLine -Text "  [10] RBAC Configuration" -InnerWidth $w -Color Cyan -Single
        Write-BoxLine -InnerWidth $w -Color Cyan -Single
        Write-BoxLine -Text "  [Q]  Quit" -InnerWidth $w -Color Cyan -Single
        Write-BoxLine -InnerWidth $w -Color Cyan -Single
        Write-BoxBottom -InnerWidth $w -Color Cyan -Single
        Write-Host ""

        $selection = Read-Host "  Enter selection [1-10, Q]"

        if ([string]::IsNullOrEmpty($selection)) { $selection = "" }
        $selectionUpper = $selection.Trim().ToUpper()

        switch ($selectionUpper) {
            "1" {
                # Content Explorer - unified configurable option
                $ceConfig = Get-CEInteractiveConfiguration
                if ($null -eq $ceConfig) { continue }
                $result.Mode = "ContentExplorer"
                $result.CEAllSITs = [bool]$ceConfig.AllSITs
                $result.CEWorkloads = $ceConfig.Workloads
                if ($ceConfig.WorkerCount -ge 2) {
                    $result.CEMultiTerminal = $true
                    $result.CEWorkerCount = $ceConfig.WorkerCount
                }
            }
            "2" {
                # Resume Previous Export - scan for ExportPhase.txt in Export-* directories
                $baseOutputDir = if ($OutputDirectory) { $OutputDirectory } else { Join-Path $scriptRoot "Output" }
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
            "3" {
                # Retry Discrepant Tasks - scan for RetryTasks.csv in Export-* directories
                $baseOutputDir = if ($OutputDirectory) { $OutputDirectory } else { Join-Path $scriptRoot "Output" }
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
            "4" {
                # Run from Task CSV - scan for RemainingTasks.csv or accept a path
                Write-Host ""
                Write-Host "  Enter path to a task CSV file, or press Enter to scan Output directory:" -ForegroundColor Yellow
                $csvInput = Read-Host "  CSV path"

                if ([string]::IsNullOrEmpty($csvInput)) {
                    # Scan for RemainingTasks.csv in Output directories
                    $baseOutputDir = if ($OutputDirectory) { $OutputDirectory } else { Join-Path $scriptRoot "Output" }
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
            "5" {
                # Activity Explorer - unified configurable option
                $aeConfig = Get-AEInteractiveConfiguration
                if ($null -eq $aeConfig) { continue }
                $result.Mode = "ActivityExplorer"
                $result.PastDays = $aeConfig.PastDays
                if ($aeConfig.WorkerCount -ge 2) {
                    $result.AEMultiTerminal = $true
                    $result.AEWorkerCount = $aeConfig.WorkerCount
                }
            }
            "6" {
                # Resume Previous AE Export - scan for ExportType.txt = "ActivityExplorer" + incomplete phase
                $baseOutputDir = if ($OutputDirectory) { $OutputDirectory } else { Join-Path $scriptRoot "Output" }
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
            "7" { $result.Mode = "DLP" }
            "8" { $result.Mode = "Labels" }
            "9" { $result.Mode = "eDiscovery" }
            "10" { $result.Mode = "RBAC" }
            "Q" { $result.Quit = $true; return $result }
            "" { $result.Quit = $true; return $result }  # Default to quit on empty input
            default {
                Write-Host "`n  Invalid selection. Exiting." -ForegroundColor Yellow
                $result.Quit = $true
                return $result
            }
        }

    } while ($null -eq $result.Mode -and -not $result.Quit)

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

function Format-CEWorkloadSummary {
    param($Workloads)
    if (-not $Workloads -or @($Workloads).Count -eq 0) { return "From config" }
    return (@($Workloads) -join ', ')
}

function Read-WorkerCountPrompt {
    param([int]$Default = 1)
    Write-Host ""
    Write-Host "  Workers: 1 = single terminal, 2-16 = multi-terminal parallel" -ForegroundColor Gray
    $wcInput = Read-Host ("  Number of workers [$Default]")
    if ([string]::IsNullOrEmpty($wcInput)) { return $Default }
    $wcVal = $wcInput -as [int]
    if ($wcVal -and $wcVal -ge 1 -and $wcVal -le 16) { return $wcVal }
    Write-Host "  Invalid worker count. Must be 1-16. Keeping previous value." -ForegroundColor Yellow
    return $Default
}

function Show-AuthModeNote {
    param([int]$WorkerCount)
    if ($WorkerCount -lt 2) { return }
    $authParams = Build-AuthParameters
    Write-Host ""
    if ($authParams.ContainsKey('AppId')) {
        Write-Host "  Authentication: Certificate (workers auto-authenticate)" -ForegroundColor Green
    } else {
        Write-Host "  Authentication: Interactive - each worker terminal will open a" -ForegroundColor Yellow
        Write-Host "  browser window requiring manual login." -ForegroundColor Yellow
    }
}

function Get-CEInteractiveConfiguration {
    <#
    .SYNOPSIS
        Interactive prompt-and-confirm flow for Content Explorer exports.
    .OUTPUTS
        Hashtable {WorkerCount, AllSITs, Workloads} or $null if cancelled.
    #>
    [CmdletBinding()]
    param()

    $config = @{
        WorkerCount = 1
        AllSITs     = $true
        Workloads   = $null
    }

    while ($true) {
        Write-Host ""
        Write-SectionHeader -Text "Content Explorer - Configure Export" -Color Cyan
        $wkrLabel = if ($config.WorkerCount -le 1) { "1 (single terminal)" } else { "{0} (multi-terminal)" -f $config.WorkerCount }
        Write-Host ("    Workers:   {0}" -f $wkrLabel) -ForegroundColor White
        Write-Host ("    SITs:      {0}" -f $(if ($config.AllSITs) { "All Sensitive Info Types (filtered by SITstoSkip.json)" } else { "From config file" })) -ForegroundColor White
        Write-Host ("    Workloads: {0}" -f (Format-CEWorkloadSummary -Workloads $config.Workloads)) -ForegroundColor White
        Show-AuthModeNote -WorkerCount $config.WorkerCount

        Write-Host ""
        $choice = (Read-Host "  [W]orkers, [S]ITs, work[L]oads, [Y] Continue, [N] Cancel").Trim().ToUpper()
        switch ($choice) {
            'W' { $config.WorkerCount = Read-WorkerCountPrompt -Default $config.WorkerCount }
            'S' {
                $sitsInput = (Read-Host ("  Scan ALL SITs? Currently {0} [Y/N]" -f $(if ($config.AllSITs) { 'Yes' } else { 'No' }))).Trim().ToUpper()
                if ($sitsInput -eq 'Y') { $config.AllSITs = $true }
                elseif ($sitsInput -eq 'N') { $config.AllSITs = $false }
            }
            'L' { $config.Workloads = Get-CEWorkloadSelection }
            'Y' { return $config }
            ''  { return $config }
            'N' { return $null }
            default { Write-Host "  Unrecognized choice." -ForegroundColor Yellow }
        }
    }
}

function Get-AEInteractiveConfiguration {
    <#
    .SYNOPSIS
        Interactive prompt-and-confirm flow for Activity Explorer exports.
    .OUTPUTS
        Hashtable {WorkerCount, PastDays} or $null if cancelled.
    #>
    [CmdletBinding()]
    param()

    $config = @{
        WorkerCount = 1
        PastDays    = 7
    }

    while ($true) {
        Write-Host ""
        Write-SectionHeader -Text "Activity Explorer - Configure Export" -Color Cyan
        $wkrLabel = if ($config.WorkerCount -le 1) { "1 (single terminal)" } else { "{0} (multi-terminal)" -f $config.WorkerCount }
        Write-Host ("    Workers:   {0}" -f $wkrLabel) -ForegroundColor White
        Write-Host ("    Past days: {0}" -f $config.PastDays) -ForegroundColor White
        Show-AuthModeNote -WorkerCount $config.WorkerCount

        Write-Host ""
        $choice = (Read-Host "  [W]orkers, [D]ays, [Y] Continue, [N] Cancel").Trim().ToUpper()
        switch ($choice) {
            'W' { $config.WorkerCount = Read-WorkerCountPrompt -Default $config.WorkerCount }
            'D' {
                $daysInput = Read-Host ("  Past days to export (1-30) [{0}]" -f $config.PastDays)
                if ($daysInput -match '^\d+$') {
                    $days = [int]$daysInput
                    if ($days -ge 1 -and $days -le 30) { $config.PastDays = $days }
                    else { Write-Host "  Out of range (1-30). Keeping previous value." -ForegroundColor Yellow }
                }
            }
            'Y' { return $config }
            ''  { return $config }
            'N' { return $null }
            default { Write-Host "  Unrecognized choice." -ForegroundColor Yellow }
        }
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

    $authConfigPath = Join-Path $scriptRoot "ConfigFiles\AuthConfig.json"

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

    $scriptPath = Join-Path $scriptRoot "Export-Compl8Configuration.ps1"

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

    $scriptPath = Join-Path $scriptRoot "Export-Compl8Configuration.ps1"
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

