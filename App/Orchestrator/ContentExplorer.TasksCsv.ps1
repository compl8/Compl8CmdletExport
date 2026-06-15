# ContentExplorer TasksCsv — run detail export from an input task CSV
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
    if (-not $script:Unattended) {
        $confirm = Read-Host "  Run these tasks? [Y/N]"
        if ([string]::IsNullOrEmpty($confirm) -or $confirm.Trim().ToUpper() -ne "Y") {
            Write-Host "  Cancelled." -ForegroundColor Yellow
            return
        }
    } else {
        Write-ExportLog -Message "Unattended: proceeding with task CSV run without confirmation (prompt E skipped)." -Level Info
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
    $configPath = Join-Path $scriptRoot "ConfigFiles" "ContentExplorerClassifiers.json"
    $ceConfig = Read-JsonConfig -Path $configPath
    $ceSettings = Get-ContentExplorerSettings -ConfigObject $ceConfig -DefaultBatchSize $script:CEDefaultBatchSize -DefaultWorkloads $script:CEDefaultWorkloads -DefaultPageSize $PageSize
    $cePageSize = $ceSettings.PageSize
    if (-not $cePageSize -or $cePageSize -lt 1) { $cePageSize = 100 }

    $progressLogPath = Join-Path (Get-LogsDir $exportDir) "ContentExplorer-Progress.log"
    $trackerPath = Join-Path (Get-CoordinationDir $exportDir) "RunTracker.json"
    $tracker = Get-ContentExplorerRunTracker -TrackerPath $trackerPath
    $telemetryDbPath = Join-Path $scriptRoot "TelemetryDB" "content-explorer-telemetry.jsonl"

    # Build work plan tasks from input CSV
    $script:AllWorkPlanTasks = @()
    foreach ($t in $inputTasks) {
        if ($t.Status -eq "Completed") { continue }
        $script:AllWorkPlanTasks += @{
            TagType               = $t.TagType
            TagName               = $t.TagName
            Workload              = $t.Workload
            Location              = if ($t.Location) { $t.Location } else { "" }
            LocationType          = if ($t.LocationType) { $t.LocationType } else { "WorkloadFallback" }
            ExpectedCount         = ($t.ExpectedCount -as [int])
            OriginalExpectedCount = if ($t.OriginalExpectedCount) { ($t.OriginalExpectedCount -as [int]) } else { ($t.ExpectedCount -as [int]) }
            PageSize              = if ($t.PageSize) { ($t.PageSize -as [int]) } else { $cePageSize }
            ExportedCount         = 0
            Locations             = @()
            Status                = "Pending"
        }
    }

    Write-ExportLog -Message ("  {0} tasks to export" -f $script:AllWorkPlanTasks.Count) -Level Info

    if ($WorkerCount -gt 0) {
        # -- Multi-terminal: spawn workers and dispatch tasks --
        $detailTasks = @($inputTasks | Where-Object { $_.Status -eq "Pending" })
        # Sort: largest first
        $detailTasks = @($detailTasks | Sort-Object { [int]$_.ExpectedCount } -Descending)
        Write-TaskCsv -Path $detailCsvPath -Tasks $inputTasks

        # Collect into a real ArrayList: the dispatch loop must share this exact instance.
        $workerProcesses = [System.Collections.ArrayList]@(Start-WorkerTerminals -ExportRunDirectory $exportDir -Count $WorkerCount)
        if ($workerProcesses.Count -eq 0) {
            Write-ExportLog -Message "  No workers spawned - aborting" -Level Error
            Write-Host "  ERROR: No workers were spawned." -ForegroundColor Red
            return
        }

        # Guarantee worker shutdown on both normal completion and exception/abort.
        try {

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
        # Thin wrapper over Read-CEDetailSignals (module). Shared verbatim with
        # Invoke-ContentExplorerResume.
        $ceDetailOnScan = {
            param($ExportDir, $WorkerDirs, $Context)
            return (Read-CEDetailSignals -ExportDir $ExportDir -WorkerDirs $WorkerDirs)
        }

        # --- Shared CE Detail Callback: OnMatchTask ---
        # Thin wrapper over Find-CEDetailTaskMatch (module).
        $ceDetailOnMatch = {
            param($Data, $Tasks, $Context)
            return (Find-CEDetailTaskMatch -Data $Data -Tasks $Tasks)
        }

        # --- CE TasksCsv Callback: OnDispatchTask ---
        $csvOnDispatch = {
            param($Worker, $NextTask, $Context)
            $taskData = New-CEDetailDispatchPayload -NextTask $NextTask
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
            -OnSelectNextTask ${function:Select-LargestPendingTask} `
            -SleepSeconds 2

        # Save final task state
        Write-TaskCsv -Path $detailCsvPath -Tasks $inputTasks
        }
        finally {
            # Guarantee worker shutdown on both normal completion and exception/abort.
            # Stop-WorkerProcesses is a no-op when $workerProcesses is empty.
            Stop-WorkerProcesses -WorkerProcesses $workerProcesses
        }
    }
    else {
        # -- Single-terminal: process detail tasks directly --
        $completedTaskCounts = @{}
        foreach ($task in $script:AllWorkPlanTasks) {
            $taskLocation = if ($task.Location) { $task.Location } else { "" }
            $taskKey = "{0}|{1}|{2}|{3}" -f $task.TagType, $task.TagName, $task.Workload, $taskLocation

            if ((-not $task.ExpectedCount -or $task.ExpectedCount -eq 0) -and $task.Status -ne "Error") {
                Write-ExportLog -Message ("    Skipping {0} / {1} - no data" -f $task.TagName, $task.Workload) -Level Info
                continue
            }

            Write-ExportLog -Message ("    Exporting: {0} / {1} (expected {2})" -f $task.TagName, $task.Workload, $task.ExpectedCount) -Level Info

            # Output to Data/ContentExplorer/TagType/TagName/
            $classifierDir = Get-CEClassifierDir $exportDir $task.TagType $task.TagName

            $telemetry = New-ContentExplorerTelemetry -TagType $task.TagType -TagName $task.TagName -Workload $task.Workload
            # Single-terminal detail splat (tasks-CSV): per-task page size (same expression as before), with telemetry.
            $taskPageSize = if ($task.PageSize -and ($task.PageSize -as [int]) -gt 0) { ($task.PageSize -as [int]) } else { $cePageSize }
            $exportParams = Build-CEDetailExportParams -Task $task -PageSize $taskPageSize `
                -ProgressLogPath $progressLogPath -TelemetryDatabasePath $telemetryDbPath `
                -OutputDirectory $classifierDir -Telemetry $telemetry

            try {
                Export-ContentExplorerWithProgress @exportParams | Out-Null
                $exportedCount = if ($task.ExportedCount) { $task.ExportedCount } else { 0 }

                if ($exportedCount -gt 0) {
                    Write-ExportLog -Message ("    Completed: {0} / {1} - {2} records" -f $task.TagName, $task.Workload, $exportedCount) -Level Success
                }

                $completedTaskCounts[$taskKey] = $exportedCount

                # Update task status in CSV
                $csvTask = $inputTasks | Where-Object {
                    $_.TagType -eq $task.TagType -and $_.TagName -eq $task.TagName -and $_.Workload -eq $task.Workload -and $(if ($_.Location) { $_.Location } else { "" }) -eq $taskLocation
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
                    $_.TagType -eq $task.TagType -and $_.TagName -eq $task.TagName -and $_.Workload -eq $task.Workload -and $(if ($_.Location) { $_.Location } else { "" }) -eq $taskLocation
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
