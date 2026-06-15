# ContentExplorer Resume — resume an incomplete export from its last phase
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
    if (-not $script:Unattended) {
        $confirm = Read-Host "  Resume this export? [Y/n]"
        if (-not [string]::IsNullOrEmpty($confirm) -and $confirm.Trim().ToUpper() -ne "Y") {
            Write-Host "  Resume cancelled." -ForegroundColor Yellow
            return
        }
    } else {
        Write-ExportLog -Message "Unattended: proceeding with resume without confirmation (prompt C skipped)." -Level Info
    }

    # Initialize logging
    Initialize-ExportLog -LogDirectory (Get-LogsDir $ExportDir)
    $script:ExportRunDirectory = $ExportDir
    $script:SharedExportDirectory = $ExportDir
    $script:ErrorLogPath = Join-Path (Get-LogsDir $ExportDir) "ExportProject-Errors.log"

    Write-ExportLog -Message ("Resuming export from phase: {0}" -f $phase) -Level Info

    # Load Content Explorer page size (manifest overrides config file for consistency)
    $configPath = Join-Path $scriptRoot "ConfigFiles" "ContentExplorerClassifiers.json"
    $resolved = Resolve-CEPageSize -ExportRunDirectory $ExportDir -ConfigPath $configPath -FallbackPageSize $PageSize
    $cePageSize = $resolved.PageSize
    $savedSettings = Get-ExportSettings -ExportRunDirectory $ExportDir
    if ($savedSettings -and $savedSettings.JsonlOutput -eq $true) { $env:COMPL8_JSONL_OUTPUT = "1" }
    $ceConfig = Read-JsonConfig -Path $configPath
    $ceSettings = Get-ContentExplorerSettings -ConfigObject $ceConfig -SavedSettings $savedSettings -DefaultBatchSize $script:CEDefaultBatchSize -DefaultWorkloads $script:CEDefaultWorkloads -DefaultPageSize $cePageSize
    $largeAllSITDetailThreshold = $ceSettings.LargeAllSITDetailThreshold
    $largeAllSITFallbackCandidates = @($ceSettings.LargeAllSITWorkloadFallbackWorkloads)
    $minLocationItems = $ceSettings.MinLocationItems
    $savedCEAllSITs = ($savedSettings -and $savedSettings.CEAllSITs -eq $true)

    $sitsToSkipPath = Join-Path $scriptRoot "ConfigFiles" "SITstoSkip.json"
    $sitsToSkip = Get-SITsToSkip -ConfigPath $sitsToSkipPath

    $progressLogPath = Join-Path (Get-LogsDir $ExportDir) "ContentExplorer-Progress.log"
    $aggregateCsvPath = Join-Path (Get-CoordinationDir $ExportDir) "ContentExplorer-Aggregates.csv"
    $trackerPath = Join-Path (Get-CoordinationDir $ExportDir) "RunTracker.json"
    $tracker = Get-ContentExplorerRunTracker -TrackerPath $trackerPath
    $telemetryDbPath = Join-Path $scriptRoot "TelemetryDB" "content-explorer-telemetry.jsonl"

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
                    # Shared aggregate pagination loop (resume strategy: single call,
                    # no per-page retry — errors propagate to the catch below, exactly
                    # as the prior inline loop did). Cookie null/empty + same-cookie
                    # guards are preserved inside Invoke-CEAggregatePaging.
                    $allAggregates = @(Invoke-CEAggregatePaging `
                        -TagType $aggTask.TagType -TagName $aggTask.TagName -Workload $aggTask.Workload `
                        -PageSize 5000 -RetryMode 'None')

                    # Write aggregate results to central CSV
                    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                    foreach ($agg in $allAggregates) {
                        $locationName = $agg.Name -replace '"', '""'
                        if ($locationName -match '[,"]') { $locationName = ('"' + $locationName + '"') }
                        $csvLine = "{0},{1},{2},{3},{4},{5}," -f $timestamp, $aggTask.TagType, (ConvertTo-CsvField $aggTask.TagName), $aggTask.Workload, $locationName, $agg.Count
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
                    $fcLine = "{0},{1},{2},{3},_FILECOUNT,{4}," -f $timestamp, $aggTask.TagType, (ConvertTo-CsvField $aggTask.TagName), $aggTask.Workload, $fileCount
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
                $allAggregateTaskData = @()
                foreach ($tagType in $tagTypes) {
                    $tagNames = @($aggTasks | Where-Object { $_.TagType -eq $tagType } | Select-Object -ExpandProperty TagName -Unique)
                    if ($tagNames.Count -eq 0) { continue }
                    $cachedResult = Import-AggregateDataFromCsv -CsvPath $aggregateCsvPath -TagType $tagType -TagNames $tagNames -Workloads $workloads
                    if ($cachedResult -and $cachedResult.TaskData) {
                        $allAggregateTaskData += @($cachedResult.TaskData.Values)
                    }
                }

                $resumeFallbackWorkloads = @()
                if ($savedCEAllSITs) {
                    $resumeSitCount = @($allAggregateTaskData | Where-Object { $_.TagType -eq "SensitiveInformationType" } | ForEach-Object { $_.TagName } | Select-Object -Unique).Count
                    if ($resumeSitCount -gt $largeAllSITDetailThreshold) {
                        $resumeFallbackWorkloads = @($workloads | Where-Object { $_ -in $largeAllSITFallbackCandidates })
                        if ($resumeFallbackWorkloads.Count -gt 0) {
                            Write-ExportLog -Message ("Large All-SIT detail strategy restored from settings: {0} SITs > threshold {1}; using workload-level detail tasks for {2}" -f $resumeSitCount, $largeAllSITDetailThreshold, ($resumeFallbackWorkloads -join ', ')) -Level Info
                        }
                    }
                }

                $script:AllWorkPlanTasks = @(New-ContentExplorerDetailTasks -WorkPlanTasks $allAggregateTaskData -DefaultPageSize $cePageSize -WorkloadFallbackWorkloads $resumeFallbackWorkloads -MinLocationItems $minLocationItems -Sort)

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
                        $dtKey = "{0}|{1}|{2}|{3}" -f $dt.TagType, $dt.TagName, $dt.Workload, $(if ($dt.Location) { $dt.Location } else { "" })
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
                        $taskKey = "{0}|{1}|{2}|{3}" -f $task.TagType, $task.TagName, $task.Workload, $(if ($task.Location) { $task.Location } else { "" })
                        if ($completedTaskKeys.ContainsKey($taskKey)) { continue }
                        if ((-not $task.ExpectedCount -or $task.ExpectedCount -eq 0) -and $task.Status -ne "Error") { continue }

                        $detailTasks += @{
                            Phase                 = "Detail"
                            TagType               = $task.TagType
                            TagName               = $task.TagName
                            Workload              = $task.Workload
                            Location              = if ($task.Location) { $task.Location } else { "" }
                            LocationType          = if ($task.LocationType) { $task.LocationType } else { "WorkloadFallback" }
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
                    # Spawn workers. Collect into a real ArrayList: the dispatch loop and the
                    # W-hotkey [ref] must share this exact instance for dynamic adds to work.
                    $workerProcesses = [System.Collections.ArrayList]@(Start-WorkerTerminals -ExportRunDirectory $ExportDir -Count $WorkerCount)
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
                    # Thin wrapper over Read-CEDetailSignals (module). Shared verbatim
                    # with Invoke-ContentExplorerFromTasksCsv.
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

                    # --- CE Resume Callback: OnDispatchTask ---
                    $ceResumeOnDispatch = {
                        param($Worker, $NextTask, $Context)
                        $taskData = New-CEDetailDispatchPayload -NextTask $NextTask
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
                        -OnSelectNextTask ${function:Select-LargestPendingTask} `
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
                    $taskLocation = if ($task.Location) { $task.Location } else { "" }
                    $taskKey = "{0}|{1}|{2}|{3}" -f $task.TagType, $task.TagName, $task.Workload, $taskLocation

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
                    # Single-terminal detail splat (resume): page size = $cePageSize, with telemetry.
                    $exportParams = Build-CEDetailExportParams -Task $task -PageSize $cePageSize `
                        -ProgressLogPath $progressLogPath -TelemetryDatabasePath $telemetryDbPath `
                        -OutputDirectory $classifierDir -Telemetry $telemetry

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
