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
    $aeConfigPath = Join-Path $scriptRoot "ConfigFiles\ActivityExplorerSelector.json"
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

        # Bounded retry-in-place for auth/token recovery: on a successful silent
        # reconnect we re-run the SAME task rather than abandoning it. A reconnected
        # worker stays alive, so the orchestrator's dead-worker reclaim never fires;
        # retrying here is the only way the interrupted day actually gets finished.
        $authRecoveryAttempts = 0
        $maxAuthRecovery = 3
        $retryTask = $true
        $workerShouldExit = $false
        while ($retryTask) {
        $retryTask = $false

        # (Re)load per-day tracker so a retry resumes from pages already saved.
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

            $completionsDir = Get-CompletionsDir $exportDir
            if (-not (Test-Path $completionsDir)) {
                New-Item -ItemType Directory -Force -Path $completionsDir | Out-Null
            }

            $resultStatus = if ($exportResult.Status) { [string]$exportResult.Status } else { "Completed" }
            $isPartialFailure = ($exportResult.PartialFailure -eq $true) -or ($resultStatus -ne "Completed")

            if ($isPartialFailure) {
                $partialErrorSummary = $null
                if ($exportResult.PartialErrors -and $exportResult.PartialErrors.Count -gt 0) {
                    $partialErrorSummary = ($exportResult.PartialErrors | Select-Object -Last 1).ErrorMessage
                }
                $errMessage = "Day {0} export status={1} after {2} pages ({3} records). Last error: {4}" -f `
                    $taskDay, $resultStatus, $exportResult.PageCount, $exportResult.TotalRecords, $partialErrorSummary

                Write-ExportLog -Message ("  Day {0} {1}: {2} records, {3} pages in {4} - emitting error signal" -f $taskDay, $resultStatus, $exportResult.TotalRecords, $exportResult.PageCount, (Format-TimeSpan -Seconds $taskElapsed.TotalSeconds)) -Level Error
                Write-ProgressEntry -Path $progressLogPath -Message ("Day {0} {1}: {2} records, {3} pages" -f $taskDay, $resultStatus, $exportResult.TotalRecords, $exportResult.PageCount)

                $errorFile = Join-Path $completionsDir ("error-ae-{0}-{1}.txt" -f $taskDay, $PID)
                $errorPayload = @{
                    Day          = $taskDay
                    ErrorMessage = $errMessage
                    ErrorType    = $resultStatus
                    RecordCount  = $exportResult.TotalRecords
                    PageCount    = $exportResult.PageCount
                }
                (ConvertTo-SignedEnvelopeJson -Payload $errorPayload -SigningKey $signalSigningKey) | Set-Content -Path $errorFile -Encoding UTF8
            }
            else {
                Write-ExportLog -Message ("  Day {0} complete: {1} records, {2} pages in {3}" -f $taskDay, $exportResult.TotalRecords, $exportResult.PageCount, (Format-TimeSpan -Seconds $taskElapsed.TotalSeconds)) -Level Success
                Write-ProgressEntry -Path $progressLogPath -Message ("Day {0} complete: {1} records, {2} pages" -f $taskDay, $exportResult.TotalRecords, $exportResult.PageCount)

                $doneFile = Join-Path $completionsDir ("ae-done-{0}-{1}.txt" -f $taskDay, $PID)
                $donePayload = @{
                    Day            = $taskDay
                    RecordCount    = $exportResult.TotalRecords
                    PageCount      = $exportResult.PageCount
                    ElapsedSeconds = [int]$taskElapsed.TotalSeconds
                }
                (ConvertTo-SignedEnvelopeJson -Payload $donePayload -SigningKey $signalSigningKey) | Set-Content -Path $doneFile -Encoding UTF8
            }

        }
        catch {
            $taskElapsed = (Get-Date) - $taskStart
            $aeErrMsg = $_.Exception.Message
            $errorInfo = Get-HttpErrorExplanation -ErrorMessage $aeErrMsg -ErrorRecord $_
            $isAuthError = ($errorInfo.Category -eq "AuthError") -or ($_.Exception -is [System.Management.Automation.CommandNotFoundException])

            if ($isAuthError) {
                # Auth/token expiry: try a silent reconnect BEFORE signalling so we
                # don't permanently error a day that was only interrupted by an
                # expired token. On success, RE-RUN the same task in place (a
                # reconnected worker stays alive, so the orchestrator never reclaims
                # its in-progress task). On failure or once retries are exhausted,
                # exit the worker so the task is reclaimed via the stale-worker lease
                # (Progress.log goes stale once the loop exits). Do NOT also write an
                # error signal here, or the day would be permanently errored.
                Write-ExportLog -Message ("  Day {0}: AUTH/CONNECTION error - attempting recovery before signalling" -f $taskDay) -Level Warning
                Write-ProgressEntry -Path $progressLogPath -Message ("Day {0}: auth error - attempting recovery" -f $taskDay)
                if (($authRecoveryAttempts -lt $maxAuthRecovery) -and (Invoke-WorkerReconnect -AuthParams $script:AuthParams)) {
                    $authRecoveryAttempts++
                    Write-ExportLog -Message ("  Day {0}: recovery successful - retrying same task (attempt {1}/{2})" -f $taskDay, $authRecoveryAttempts, $maxAuthRecovery) -Level Success
                    Write-ProgressEntry -Path $progressLogPath -Message ("Day {0}: auth recovered, retrying task (attempt {1}/{2})" -f $taskDay, $authRecoveryAttempts, $maxAuthRecovery)
                    $lastWorkerActivity = Get-Date
                    $retryTask = $true
                    continue
                }
                Write-ExportLog -Message ("  Day {0}: recovery failed/exhausted - worker exiting (task returned to queue)" -f $taskDay) -Level Error
                $workerShouldExit = $true
            }
            else {

            # Non-auth error: write error signal so the orchestrator records the failure.
            Write-ExportLog -Message ("  Day {0} FAILED: {1}" -f $taskDay, $aeErrMsg) -Level Error
            Write-ProgressEntry -Path $progressLogPath -Message ("Day {0} FAILED: {1}" -f $taskDay, $aeErrMsg)

            $completionsDir = Get-CompletionsDir $exportDir
            if (-not (Test-Path $completionsDir)) {
                New-Item -ItemType Directory -Force -Path $completionsDir | Out-Null
            }
            $errorFile = Join-Path $completionsDir ("error-ae-{0}-{1}.txt" -f $taskDay, $PID)
            $errorPayload = @{
                Day          = $taskDay
                ErrorMessage = $aeErrMsg
                ErrorType    = $_.Exception.GetType().Name
            }
            (ConvertTo-SignedEnvelopeJson -Payload $errorPayload -SigningKey $signalSigningKey) | Set-Content -Path $errorFile -Encoding UTF8

            if ($script:ErrorLogPath) {
                Write-ExportErrorLog -ErrorLogPath $script:ErrorLogPath -Context "AE Worker Day Export" -TaskKey $taskDay -ErrorRecord $_ -AdditionalData @{ Day = $taskDay }
            }
            if ($script:WorkerErrorLogPath) {
                Write-ExportErrorLog -ErrorLogPath $script:WorkerErrorLogPath -Context "AE Worker Day Export" -TaskKey $taskDay -ErrorRecord $_ -AdditionalData @{ Day = $taskDay }
            }
            }  # end non-auth error else
        }
        }  # end retry-in-place loop

        Complete-WorkerTask -WorkerDir $workerDir
        $lastWorkerActivity = Get-Date
        if ($workerShouldExit) { break }
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
            $aeConfigPath = Join-Path $scriptRoot "ConfigFiles\ActivityExplorerSelector.json"
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
        $configPath = Join-Path $scriptRoot "ConfigFiles\ActivityExplorerSelector.json"
        $selectorConfig = Read-JsonConfig -Path $configPath
        Save-ExportSettings -ExportRunDirectory $exportDir -ExportType "ActivityExplorer" -Settings @{
            PastDays       = $PastDays
            PageSize       = $PageSize
            SelectorConfig = $selectorConfig
        }

        # SIT reference snapshot: GUID->name map (flat list + rule packages) written to the
        # export root so downstream analytics resolve SIT/sub-entity GUIDs to display names.
        Export-SitReferenceSnapshot -ExportRunDirectory $exportDir | Out-Null
    }

    # ==== Shared multi-terminal dispatch (both fresh and resume paths) ====

    # Ensure Completions directory exists
    $completionsDir = Get-CompletionsDir $exportDir
    if (-not (Test-Path $completionsDir)) {
        New-Item -ItemType Directory -Force -Path $completionsDir | Out-Null
    }

    Write-ExportLog -Message ("Task CSV: {0}" -f $aeTaskCsvPath) -Level Info

    # Spawn workers. Collect into a real ArrayList: the dispatch loop must share this
    # exact instance (matches the CE orchestrator paths).
    $workerProcesses = [System.Collections.ArrayList]@(Start-WorkerTerminals -ExportRunDirectory $exportDir -Count $actualWorkers)
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

    # SIT reference snapshot: GUID->name map (flat list + rule packages) written to the
    # export root so downstream analytics resolve SIT/sub-entity GUIDs to display names.
    Export-SitReferenceSnapshot -ExportRunDirectory $script:ExportRunDirectory | Out-Null

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

