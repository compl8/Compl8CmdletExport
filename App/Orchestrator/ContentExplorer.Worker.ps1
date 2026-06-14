# ContentExplorer Worker — file-drop worker (Aggregate + Detail phases)
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
    $configPath = Join-Path $scriptRoot "ConfigFiles" "ContentExplorerClassifiers.json"
    $ceConfig = Read-JsonConfig -Path $configPath
    if (-not $ceConfig -and -not $savedSettings) {
        Write-ExportLog -Message "ERROR: Cannot read Content Explorer config" -Level Error
        return
    }

    # Extract config settings (manifest overrides config file)
    if ($savedSettings) {
        Write-ExportLog -Message "CE Worker using saved settings from ExportSettings.json" -Level Info
        if ($savedSettings.JsonlOutput -eq $true) { $env:COMPL8_JSONL_OUTPUT = "1" }
    }
    else {
        Write-ExportLog -Message "CE Worker using current config file (no manifest found)" -Level Warning
    }
    $ceSettings = Get-ContentExplorerSettings -ConfigObject $ceConfig -SavedSettings $savedSettings -DefaultBatchSize $script:CEDefaultBatchSize -DefaultWorkloads $script:CEDefaultWorkloads -DefaultPageSize $PageSize
    $batchSize = $ceSettings.BatchSize
    $workloads = @($ceSettings.Workloads)
    $cePageSize = $ceSettings.PageSize
    $largeAllSITDetailThreshold = $ceSettings.LargeAllSITDetailThreshold
    $largeAllSITFallbackCandidates = @($ceSettings.LargeAllSITWorkloadFallbackWorkloads)

    # Per-worker run tracker (in worker folder)
    $trackerPath = Join-Path $workerDir "RunTracker.json"
    $tracker = Get-ContentExplorerRunTracker -TrackerPath $trackerPath

    # Telemetry setup
    $telemetryDbPath = Join-Path $scriptRoot "TelemetryDB" "content-explorer-telemetry.jsonl"

    # Aggregate data (loaded once when entering Detail phase)
    $aggregateDataLoaded = $false
    $aggregateData = $null

    $tasksExported = 0
    $cmdletNotFoundCount = 0
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
        $sanitizedKey = ConvertTo-SafeDirectoryName -Name ("{0}-{1}-{2}" -f $taskTagType, $taskTagName, $taskWorkload)

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

            try {
                # Shared aggregate pagination loop (worker reconnect-in-place strategy).
                # $script:AuthParams / $script:ErrorLogPath / $script:WorkerErrorLogPath
                # live in the App/orchestrator script scope, NOT the module scope, so
                # they MUST be passed explicitly — a module function cannot see them.
                $allAggregates = @(Invoke-CEAggregatePaging `
                    -TagType $taskTagType -TagName $taskTagName -Workload $taskWorkload `
                    -PageSize 5000 -RetryMode 'WorkerReconnect' `
                    -AuthParams $script:AuthParams -WriteWorkerErrorLog `
                    -ErrorLogPath $script:ErrorLogPath -WorkerErrorLogPath $script:WorkerErrorLogPath `
                    -TaskKey $taskKey)

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
                    $csvLine = "{0},{1},{2},{3},{4},{5}," -f $timestamp, $taskTagType, (ConvertTo-CsvField $taskTagName), $taskWorkload, $locationName, $agg.Count
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
                $csvLine = "{0},{1},{2},{3},_FILECOUNT,{4}," -f $timestamp, $taskTagType, (ConvertTo-CsvField $taskTagName), $taskWorkload, $fileCount
                [void]$csvSb.AppendLine($csvLine)

                # Write marker row for zero-result tasks so orchestrator can identify and complete them
                if ($allAggregates.Count -eq 0) {
                    $csvLine = "{0},{1},{2},{3},NONE,0," -f $timestamp, $taskTagType, (ConvertTo-CsvField $taskTagName), $taskWorkload
                    [void]$csvSb.AppendLine($csvLine)
                }

                # Write per-worker aggregate CSV atomically (temp+move) so the
                # orchestrator never reads a partial-but-parseable file mid-write.
                $aggTempPath = "{0}.tmp.{1}" -f $aggCsvPath, $PID
                [System.IO.File]::WriteAllText($aggTempPath, $csvSb.ToString(), [System.Text.Encoding]::UTF8)
                [System.IO.File]::Move($aggTempPath, $aggCsvPath, $true)

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

                    # Write per-worker aggregate CSV with error row (atomic temp+move)
                    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                    $escapedError = $aggError -replace '"', '""'
                    $csvSb = [System.Text.StringBuilder]::new()
                    [void]$csvSb.AppendLine("Timestamp,TagType,TagName,Workload,Location,Count,Error")
                    $errorCsvLine = '{0},{1},{2},{3},ERROR,0,"{4}"' -f $timestamp, $taskTagType, (ConvertTo-CsvField $taskTagName), $taskWorkload, $escapedError
                    [void]$csvSb.AppendLine($errorCsvLine)
                    $aggErrorTempPath = "{0}.tmp.{1}" -f $aggCsvPath, $PID
                    [System.IO.File]::WriteAllText($aggErrorTempPath, $csvSb.ToString(), [System.Text.Encoding]::UTF8)
                    [System.IO.File]::Move($aggErrorTempPath, $aggCsvPath, $true)

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
            # Deterministic SHA256 hash (not GetHashCode, which is randomized
            # per-process in .NET 5+) so the suffix matches the page-file prefix
            # built inside Export-ContentExplorerWithProgress across processes.
            $locationSuffix = if ($task.Location) { "-" + (Get-DeterministicNameHash -Name $task.Location) } else { "" }
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

            # Bounded retry-in-place for auth/token recovery (see AE worker for
            # rationale): a reconnected worker stays alive, so the orchestrator's
            # dead-worker reclaim never fires - the task must be finished here.
            # $detailTask is rebuilt each attempt so its Status/ExportedCount reset.
            $authRecoveryAttempts = 0
            $maxAuthRecovery = 3
            $retryTask = $true
            $workerShouldExit = $false
            while ($retryTask) {
            $retryTask = $false

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
                # On an auth-recovery retry (attempt > 0), clear the failed attempt's
                # partial pages before re-running. First attempts and other callers
                # (resume, retry-bucket) do NOT pass this so existing pages are
                # preserved - protects retry-bucket from losing records on a shrink.
                $cleanFlag = @{}
                if ($authRecoveryAttempts -gt 0) { $cleanFlag['CleanPriorPages'] = $true }
                Export-ContentExplorerWithProgress -Task $detailTask -PageSize $taskPageSize -ProgressLogPath $progressLogPath -Telemetry $telemetry -TelemetryDatabasePath $telemetryDbPath -OutputDirectory $classifierDir @locationParams @cleanFlag | Out-Null

                $exportedCount = if ($detailTask.ExportedCount) { $detailTask.ExportedCount } else { 0 }
                $taskStatus = if ($detailTask.Status) { [string]$detailTask.Status } else { "Completed" }
                $isFailureOutcome = ($taskStatus -eq "Failed" -or $taskStatus -eq "PartialFailure")

                # Update run tracker regardless of outcome (partial data on disk is still useful)
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
                        Status          = $taskStatus
                    }
                }

                Save-ContentExplorerRunTracker -Tracker $tracker -TrackerPath $trackerPath

                $detailTaskElapsed = ((Get-Date) - $detailTaskStartTime).TotalSeconds

                if ($isFailureOutcome) {
                    # Export function set Failed/PartialFailure without throwing - route to error signal
                    $errMsg = "Detail export status={0} after {1} records, {2} pages" -f $taskStatus, $exportedCount, $detailTask.TotalPages
                    Write-ExportLog -Message ("    DETAIL {0}: {1} records (partial data preserved)" -f $taskStatus.ToUpper(), $exportedCount) -Level Error
                    Write-ProgressEntry -Path $progressLogPath -Message ("Detail {0}: {1} -> {2} records" -f $taskStatus, $taskKey, $exportedCount)

                    $errorPayload = @{
                        Timestamp    = (Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff")
                        TaskKey      = $taskKey
                        TagType      = $taskTagType
                        TagName      = $taskTagName
                        Workload     = $taskWorkload
                        Location     = if ($task.Location) { $task.Location } else { "" }
                        LocationType = if ($task.LocationType) { $task.LocationType } else { "" }
                        Error        = $errMsg
                        RecordCount  = $exportedCount
                        Status       = $taskStatus
                    }
                    $errorJson = ConvertTo-SignedEnvelopeJson -Payload $errorPayload -SigningKey $signalSigningKey
                    [System.IO.File]::WriteAllText($detailErrorPath, $errorJson, [System.Text.Encoding]::UTF8)
                }
                else {
                    if ($exportedCount -gt 0) {
                        Write-ExportLog -Message ("    -> Exported {0} records to {1}" -f $exportedCount, $classifierDir) -Level Success
                    }
                    else {
                        Write-ExportLog -Message ("    -> No records exported for {0}" -f $taskKey) -Level Info
                    }

                    Write-ProgressEntry -Path $progressLogPath -Message ("Detail complete: {0} -> {1} records" -f $taskKey, $exportedCount)

                    # Write detail-done signal file (orchestrator watches for these)
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
            }
            catch {
                $detailError = $_.Exception.Message
                $isCmdletNotFound = $_.Exception -is [System.Management.Automation.CommandNotFoundException]
                $detailErrInfo = Get-HttpErrorExplanation -ErrorMessage $detailError -ErrorRecord $_
                $isAuthError = ($detailErrInfo.Category -eq "AuthError")

                if ($isAuthError) {
                    # Auth/token expiry: attempt a silent reconnect, then RE-RUN the
                    # same task in place. A reconnected worker stays alive, so the
                    # orchestrator never reclaims its in-progress task; retrying here
                    # is the only way the task gets finished. On failure or once
                    # retries are exhausted, exit so the stale-worker lease reclaims
                    # it (Progress.log goes stale once the loop exits).
                    Write-ExportLog -Message ("    DETAIL AUTH EXPIRED: {0} - attempting reconnect" -f $taskKey) -Level Warning
                    Write-ProgressEntry -Path $progressLogPath -Message ("Auth expired: {0} -> attempting recovery" -f $taskKey)
                    if (($authRecoveryAttempts -lt $maxAuthRecovery) -and (Invoke-WorkerReconnect -AuthParams $script:AuthParams)) {
                        $authRecoveryAttempts++
                        Write-ExportLog -Message ("    DETAIL recovery successful - retrying same task (attempt {0}/{1})" -f $authRecoveryAttempts, $maxAuthRecovery) -Level Success
                        Write-ProgressEntry -Path $progressLogPath -Message ("Auth recovered: {0} -> retrying task (attempt {1}/{2})" -f $taskKey, $authRecoveryAttempts, $maxAuthRecovery)
                        $retryTask = $true
                        continue
                    }
                    Write-ExportLog -Message "Worker exiting: auth recovery failed/exhausted (task returned to queue)" -Level Error
                    $workerShouldExit = $true
                }
                elseif ($isCmdletNotFound) {
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
            }  # end retry-in-place loop

            Complete-WorkerTask -WorkerDir $workerDir
            $lastWorkerActivity = Get-Date
            if ($workerShouldExit) { break }

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
