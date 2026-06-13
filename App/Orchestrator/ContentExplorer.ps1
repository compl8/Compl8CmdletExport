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
                        $newAggCookie = $metadata.PageCookie
                        if ([string]::IsNullOrEmpty($newAggCookie)) {
                            throw "MorePagesAvailable=true but PageCookie is null/empty - cannot advance aggregate cursor"
                        }
                        if ($newAggCookie -eq $pageCookie) {
                            throw "API returned same PageCookie as previous aggregate page - cursor stuck"
                        }
                        $pageCookie = $newAggCookie
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
                    $exportParams = @{
                        Task                 = $task
                        PageSize             = $cePageSize
                        ProgressLogPath      = $progressLogPath
                        Telemetry            = $telemetry
                        TelemetryDatabasePath = $telemetryDbPath
                        AdaptivePageSize     = $true
                        OutputDirectory      = $classifierDir
                    }
                    if ($task.LocationType -eq "SiteUrl" -and $task.Location) {
                        $exportParams["SiteUrl"] = $task.Location
                    } elseif ($task.LocationType -eq "UPN" -and $task.Location) {
                        $exportParams["UserPrincipalName"] = $task.Location
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
    $configPath = Join-Path $scriptRoot "ConfigFiles" "ContentExplorerClassifiers.json"
    $resolved = Resolve-CEPageSize -ExportRunDirectory $ExportDir -ConfigPath $configPath -FallbackPageSize $PageSize
    $cePageSize = $resolved.PageSize

    $progressLogPath = Join-Path (Get-LogsDir $ExportDir) "ContentExplorer-Progress.log"
    $aggregateCsvPath = Join-Path (Get-CoordinationDir $ExportDir) "ContentExplorer-Aggregates.csv"
    $trackerPath = Join-Path (Get-CoordinationDir $ExportDir) "RunTracker.json"
    $tracker = Get-ContentExplorerRunTracker -TrackerPath $trackerPath
    $telemetryDbPath = Join-Path $scriptRoot "TelemetryDB" "content-explorer-telemetry.jsonl"

    # Build work plan tasks from retry CSV (reuse aggregate data, skip aggregation)
    $script:AllWorkPlanTasks = @()
    $workloads = @($script:CEDefaultWorkloads)

    foreach ($rt in $retryTasks) {
        $rtLocation = if ($rt.Location) { $rt.Location } else { "" }
        $rtLocationType = if ($rt.LocationType) { $rt.LocationType } else { "WorkloadFallback" }

        $script:AllWorkPlanTasks += @{
            TagType               = $rt.TagType
            TagName               = $rt.TagName
            Workload              = $rt.Workload
            Location              = $rtLocation
            LocationType          = $rtLocationType
            ExpectedCount         = ($rt.OriginalExpectedCount -as [int])
            OriginalExpectedCount = ($rt.OriginalExpectedCount -as [int])
            PageSize              = if ($rt.PageSize) { ($rt.PageSize -as [int]) } else { $cePageSize }
            ExportedCount         = 0
            Status                = "Pending"
        }
    }

    Write-ExportLog -Message ("  Built {0} retry tasks" -f $script:AllWorkPlanTasks.Count) -Level Info

    # Set phase to Detail
    Write-ExportPhase -ExportDir $ExportDir -Phase "Detail"

    # Run detail export for retry tasks (single-terminal)
    $completedTaskCounts = @{}

    foreach ($task in @($script:AllWorkPlanTasks)) {
        $taskLocation = if ($task.Location) { $task.Location } else { "" }
        $taskKey = "{0}|{1}|{2}|{3}" -f $task.TagType, $task.TagName, $task.Workload, $taskLocation

        if (-not $task.ExpectedCount -or $task.ExpectedCount -eq 0) {
            Write-ExportLog -Message ("    Skipping {0} / {1} - no data" -f $task.TagName, $task.Workload) -Level Info
            continue
        }

        $locationLabel = if ($taskLocation) { " @ $taskLocation" } else { "" }
        Write-ExportLog -Message ("    Retrying: {0} / {1}{2} (expected {3})" -f $task.TagName, $task.Workload, $locationLabel, $task.ExpectedCount) -Level Info

        # Export task with progress tracking — output to Data/ContentExplorer/TagType/TagName/
        $classifierDir = Get-CEClassifierDir $ExportDir $task.TagType $task.TagName

        $taskPageSize = if ($task.PageSize -and ($task.PageSize -as [int]) -gt 0) { ($task.PageSize -as [int]) } else { $cePageSize }
        $exportParams = @{
            Task                  = $task
            PageSize              = $taskPageSize
            ProgressLogPath       = $progressLogPath
            AdaptivePageSize      = $true
            TelemetryDatabasePath = $telemetryDbPath
            OutputDirectory       = $classifierDir
        }
        if ($task.LocationType -eq "SiteUrl" -and $task.Location) {
            $exportParams["SiteUrl"] = $task.Location
        } elseif ($task.LocationType -eq "UPN" -and $task.Location) {
            $exportParams["UserPrincipalName"] = $task.Location
        }

        try {
            Export-ContentExplorerWithProgress @exportParams | Out-Null
            $exportedCount = if ($task.ExportedCount) { $task.ExportedCount } else { 0 }
            $completedTaskCounts[$taskKey] = $exportedCount

            if ($exportedCount -gt 0) {
                Write-ExportLog -Message ("    Completed: {0} / {1}{2} - {3} records" -f $task.TagName, $task.Workload, $locationLabel, $exportedCount) -Level Success
            }
            else {
                Write-ExportLog -Message ("    Completed: {0} / {1}{2} - 0 records" -f $task.TagName, $task.Workload, $locationLabel) -Level Info
            }
        }
        catch {
            Write-ExportLog -Message ("    FAILED: {0} / {1}{2} - {3}" -f $task.TagName, $task.Workload, $locationLabel, $_.Exception.Message) -Level Error
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
            $rtLocation = if ($rt.Location) { $rt.Location } else { "" }
            $taskKey = "{0}|{1}|{2}|{3}" -f $rt.TagType, $rt.TagName, $rt.Workload, $rtLocation
            if ($completedTaskCounts.ContainsKey($taskKey)) {
                $matchTask = $detailTasks | Where-Object {
                    $_.TagType -eq $rt.TagType -and $_.TagName -eq $rt.TagName -and $_.Workload -eq $rt.Workload -and $(if ($_.Location) { $_.Location } else { "" }) -eq $rtLocation
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
            -OnSelectNextTask ${function:Select-LargestPendingTask} `
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
            $exportParams = @{
                Task                  = $task
                PageSize              = if ($task.PageSize -and ($task.PageSize -as [int]) -gt 0) { ($task.PageSize -as [int]) } else { $cePageSize }
                ProgressLogPath       = $progressLogPath
                Telemetry             = $telemetry
                TelemetryDatabasePath = $telemetryDbPath
                AdaptivePageSize      = $true
                OutputDirectory       = $classifierDir
            }
            if ($task.LocationType -eq "SiteUrl" -and $task.Location) {
                $exportParams["SiteUrl"] = $task.Location
            } elseif ($task.LocationType -eq "UPN" -and $task.Location) {
                $exportParams["UserPrincipalName"] = $task.Location
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
    $exportDir = $script:SharedExportDirectory

    # ... [unchanged: aggregate caching check] ...
    # Check for recent aggregate CSVs to reuse (no more "join existing export" -- that concept
    # is removed because there is no ExportProject.json to join).
    # Instead, we only offer to reuse aggregate data from previous exports.

    # Get current tenant info for filtering and logging
    $currentTenant = Get-Compl8TenantInfo
    if ($currentTenant) {
        Write-ExportLog -Message ("Current tenant: {0} ({1})" -f $currentTenant.TenantDomain, $currentTenant.TenantId) -Level Info
    }

    # SIT reference snapshot: GUID->name map (flat list + rule packages) written to the
    # export root so downstream analytics resolve SIT/sub-entity GUIDs to display names.
    Export-SitReferenceSnapshot -ExportRunDirectory $script:ExportRunDirectory | Out-Null

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
    $ceSettings = Get-ContentExplorerSettings -ConfigObject $config -DefaultBatchSize $script:CEDefaultBatchSize -DefaultWorkloads $script:CEDefaultWorkloads -DefaultPageSize 1000
    $batchSize = $ceSettings.BatchSize
    $workloads = @($ceSettings.Workloads)
    $cePageSize = $ceSettings.PageSize
    $largeAllSITDetailThreshold = $ceSettings.LargeAllSITDetailThreshold
    $largeAllSITFallbackCandidates = @($ceSettings.LargeAllSITWorkloadFallbackWorkloads)
    $minLocationItems = $ceSettings.MinLocationItems

    # Apply menu/CLI workload selection (overrides config)
    if ($CEWorkloads) {
        $workloads = @($CEWorkloads)
    }
    # CLI threshold override (-1 = not specified; 0 = explicit "export everything")
    if ($CEMinLocationItems -ge 0) {
        $minLocationItems = $CEMinLocationItems
    }

    # Save export settings manifest for resume consistency
    Save-ExportSettings -ExportRunDirectory $script:ExportRunDirectory -ExportType "ContentExplorer" -Settings @{
        Workloads = $workloads
        CEAllSITs = [bool]$CEAllSITs
        CEAllTCs  = [bool]$CEAllTCs
        BatchSize = $batchSize
        PageSize  = $cePageSize
        MinLocationItems = $minLocationItems
        LargeAllSITDetailThreshold = $largeAllSITDetailThreshold
        LargeAllSITWorkloadFallbackWorkloads = $largeAllSITFallbackCandidates
        JsonlOutput = ($env:COMPL8_JSONL_OUTPUT -eq "1")
    }

    Write-ExportLog -Message ("Default page size: {0} (adaptive sizing selects optimal size per workload)" -f $cePageSize) -Level Info
    if ($minLocationItems -gt 0) {
        Write-ExportLog -Message ("Minimum location size: {0} item(s) - smaller locations are skipped at planning" -f $minLocationItems) -Level Info
    }
    Write-ExportLog -Message ("Workloads: {0}" -f ($workloads -join ', ')) -Level Info
    if ($CEAllSITs) {
        Write-ExportLog -Message "MODE: All SITs - Auto-discovering all Sensitive Information Types" -Level Info
    }
    elseif ($CEAllTCs) {
        Write-ExportLog -Message "MODE: All TCs - Auto-discovering all Trainable Classifiers (other tag types skipped)" -Level Info
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
        @{ TagType = "TrainableClassifier"; ConfigSection = "TrainableClassifiers"; DiscoverCmd = { Get-TrainableClassifiersFromCache } }
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
        # CEAllTCs mode: only process TrainableClassifier
        if ($CEAllTCs -and $tagType -ne "TrainableClassifier") {
            Write-ExportLog -Message ("`n--- {0} --- (skipped in All TCs mode)" -f $tagType) -Level Info
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
        elseif ($CEAllTCs -and $tagType -eq "TrainableClassifier") {
            $autoDiscover = $true
            Write-ExportLog -Message "  All TCs mode: forcing auto-discovery from cache" -Level Info
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

    # Large all-SIT runs create punishing SIT x location task fanout for mailbox-like
    # workloads. The cmdlet still requires TagType/TagName, so use one detail task per
    # SIT/workload for configured workloads instead of per-location tasks.
    # discoveredTagsByType at this point already reflects SITstoSkip filtering and the
    # empty-name filter — i.e. the SITs we will actually query, not the unfiltered list.
    $largeAllSITDetailFallbackWorkloads = @()
    if ($CEAllSITs -and $discoveredTagsByType.ContainsKey("SensitiveInformationType")) {
        $sitCountForPlanning = @($discoveredTagsByType["SensitiveInformationType"]).Count
        if ($sitCountForPlanning -gt $largeAllSITDetailThreshold) {
            $largeAllSITDetailFallbackWorkloads = @($workloads | Where-Object { $_ -in $largeAllSITFallbackCandidates })
            if ($largeAllSITDetailFallbackWorkloads.Count -gt 0) {
                Write-ExportLog -Message ("Large All-SIT detail strategy enabled: {0} SITs > threshold {1}; using workload-level detail tasks for {2}" -f $sitCountForPlanning, $largeAllSITDetailThreshold, ($largeAllSITDetailFallbackWorkloads -join ', ')) -Level Info
            }
        }
    }

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
        # Collect into a real ArrayList: the dispatch loop and the W-hotkey [ref] must
        # share this exact instance for dynamically added workers to be dispatchable.
        $workerProcesses = [System.Collections.ArrayList]@(Start-WorkerTerminals -ExportRunDirectory $script:SharedExportDirectory -Count $WorkerCount)
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
                    # Claim the worker file before reading: rename .csv -> .processing.
                    # If we later crash or Add-Content throws, the next scan won't see
                    # the same file as .csv and re-import it, so central aggregate CSV
                    # rows can't be duplicated. On success we rename .processing -> .done.
                    $processingPath = $aggFile.FullName -replace '\.csv$', '.processing'
                    try {
                        [System.IO.File]::Move($aggFile.FullName, $processingPath, $true)
                    }
                    catch {
                        Write-ExportLog -Message ("  Warning: Could not claim aggregate file {0}: {1}" -f $aggFile.Name, $_.Exception.Message) -Level Warning -LogOnly
                        continue
                    }

                    try {
                        $aggFileContent = @(Import-Csv -Path $processingPath -Encoding UTF8 -ErrorAction Stop)
                        if ($aggFileContent.Count -gt 0) {
                            $fileTagType = $aggFileContent[0].TagType
                            $fileTagName = $aggFileContent[0].TagName
                            $fileWorkload = $aggFileContent[0].Workload

                            # Check for error CSV (Location=ERROR)
                            $errorRow = $aggFileContent | Where-Object { $_.Location -eq 'ERROR' } | Select-Object -First 1
                            if ($errorRow) {
                                $errMsg = if ($errorRow.Error) { $errorRow.Error } else { "Aggregate failed (error CSV from worker)" }
                                # Write error to central aggregate CSV
                                $errorCsvLine = '{0},{1},{2},{3},ERROR,0,"{4}"' -f $errorRow.Timestamp, $fileTagType, (ConvertTo-CsvField $fileTagName), $fileWorkload, ($errMsg -replace '"', '""')
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
                                    $csvLine = "{0},{1},{2},{3},{4},{5}" -f $row.Timestamp, $row.TagType, (ConvertTo-CsvField $row.TagName), $row.Workload, $locationName, $row.Count
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
                        # Success: rename .processing -> .done. Leave .processing in
                        # place on any exception so the operator can diagnose.
                        $donePath = $aggFile.FullName -replace '\.csv$', '.done'
                        try {
                            [System.IO.File]::Move($processingPath, $donePath, $true)
                        }
                        catch {
                            Write-ExportLog -Message ("  Warning: Could not finalize aggregate file {0} -> .done: {1}" -f $aggFile.Name, $_.Exception.Message) -Level Warning -LogOnly
                        }
                    }
                    catch {
                        Write-ExportLog -Message ("  Warning: Could not process aggregate output {0} (left at .processing for diagnosis): {1}" -f $aggFile.Name, $_.Exception.Message) -Level Warning -LogOnly
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
                            $errorCsvLine = '{0},{1},{2},{3},ERROR,0,"{4}"' -f $timestamp, $errData.TagType, (ConvertTo-CsvField $errData.TagName), $errData.Workload, ($errData.Error -replace '"', '""')
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

            if ($CompletedTask.Phase -eq "Detail") {
                # Preserve the original expected count, overwrite ExpectedCount with
                # the actual exported record count. This matches single-terminal
                # behavior (see Invoke-ContentExplorerResume / Invoke-ContentExplorerExport
                # single-pass loop) and is what Get-RetryBucketTasks expects when
                # reading the completed DetailTasks.csv back in.
                if (-not $CompletedTask.OriginalExpectedCount -or ($CompletedTask.OriginalExpectedCount -as [int]) -eq 0) {
                    $CompletedTask.OriginalExpectedCount = $CompletedTask.ExpectedCount
                }
                $actualCount = 0
                if ($CompletionData -and $CompletionData.ContainsKey('RecordCount')) {
                    $actualCount = $CompletionData.RecordCount -as [int]
                    if ($null -eq $actualCount) { $actualCount = 0 }
                }
                $CompletedTask.ExpectedCount = $actualCount

                # Populate the summary/aggregate-progress map so the multi-terminal
                # summary (Get-ContentExplorerAggregateProgress at end of run) sees
                # the same per-task counts that the single-terminal path tracks.
                if ($Context.CompletedTaskCounts) {
                    $locKey = if ($CompletedTask.Location) { $CompletedTask.Location } else { "" }
                    $taskKey = "{0}|{1}|{2}|{3}" -f $CompletedTask.TagType, $CompletedTask.TagName, $CompletedTask.Workload, $locKey
                    $Context.CompletedTaskCounts[$taskKey] = $actualCount
                }
                return @()
            }

            # Only generate detail tasks for aggregate completions
            if ($CompletedTask.Phase -ne "Aggregate") { return @() }

            $newTasks = @()
            $aggCsvPath = $Context.AggregateCsvPath
            $cePageSize = $Context.DefaultPageSize
            $detailFallbackWorkloads = @($Context.DetailWorkloadFallbackWorkloads)
            $useWorkloadFallbackDetail = ($detailFallbackWorkloads.Count -gt 0 -and $CompletedTask.Workload -in $detailFallbackWorkloads)

            $totalCount = $CompletedTask.ExpectedCount -as [int]

            # Skip zero-count aggregates (no data to export)
            if (-not $totalCount -or $totalCount -le 0) {
                Write-ExportLog -Message ("  Aggregate {0}/{1} completed with 0 files - skipping detail tasks" -f $CompletedTask.TagName, $CompletedTask.Workload) -Level Info -LogOnly
                return @()
            }

            try {
                $importResult = Import-AggregateDataFromCsv -CsvPath $aggCsvPath `
                    -TagType $CompletedTask.TagType -TagNames @($CompletedTask.TagName) -Workloads @($CompletedTask.Workload)

                $newTasks = @(New-ContentExplorerDetailTasks `
                    -WorkPlanTasks @($importResult.TaskData.Values) `
                    -DefaultPageSize $cePageSize `
                    -WorkloadFallbackWorkloads $detailFallbackWorkloads `
                    -MinLocationItems ([int]$Context.MinLocationItems))

                if ($useWorkloadFallbackDetail) {
                    foreach ($taskData in @($importResult.TaskData.Values)) {
                        if ($taskData.Locations -and @($taskData.Locations).Count -gt 0) {
                            Write-ExportLog -Message ("  Large All-SIT detail strategy: generated one workload-level detail task for {0}/{1} instead of {2} location tasks" -f $taskData.TagName, $taskData.Workload, @($taskData.Locations).Count) -Level Info -LogOnly
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
                            # Splat the hashtable by named keys. @($Context.AuthParams)
                            # wraps it in a one-element array and splats positionally,
                            # so cert-auth params were effectively dropped.
                            $keepaliveAuthParams = $Context.AuthParams
                            Connect-Compl8Compliance @keepaliveAuthParams -LogOnly
                            $Context.LastKeepalive = Get-Date
                        }
                    } catch {
                        Write-ExportLog -Message ("  Session reconnection failed (non-fatal): {0}" -f $_.Exception.Message) -Level Warning -LogOnly
                        $Context.LastKeepalive = Get-Date
                    }
                }
            }

            # Dynamic worker spawning via W hotkey. Pass the original [ref]s straight
            # through: re-wrapping with [ref]$Context.X.Value creates a ref to a temporary,
            # so NextWorkerNumber increments would be silently lost.
            Test-AddWorkerKeypress -ExportRunDirectory $Context.ExportRunDirectory `
                -WorkerProcesses $Context.WorkerProcessesRef -NextWorkerNumber $Context.NextWorkerNumberRef
            # WorkerProcessesRef wraps the same ArrayList instance the dispatch engine
            # iterates, so workers added here are dispatchable on the next iteration.

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
            MinLocationItems      = $minLocationItems
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
            DetailWorkloadFallbackWorkloads = @($largeAllSITDetailFallbackWorkloads)
            CompletedTaskCounts   = $completedTaskCounts
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
            -OnSelectNextTask ${function:Select-LargestPendingTask} `
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

    if (-not $isMultiTerminal) {
        $script:AllWorkPlanTasks = @(New-ContentExplorerDetailTasks `
            -WorkPlanTasks @($script:AllWorkPlanTasks) `
            -DefaultPageSize $cePageSize `
            -WorkloadFallbackWorkloads $largeAllSITDetailFallbackWorkloads `
            -MinLocationItems $minLocationItems `
            -Sort)

        $singleDetailCsvPath = Join-Path (Get-CoordinationDir $script:SharedExportDirectory) "DetailTasks.csv"
        if ($script:AllWorkPlanTasks.Count -gt 0) {
            Write-TaskCsv -Path $singleDetailCsvPath -Tasks $script:AllWorkPlanTasks
            Write-ExportLog -Message ("  Planned {0} single-terminal detail tasks" -f $script:AllWorkPlanTasks.Count) -Level Info
        }
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

        $singleDetailCsvPath = Join-Path (Get-CoordinationDir $script:SharedExportDirectory) "DetailTasks.csv"
        $detailCsvFlushInterval = 10
        $tasksSinceFlush = 0

        foreach ($task in @($script:AllWorkPlanTasks)) {
            $taskLocation = if ($task.Location) { $task.Location } else { "" }
            $taskKey = "{0}|{1}|{2}|{3}" -f $task.TagType, $task.TagName, $task.Workload, $taskLocation

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
            $taskPageSize = if ($task.PageSize -and ($task.PageSize -as [int]) -gt 0) { ($task.PageSize -as [int]) } elseif ($task.Status -eq "Error") { [Math]::Max(500, $cePageSize) } else { $cePageSize }
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
            if ($task.LocationType -eq "SiteUrl" -and $task.Location) {
                $exportParams["SiteUrl"] = $task.Location
            } elseif ($task.LocationType -eq "UPN" -and $task.Location) {
                $exportParams["UserPrincipalName"] = $task.Location
            }

            $exportFailed = $false
            try {
                Export-ContentExplorerWithProgress @exportParams | Out-Null
            }
            catch {
                $exportFailed = $true
                Write-ExportLog -Message ("    FAILED: {0} / {1} - {2}" -f $task.TagName, $task.Workload, $_.Exception.Message) -Level Error
                if ($script:ErrorLogPath) {
                    Write-ExportErrorLog -ErrorLogPath $script:ErrorLogPath -Context "Single-Terminal Detail Export" -TaskKey $taskKey -ErrorRecord $_
                }
                # Flush CSV immediately on error so external observers see the failure state
                if (Test-Path $singleDetailCsvPath) {
                    Write-TaskCsv -Path $singleDetailCsvPath -Tasks $script:AllWorkPlanTasks
                    $tasksSinceFlush = 0
                }
            }

            # Track exported count for this task
            $exportedCount = if ($task.ExportedCount) { $task.ExportedCount } else { 0 }
            $completedTaskCounts[$taskKey] = $exportedCount
            if (-not $task.OriginalExpectedCount -or ($task.OriginalExpectedCount -as [int]) -eq 0) {
                $task.OriginalExpectedCount = $task.ExpectedCount
            }
            $task.ExpectedCount = $exportedCount

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

            # Throttle DetailTasks.csv flushes: every 10 tasks + on error (above) + at end (below).
            # The run tracker JSON is the source of truth for resume; this CSV is for observers.
            $tasksSinceFlush++
            if ($tasksSinceFlush -ge $detailCsvFlushInterval -and (Test-Path $singleDetailCsvPath)) {
                Write-TaskCsv -Path $singleDetailCsvPath -Tasks $script:AllWorkPlanTasks
                $tasksSinceFlush = 0
            }
        }

        # Final flush after the loop
        if ($tasksSinceFlush -gt 0 -and (Test-Path $singleDetailCsvPath)) {
            Write-TaskCsv -Path $singleDetailCsvPath -Tasks $script:AllWorkPlanTasks
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

    # Aggregate-delta + watermark save (foundation for incremental support)
    try {
        $detailCsvPathFinal = Join-Path (Get-CoordinationDir $exportDir) "DetailTasks.csv"
        if (Test-Path $detailCsvPathFinal) {
            $finalDetailTasks = @(Read-TaskCsv -Path $detailCsvPathFinal)

            # Write aggregate-delta report comparing this run's counts to prior watermarks
            $priorWatermarks = Read-Watermarks -ScriptRoot $scriptRoot -TenantPrefix $script:TenantPrefix
            $deltaInput = @($finalDetailTasks | ForEach-Object {
                [PSCustomObject]@{
                    TagType        = $_.TagType
                    TagName        = $_.TagName
                    Workload       = $_.Workload
                    Location       = $_.Location
                    ExpectedCount  = if ($_.OriginalExpectedCount) { $_.OriginalExpectedCount } else { $_.ExpectedCount }
                }
            })
            Write-AggregateDeltaReport -ExportDir $exportDir -Watermarks $priorWatermarks -AggregateTasks $deltaInput

            # Persist new watermarks for the next run
            $wasFullRun = -not ($env:COMPL8_INCREMENTAL -eq "1") -or ($env:COMPL8_FORCE_FULL_REBUILD -eq "1")
            Save-WatermarksFromDetailTasks -ScriptRoot $scriptRoot -TenantPrefix $script:TenantPrefix -DetailTasks $finalDetailTasks -WasFullRun:$wasFullRun
            Write-ExportLog -Message "Saved tenant watermarks for future incremental runs" -Level Info
        }
    }
    catch {
        Write-ExportLog -Message ("Watermark/delta processing failed: {0}" -f $_.Exception.Message) -Level Warning
    }

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

