# ContentExplorer Export — orchestrator for single/multi-terminal exports
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
                        Write-ExportLog -Message ("  Warning: Could not parse aggregate error file {0} - quarantined" -f $errFile.Name) -Level Warning -LogOnly
                        try { [System.IO.File]::Move($errFile.FullName, ($errFile.FullName + '.invalid'), $true) }
                        catch { try { Remove-Item -Path $errFile.FullName -Force -ErrorAction SilentlyContinue } catch {} }
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
                    Write-ExportLog -Message ("  Warning: Could not parse detail done file {0} - quarantined" -f $doneFile.Name) -Level Warning -LogOnly
                    try { [System.IO.File]::Move($doneFile.FullName, ($doneFile.FullName + '.invalid'), $true) }
                    catch { try { Remove-Item -Path $doneFile.FullName -Force -ErrorAction SilentlyContinue } catch {} }
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
                    Write-ExportLog -Message ("  Warning: Could not parse detail error file {0} - quarantined" -f $errFile.Name) -Level Warning -LogOnly
                    try { [System.IO.File]::Move($errFile.FullName, ($errFile.FullName + '.invalid'), $true) }
                    catch { try { Remove-Item -Path $errFile.FullName -Force -ErrorAction SilentlyContinue } catch {} }
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
                $taskData = New-CEDetailDispatchPayload -NextTask $NextTask
            }
            $sent = Send-WorkerTask -WorkerDir $Worker.WorkerDir -TaskData $taskData -ExportDir $Context.ExportDir
            if ($sent) {
                $Context.DispatchTimes[$Worker.PID] = Get-Date
            }
            return $sent
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

            # Single-terminal detail splat (fresh): no telemetry object (matches prior inline build).
            $exportParams = Build-CEDetailExportParams -Task $task -PageSize $taskPageSize `
                -ProgressLogPath $progressLogPath -TelemetryDatabasePath $telemetryDbPath `
                -OutputDirectory $classifierDir

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

