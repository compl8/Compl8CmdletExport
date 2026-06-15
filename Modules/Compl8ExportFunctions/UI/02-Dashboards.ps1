#region Dashboard Functions

function Reset-OrchestratorDashboard {
    <#
    .SYNOPSIS
        Resets the Content Explorer orchestrator dashboard line counter.
        Call before entering a new dashboard loop to prevent cursor misalignment.
    #>
    $script:DashboardLineCount = 0
    # Drop any ETA sampler history from a prior loop so the recent-rate window
    # measures only this run/phase (see Get-ProgressEta).
    $script:CeAggEtaState = @{}
    $script:CeDetailEtaState = @{}
}

function Reset-AEDashboard {
    <#
    .SYNOPSIS
        Resets the Activity Explorer dashboard line counter.
        Call before entering a new dashboard loop to prevent cursor misalignment.
    #>
    $script:AEDashboardLineCount = 0
    $script:AeEtaState = @{}
}

function Get-ProgressEta {
    <#
    .SYNOPSIS
        Estimates time-to-completion from a recent (trailing-window) throughput
        rate, rather than a lifetime cumulative average.

    .DESCRIPTION
        The old dashboard ETA divided total-elapsed by total-completed -- a
        cumulative average. That is badly biased on this tool's bursty,
        non-stationary throughput: worker ramp-up/auth, Security & Compliance
        429/503 throttle waves (60s backoffs), and live worker-count changes
        (the [W] add-worker hotkey, or resume with fewer workers). A lifetime
        average bakes in the slow start forever and barely reacts when a throttle
        wave hits or clears, so the ETA is wrong in both directions.

        This measures the rate over a trailing time window and lightly smooths it
        (EWMA) so the estimate tracks the *current* pace. All history lives in the
        caller-supplied -State hashtable (mutated in place) and "now" is injected,
        so the math is deterministic and unit-testable; the dashboards keep one
        -State per phase in script scope.

        Fallback ladder:
          * 'window'     - trailing-window rate, once there is >= MinSpanSeconds
                           of recent movement (the normal case).
          * 'cumulative' - before that, a rate over the retained samples (matches
                           the old behaviour during the first ~MinSpanSeconds).
          * 'none'       - nothing has moved yet; Ready is $false and callers may
                           show a seed estimate instead.

    .PARAMETER State
        Hashtable persisted across calls. Keys 'Samples' (a generic List of
        timestamp/units points) and 'Ewma' are created and updated in place.

    .PARAMETER Now
        Current time, injected for testability.

    .PARAMETER CompletedUnits
        Cumulative completed work in the phase's natural unit (tasks, items, or
        percent). The rate is derived from the change in this value over time, so
        a constant resume baseline is harmless.

    .PARAMETER RemainingUnits
        Work left, same unit. ETA = RemainingUnits / rate.

    .PARAMETER WindowSeconds
        Trailing window for the recent-rate measurement (default 120s).

    .PARAMETER MinSpanSeconds
        Minimum span of recent movement before the windowed rate is trusted
        (default 20s); below this the cumulative fallback is used.

    .PARAMETER EwmaAlpha
        Smoothing factor for the windowed rate (0..1; higher = more responsive).

    .PARAMETER MaxEtaSeconds
        Clamp so a near-zero rate cannot print an absurd ETA (default 30 days).

    .OUTPUTS
        PSCustomObject with EtaSeconds, RatePerSecond, RatePerMinute,
        Source ('window'|'cumulative'|'none'), and Ready (bool).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$State,
        [Parameter(Mandatory)][datetime]$Now,
        [Parameter(Mandatory)][double]$CompletedUnits,
        [Parameter(Mandatory)][double]$RemainingUnits,
        [double]$WindowSeconds = 120,
        [double]$MinSpanSeconds = 20,
        [double]$EwmaAlpha = 0.4,
        [double]$MaxEtaSeconds = 2592000
    )

    if (-not $State.ContainsKey('Samples') -or $null -eq $State['Samples']) {
        $State['Samples'] = [System.Collections.Generic.List[object]]::new()
    }
    if (-not $State.ContainsKey('Ewma')) { $State['Ewma'] = $null }
    $samples = $State['Samples']

    # Record this frame, then drop anything older than the window (always keep the
    # newest point). The oldest retained sample is therefore within WindowSeconds,
    # so the measured span never exceeds the window. Bounded memory.
    $samples.Add([pscustomobject]@{ T = $Now; Units = $CompletedUnits })
    $cutoff = $Now.AddSeconds(-$WindowSeconds)
    while ($samples.Count -gt 1 -and $samples[0].T -lt $cutoff) {
        $samples.RemoveAt(0)
    }

    $oldest = $samples[0]
    $spanSec = ($Now - $oldest.T).TotalSeconds
    $deltaUnits = $CompletedUnits - $oldest.Units

    $rate = $null
    $source = 'none'
    if ($spanSec -ge $MinSpanSeconds -and $deltaUnits -gt 0) {
        $windowRate = $deltaUnits / $spanSec
        if ($null -eq $State['Ewma']) {
            $State['Ewma'] = $windowRate
        }
        else {
            $State['Ewma'] = ($EwmaAlpha * $windowRate) + ((1 - $EwmaAlpha) * $State['Ewma'])
        }
        $rate = $State['Ewma']
        $source = 'window'
    }
    elseif ($spanSec -ge $MinSpanSeconds -and $null -ne $State['Ewma']) {
        # Confirmed stall: >= MinSpan of retained window with NO forward progress
        # ($deltaUnits <= 0, having fallen through the window branch). Blend a zero
        # rate into the EWMA so it decays toward the true (zero) throughput instead of
        # freezing at the pre-stall value. Without this, the first sample after the
        # stall blends against a stale, optimistic rate and the "last Nm" ETA briefly
        # understates a throttle wave -- the exact case this estimator exists to track.
        $State['Ewma'] = (1 - $EwmaAlpha) * $State['Ewma']
        $rate = $State['Ewma']
        $source = 'window'
    }
    elseif ($spanSec -gt 0 -and $deltaUnits -gt 0) {
        # Warm-up: not enough recent span yet. Use a cumulative rate over the
        # retained samples and seed the EWMA so the window handoff is smooth.
        $rate = $deltaUnits / $spanSec
        $State['Ewma'] = $rate
        $source = 'cumulative'
    }

    $etaSeconds = $null
    $ready = $false
    if ($RemainingUnits -le 0) {
        $etaSeconds = 0.0
        $ready = $true
    }
    elseif ($null -ne $rate -and $rate -gt 0) {
        $etaSeconds = $RemainingUnits / $rate
        if ($etaSeconds -gt $MaxEtaSeconds) { $etaSeconds = $MaxEtaSeconds }
        $ready = $true
    }

    [pscustomobject]@{
        EtaSeconds    = $etaSeconds
        RatePerSecond = $rate
        RatePerMinute = if ($null -ne $rate) { $rate * 60 } else { $null }
        Source        = $source
        Ready         = $ready
    }
}

function Show-OrchestratorDashboard {
    <#
    .SYNOPSIS
        Displays a compact progress dashboard for the orchestrator.
        Redraws in place using cursor positioning for a static display.
    #>
    param(
        [string]$Phase,
        [int]$Completed,
        [int]$Total,
        [array]$Workers,
        [array]$RecentErrors,
        [array]$RecentActivity,
        [array]$DispatchLog,
        [datetime]$ExportStartTime,
        [datetime]$PhaseStartTime,
        [long]$CompletedItems = 0,
        [long]$TotalItems = 0,
        [int]$RemainingAggregates = 0,
        [array]$DetailTasks = @(),
        [hashtable]$ClassifierGroups = @{},
        [int]$TotalLocations = 0,
        [int]$TotalCompleted = 0,
        [int]$TotalErrors = 0,
        [int]$TotalActive = 0,
        [int]$CompletedBaseline = 0,
        [long]$CompletedItemsBaseline = 0
    )

    $pct = if ($Total -gt 0) { [Math]::Round(($Completed / $Total) * 100, 1) } else { 0 }
    $bar = Format-ProgressBar -Percent $pct
    $width = (Get-TerminalSize).Width

    # Build all output lines with their colors
    $lines = [System.Collections.ArrayList]::new()

    [void]$lines.Add(@{ Text = ""; Color = "White" })
    [void]$lines.Add(@{ Text = ("  Content Explorer - Phase: {0} [{1}/{2}] {3} {4}%" -f $Phase.ToUpper(), $Completed, $Total, $bar, $pct); Color = "Cyan" })

    if ($RemainingAggregates -gt 0) {
        [void]$lines.Add(@{ Text = ("  ** {0} aggregate task(s) completing in background **" -f $RemainingAggregates); Color = "Yellow" })
    }

    # Timing line: export start, phase elapsed, ETA
    $now = Get-Date
    $exportElapsed = if ($ExportStartTime -ne [datetime]::MinValue) { $now - $ExportStartTime } else { $null }
    $phaseElapsed = if ($PhaseStartTime -ne [datetime]::MinValue) { $now - $PhaseStartTime } else { $null }

    $timingParts = @()
    if ($ExportStartTime -ne [datetime]::MinValue) {
        $timingParts += "Started: {0}" -f $ExportStartTime.ToString("HH:mm:ss")
    }
    if ($phaseElapsed) {
        $timingParts += "Phase: {0}" -f (Format-TimeSpan -Seconds $phaseElapsed.TotalSeconds)
    }
    if ($exportElapsed) {
        $timingParts += "Total: {0}" -f (Format-TimeSpan -Seconds $exportElapsed.TotalSeconds)
    }

    if ($timingParts.Count -gt 0) {
        [void]$lines.Add(@{ Text = ("  {0}" -f ($timingParts -join "  |  ")); Color = "DarkGray" })
    }

    # ETA — windowed (recent-rate) throughput via Get-ProgressEta, not a lifetime
    # average, so the estimate tracks the current pace under bursty throttling and
    # live worker-count changes. Session-only completions gate the warm-up; the
    # rate itself comes from movement in the windowed sampler.
    $sessionCompleted = $Completed - $CompletedBaseline
    $sessionCompletedItems = $CompletedItems - $CompletedItemsBaseline
    if (-not $script:CeAggEtaState) { $script:CeAggEtaState = @{} }
    if (-not $script:CeDetailEtaState) { $script:CeDetailEtaState = @{} }
    $etaText = $null
    $winLabel = "last 2m"

    if ($Phase -eq "Aggregate") {
        # Units = tasks. Sample every frame while the phase is incomplete -- even before
        # the first completion -- so the baseline frame is seeded at phase start and the
        # first completed task already has something to measure against. Keep the DISPLAY
        # decision separate: only show an ETA once there is real session progress.
        if ($Completed -lt $Total) {
            $eta = Get-ProgressEta -State $script:CeAggEtaState -Now $now `
                -CompletedUnits ([double]$Completed) -RemainingUnits ([double]($Total - $Completed))
            if ($sessionCompleted -gt 0 -and $eta.Ready -and $null -ne $eta.RatePerMinute) {
                $src = if ($eta.Source -eq 'window') { $winLabel } else { 'avg' }
                $etaText = "ETA: {0}  ({1:N1} tasks/min · {2})" -f (Format-TimeSpan -Seconds $eta.EtaSeconds), $eta.RatePerMinute, $src
            }
        }
    }
    elseif ($Phase -eq "Detail") {
        # Units = items (records). The windowed rate replaces the old seed/measured
        # blend; the 40s-per-1000-items seed only covers the no-items-yet warm-up.
        if ($TotalItems -gt 0 -and $CompletedItems -lt $TotalItems) {
            $remainingItems = [double]($TotalItems - $CompletedItems)
            $eta = Get-ProgressEta -State $script:CeDetailEtaState -Now $now `
                -CompletedUnits ([double]$CompletedItems) -RemainingUnits $remainingItems
            $itemPct = [Math]::Round(($CompletedItems / $TotalItems) * 100, 1)
            if ($eta.Ready -and $null -ne $eta.RatePerMinute) {
                $src = if ($eta.Source -eq 'window') { $winLabel } else { 'avg' }
                $etaText = "ETA: {0}  ({1:N0}/{2:N0} items, {3}%  ·  {4:N0} items/min · {5})" -f (Format-TimeSpan -Seconds $eta.EtaSeconds), $CompletedItems, $TotalItems, $itemPct, $eta.RatePerMinute, $src
            }
            else {
                $etaSeconds = 0.04 * $remainingItems   # 40s / 1000 items seed
                $etaText = "ETA: ~{0}  (initial estimate, {1:N0} items)" -f (Format-TimeSpan -Seconds $etaSeconds), $TotalItems
            }
        }
    }

    if ($etaText) {
        [void]$lines.Add(@{ Text = ("  {0}" -f $etaText); Color = "Yellow" })
    }

    [void]$lines.Add(@{ Text = ("  Updated: {0}  [W] Add worker" -f (Get-Date -Format "HH:mm:ss")); Color = "DarkGray" })
    [void]$lines.Add(@{ Text = ""; Color = "White" })

    # --- Aggregated Detail Progress by Classifier ---
    if ($Phase -eq "Detail" -and (($ClassifierGroups.Count -gt 0) -or ($DetailTasks -and $DetailTasks.Count -gt 0))) {
        # Use pre-computed classifier groups if provided, otherwise build from DetailTasks
        if ($ClassifierGroups.Count -gt 0) {
            $classifierGroups = $ClassifierGroups
        } else {
            # Group tasks by TagName + Workload
            $classifierGroups = @{}
            foreach ($dt in $DetailTasks) {
                $groupKey = "{0} / {1}" -f $dt.TagName, $dt.Workload
                if (-not $classifierGroups.ContainsKey($groupKey)) {
                    $classifierGroups[$groupKey] = @{
                        TagName        = $dt.TagName
                        Workload       = $dt.Workload
                        Completed      = 0
                        InProgress     = 0
                        Pending        = 0
                        Error          = 0
                        Total          = 0
                        TotalFiles     = [long]0
                        CompletedFiles = [long]0
                        IsFallback     = $false
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
                if ($dt.LocationType -eq "WorkloadFallback") {
                    $classifierGroups[$groupKey].IsFallback = $true
                }
            }
        }

        # Sort by total file count descending (largest classifiers first), top 10
        $sortedGroups = @($classifierGroups.GetEnumerator() | Sort-Object { $_.Value.TotalFiles } -Descending)
        $maxClassifierRows = 10

        # Compute column width for alignment
        $displayGroups = if ($sortedGroups.Count -le $maxClassifierRows) { $sortedGroups } else { @($sortedGroups | Select-Object -First ($maxClassifierRows - 1)) }
        $maxKeyLen = 0
        foreach ($g in $displayGroups) {
            $prefix = if ($g.Value.IsFallback) { "[Fallback] " } else { "" }
            $keyLen = $prefix.Length + $g.Key.Length
            if ($keyLen -gt $maxKeyLen) { $maxKeyLen = $keyLen }
        }
        $maxKeyLen = [Math]::Max($maxKeyLen, 20)
        $maxKeyLen = [Math]::Min($maxKeyLen, 50)

        $hBar = [string][char]0x2500
        [void]$lines.Add(@{ Text = "  $($hBar * 2) Detail Progress by Classifier (top 10 by file count) $($hBar * 10)"; Color = "DarkCyan" })

        # Use pre-computed totals if provided, otherwise count from DetailTasks
        if ($TotalLocations -gt 0) {
            $totalLocations = $TotalLocations
            $totalCompleted = $TotalCompleted
            $totalErrors = $TotalErrors
            $totalActive = $TotalActive
        } else {
            $totalLocations = $DetailTasks.Count
            $totalCompleted = @($DetailTasks | Where-Object { $_.Status -eq "Completed" }).Count
            $totalErrors = @($DetailTasks | Where-Object { $_.Status -eq "Error" }).Count
            $totalActive = @($DetailTasks | Where-Object { $_.Status -eq "InProgress" }).Count
        }

        $displayCount = 0
        foreach ($g in $displayGroups) {
            $grp = $g.Value
            $prefix = if ($grp.IsFallback) { "[Fallback] " } else { "" }
            $label = "{0}{1}" -f $prefix, $g.Key
            $padLabel = $label.PadRight($maxKeyLen)

            # File progress
            $filePct = if ($grp.TotalFiles -gt 0) { [Math]::Round(($grp.CompletedFiles / $grp.TotalFiles) * 100, 1) } else { 0 }
            # Location progress
            $locPct = if ($grp.Total -gt 0) { [Math]::Round(($grp.Completed / $grp.Total) * 100, 1) } else { 0 }

            if ($grp.Completed -eq $grp.Total -and $grp.Error -eq 0) {
                $detail = "done  {0:N0} files, {1:N0} locs" -f $grp.TotalFiles, $grp.Total
                $lineColor = "DarkGreen"
            }
            elseif ($grp.Total -gt 0) {
                $detail = "{0:N0}/{1:N0} files ({2}%)  {3:N0}/{4:N0} locs ({5}%)" -f $grp.CompletedFiles, $grp.TotalFiles, $filePct, $grp.Completed, $grp.Total, $locPct
                if ($grp.InProgress -gt 0) {
                    $detail += "  [{0} active]" -f $grp.InProgress
                }
                if ($grp.Error -gt 0) {
                    $detail += "  [{0} err]" -f $grp.Error
                }
                $lineColor = if ($grp.InProgress -gt 0) { "Green" } else { "Gray" }
            }
            else {
                $detail = "0 files, 0 locs"
                $lineColor = "Gray"
            }

            [void]$lines.Add(@{ Text = ("    {0}  {1}" -f $padLabel, $detail); Color = $lineColor })
            $displayCount++
        }

        # Summary line for remaining groups not shown
        if ($sortedGroups.Count -gt $maxClassifierRows) {
            $remainingCount = $sortedGroups.Count - ($maxClassifierRows - 1)
            $remainingFiles = [long]0
            $remainingLocs = 0
            foreach ($g in @($sortedGroups | Select-Object -Skip ($maxClassifierRows - 1))) {
                $remainingFiles += $g.Value.TotalFiles
                $remainingLocs += $g.Value.Total
            }
            [void]$lines.Add(@{ Text = ("    ... and {0} more classifiers ({1:N0} files, {2:N0} locations)" -f $remainingCount, $remainingFiles, $remainingLocs); Color = "DarkGray" })
        }

        # Total summary line
        $activeWorkerCount = if ($Workers) { @($Workers | Where-Object { $_.State -eq "Busy" }).Count } else { 0 }
        [void]$lines.Add(@{ Text = ("  Total: {0:N0}/{1:N0} locations completed | {2} errors | {3} active workers" -f $totalCompleted, $totalLocations, $totalErrors, $activeWorkerCount); Color = "Cyan" })
        [void]$lines.Add(@{ Text = ""; Color = "White" })
    }

    if ($Workers -and $Workers.Count -gt 0) {
        [void]$lines.Add(@{ Text = "  Workers:"; Color = "Gray" })
        # Determine if extended columns are available (detail phase provides them)
        $hasExtended = $Workers[0].ContainsKey('Expected')
        if ($hasExtended) {
            [void]$lines.Add(@{ Text = ("    {0,-8} {1,-10} {2,-10} {3,-10} {4,-16} {5,-8} {6,-8} {7}" -f "PID", "Status", "Time", "Expected", "Progress", "PgSize", "PgTime", "Current Task"); Color = "DarkGray" })
        }
        else {
            [void]$lines.Add(@{ Text = ("    {0,-8} {1,-15} {2}" -f "PID", "Status", "Current Task"); Color = "DarkGray" })
        }
        foreach ($w in $Workers) {
            $color = switch ($w.State) {
                "Busy"    { "Green" }
                "Idle"    { "Yellow" }
                "Dead"    { "Red" }
                default   { "Gray" }
            }
            if ($hasExtended) {
                $timeCol = if ($w.TaskTime) { $w.TaskTime } else { "-" }
                $expCol = if ($w.Expected) { $w.Expected } else { "-" }
                $progCol = if ($w.Progress) { $w.Progress } else { "-" }
                $pgCol = if ($w.PageSize) { $w.PageSize } else { "-" }
                $ptCol = if ($w.LastPage) { $w.LastPage } else { "-" }
                [void]$lines.Add(@{ Text = ("    {0,-8} {1,-10} {2,-10} {3,-10} {4,-16} {5,-8} {6,-8} {7}" -f $w.PID, $w.State, $timeCol, $expCol, $progCol, $pgCol, $ptCol, ($w.CurrentTask ?? "-")); Color = $color })
            }
            else {
                [void]$lines.Add(@{ Text = ("    {0,-8} {1,-15} {2}" -f $w.PID, $w.State, ($w.CurrentTask ?? "-")); Color = $color })
            }
        }
    }

    if ($DispatchLog -and $DispatchLog.Count -gt 0) {
        [void]$lines.Add(@{ Text = ""; Color = "White" })
        $dispatchSlice = @($DispatchLog | Select-Object -Last 4)
        [void]$lines.Add(@{ Text = ("  Dispatch Log ({0} total):" -f $DispatchLog.Count); Color = "DarkMagenta" })
        foreach ($d in $dispatchSlice) {
            [void]$lines.Add(@{ Text = ("    {0}  PID {1,-6} -> {2}" -f $d.Time, $d.PID, $d.Task); Color = "DarkGray" })
        }
    }

    if ($RecentActivity -and $RecentActivity.Count -gt 0) {
        [void]$lines.Add(@{ Text = ""; Color = "White" })
        $activitySlice = @($RecentActivity | Select-Object -Last 4)
        [void]$lines.Add(@{ Text = ("  Recent Activity ({0} total):" -f $RecentActivity.Count); Color = "DarkCyan" })
        foreach ($act in $activitySlice) {
            [void]$lines.Add(@{ Text = ("    {0}  {1}" -f $act.Time, $act.Message); Color = "DarkGray" })
        }
    }

    if ($RecentErrors -and $RecentErrors.Count -gt 0) {
        [void]$lines.Add(@{ Text = ""; Color = "White" })
        $errorSlice = @($RecentErrors | Select-Object -Last 4)
        [void]$lines.Add(@{ Text = ("  Recent Errors ({0} total):" -f $RecentErrors.Count); Color = "Red" })
        foreach ($err in $errorSlice) {
            [void]$lines.Add(@{ Text = ("    {0}  {1}" -f $err.Time, $err.Message); Color = "DarkRed" })
        }
    }
    [void]$lines.Add(@{ Text = ""; Color = "White" })

    $script:DashboardLineCount = Write-DashboardFrame -Lines $lines -PreviousLineCount $(if ($script:DashboardLineCount) { $script:DashboardLineCount } else { 0 })
}

function Show-AEDashboard {
    <#
    .SYNOPSIS
        Displays a compact progress dashboard for Activity Explorer multi-terminal export.
        Redraws in place using cursor positioning for a static display.
    #>
    param(
        [string]$Phase,
        [int]$Completed,
        [int]$Total,
        [array]$Workers,
        [array]$DayTasks,
        [array]$RecentActivity,
        [array]$RecentErrors,
        [datetime]$ExportStartTime,
        [long]$TotalRecords = 0,
        [double]$WeightedPct = -1
    )

    $width = (Get-TerminalSize).Width

    $lines = [System.Collections.ArrayList]::new()

    # Progress bar: use weighted percentage (per-day record progress) when available
    [void]$lines.Add(@{ Text = ""; Color = "White" })
    if ($WeightedPct -ge 0) {
        $pct = [Math]::Round($WeightedPct, 1)
        $bar = Format-ProgressBar -Percent $pct
        $recText = if ($TotalRecords -gt 0) { "  ({0:N0} records)" -f $TotalRecords } else { "" }
        [void]$lines.Add(@{ Text = ("  Activity Explorer - [{0}/{1} days] {2} {3}%{4}" -f $Completed, $Total, $bar, $pct, $recText); Color = "Cyan" })
    }
    else {
        $pct = if ($Total -gt 0) { [Math]::Round(($Completed / $Total) * 100, 1) } else { 0 }
        $bar = Format-ProgressBar -Percent $pct
        [void]$lines.Add(@{ Text = ("  Activity Explorer - [{0}/{1} days] {2} {3}%" -f $Completed, $Total, $bar, $pct); Color = "Cyan" })
    }

    # Timing + ETA
    $now = Get-Date
    $elapsed = if ($ExportStartTime -ne [datetime]::MinValue) { $now - $ExportStartTime } else { $null }
    $timingParts = @()
    if ($ExportStartTime -ne [datetime]::MinValue) {
        $timingParts += "Started: {0}" -f $ExportStartTime.ToString("HH:mm:ss")
    }
    if ($elapsed) {
        $timingParts += "Elapsed: {0}" -f (Format-TimeSpan -Seconds $elapsed.TotalSeconds)
    }

    # ETA — windowed weighted-progress rate (recent %/min) via Get-ProgressEta,
    # not a lifetime average, so it tracks the current per-day pace.
    if (-not $script:AeEtaState) { $script:AeEtaState = @{} }
    $etaText = $null
    if ($pct -gt 0 -and $pct -lt 100) {
        $eta = Get-ProgressEta -State $script:AeEtaState -Now $now `
            -CompletedUnits ([double]$pct) -RemainingUnits ([double](100 - $pct))
        if ($eta.Ready -and $null -ne $eta.RatePerMinute) {
            $src = if ($eta.Source -eq 'window') { 'last 2m' } else { 'avg' }
            $etaText = "ETA: {0}  ({1:N1}%/min · {2})" -f (Format-TimeSpan -Seconds $eta.EtaSeconds), $eta.RatePerMinute, $src
        }
    }

    if ($timingParts.Count -gt 0) {
        [void]$lines.Add(@{ Text = ("  {0}" -f ($timingParts -join "  |  ")); Color = "DarkGray" })
    }
    if ($etaText) {
        [void]$lines.Add(@{ Text = ("  {0}" -f $etaText); Color = "Yellow" })
    }

    [void]$lines.Add(@{ Text = ("  Updated: {0}" -f (Get-Date -Format "HH:mm:ss")); Color = "DarkGray" })
    [void]$lines.Add(@{ Text = ""; Color = "White" })

    # Task status summary (compact: one line instead of full day table)
    if ($DayTasks -and $DayTasks.Count -gt 0) {
        $completedDays = @($DayTasks | Where-Object { $_.Status -eq "Completed" }).Count
        $activeDays = @($DayTasks | Where-Object { $_.Status -eq "InProgress" }).Count
        $pendingDays = @($DayTasks | Where-Object { $_.Status -eq "Pending" }).Count
        $errorDays = @($DayTasks | Where-Object { $_.Status -eq "Error" }).Count
        $summaryParts = @()
        if ($completedDays -gt 0) { $summaryParts += "Completed: $completedDays" }
        if ($activeDays -gt 0) { $summaryParts += "Active: $activeDays" }
        if ($pendingDays -gt 0) { $summaryParts += "Pending: $pendingDays" }
        if ($errorDays -gt 0) { $summaryParts += "Errors: $errorDays" }
        [void]$lines.Add(@{ Text = ("  {0}" -f ($summaryParts -join "  |  ")); Color = "White" })
        [void]$lines.Add(@{ Text = ""; Color = "White" })
    }

    # Worker status
    if ($Workers -and $Workers.Count -gt 0) {
        $hBar = [string][char]0x2500
        [void]$lines.Add(@{ Text = "  $($hBar * 2) Workers $($hBar * 10)"; Color = "DarkCyan" })
        [void]$lines.Add(@{ Text = ("    {0,-8} {1,-12} {2,-14} {3,8} {4,12} {5,8}" -f "PID", "Status", "Current Day", "Pages", "Records", "%"); Color = "DarkGray" })
        foreach ($w in $Workers) {
            $color = switch ($w.State) {
                "Busy"    { "Green" }
                "Idle"    { "Yellow" }
                "Dead"    { "Red" }
                default   { "Gray" }
            }
            $recText = if ($w.Records) { "{0:N0}" -f [long]$w.Records } else { "-" }
            $pctText = if ($w.RecordPct) { "{0}%" -f $w.RecordPct } else { "-" }
            [void]$lines.Add(@{ Text = ("    {0,-8} {1,-12} {2,-14} {3,8} {4,12} {5,8}" -f $w.PID, $w.State, ($w.CurrentDay ?? "-"), ($w.Pages ?? "-"), $recText, $pctText); Color = $color })
        }
        [void]$lines.Add(@{ Text = ""; Color = "White" })
    }

    # Recent activity
    if ($RecentActivity -and $RecentActivity.Count -gt 0) {
        $activitySlice = @($RecentActivity | Select-Object -Last 4)
        [void]$lines.Add(@{ Text = ("  Recent Activity ({0} total):" -f $RecentActivity.Count); Color = "DarkCyan" })
        foreach ($act in $activitySlice) {
            [void]$lines.Add(@{ Text = ("    {0}  {1}" -f $act.Time, $act.Message); Color = "DarkGray" })
        }
    }

    # Recent errors
    if ($RecentErrors -and $RecentErrors.Count -gt 0) {
        [void]$lines.Add(@{ Text = ""; Color = "White" })
        $errorSlice = @($RecentErrors | Select-Object -Last 3)
        [void]$lines.Add(@{ Text = ("  Recent Errors ({0} total):" -f $RecentErrors.Count); Color = "Red" })
        foreach ($err in $errorSlice) {
            [void]$lines.Add(@{ Text = ("    {0}  {1}" -f $err.Time, $err.Message); Color = "DarkRed" })
        }
    }
    [void]$lines.Add(@{ Text = ""; Color = "White" })

    $script:AEDashboardLineCount = Write-DashboardFrame -Lines $lines -PreviousLineCount $(if ($script:AEDashboardLineCount) { $script:AEDashboardLineCount } else { 0 })
}

#endregion

