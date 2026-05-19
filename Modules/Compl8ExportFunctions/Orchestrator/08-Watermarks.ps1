#region Content Explorer - Watermarks (foundation for incremental exports)

function Get-WatermarkPath {
    <#
    .SYNOPSIS
        Returns the path to the per-tenant watermark file under ConfigFiles\.
    #>
    param(
        [Parameter(Mandatory)][string]$ScriptRoot,
        [string]$TenantPrefix
    )
    $suffix = if ([string]::IsNullOrWhiteSpace($TenantPrefix)) { "default" } else { $TenantPrefix }
    return Join-Path $ScriptRoot "ConfigFiles" ("Watermarks-{0}.local.json" -f $suffix)
}

function Read-Watermarks {
    <#
    .SYNOPSIS
        Loads the watermark file for a tenant. Returns an empty structure if missing.
    .OUTPUTS
        Hashtable: TenantPrefix, LastRunAt, LastFullRunAt, Tasks (hashtable keyed
        by "TagType|TagName|Workload").
    #>
    param(
        [Parameter(Mandatory)][string]$ScriptRoot,
        [string]$TenantPrefix
    )
    $path = Get-WatermarkPath -ScriptRoot $ScriptRoot -TenantPrefix $TenantPrefix
    if (-not (Test-Path $path)) {
        return @{ TenantPrefix = $TenantPrefix; LastRunAt = $null; LastFullRunAt = $null; Tasks = @{} }
    }
    try {
        $raw = Get-Content -Path $path -Raw -Encoding UTF8 | ConvertFrom-Json
    }
    catch {
        Write-Warning ("Watermarks file is malformed, ignoring: {0}" -f $_.Exception.Message)
        return @{ TenantPrefix = $TenantPrefix; LastRunAt = $null; LastFullRunAt = $null; Tasks = @{} }
    }
    $tasks = @{}
    if ($raw.Tasks) {
        foreach ($prop in $raw.Tasks.PSObject.Properties) {
            $tasks[$prop.Name] = @{
                AggregateCount = ($prop.Value.AggregateCount -as [long])
                ExportedCount  = ($prop.Value.ExportedCount -as [long])
                LastRunAt      = $prop.Value.LastRunAt
            }
        }
    }
    return @{
        TenantPrefix  = $raw.TenantPrefix
        LastRunAt     = $raw.LastRunAt
        LastFullRunAt = $raw.LastFullRunAt
        Tasks         = $tasks
    }
}

function Write-Watermarks {
    <#
    .SYNOPSIS
        Persists watermarks for a tenant. Atomic write via temp+rename.
    #>
    param(
        [Parameter(Mandatory)][string]$ScriptRoot,
        [string]$TenantPrefix,
        [Parameter(Mandatory)][hashtable]$Tasks,
        [switch]$WasFullRun
    )
    $path = Get-WatermarkPath -ScriptRoot $ScriptRoot -TenantPrefix $TenantPrefix
    $existing = Read-Watermarks -ScriptRoot $ScriptRoot -TenantPrefix $TenantPrefix
    $nowIso = (Get-Date).ToString("o")
    $payload = @{
        TenantPrefix  = $TenantPrefix
        LastRunAt     = $nowIso
        LastFullRunAt = if ($WasFullRun) { $nowIso } else { $existing.LastFullRunAt }
        Tasks         = $Tasks
    }
    $tmpPath = $path + ".tmp.$PID"
    try {
        $payload | ConvertTo-Json -Depth 6 | Set-Content -Path $tmpPath -Encoding UTF8
        # 3-arg File.Move (overwrite=true) is atomic on NTFS under .NET 5+
        # and closes the delete-before-move crash window.
        [System.IO.File]::Move($tmpPath, $path, $true)
    }
    catch {
        Write-Warning ("Could not save watermarks: {0}" -f $_.Exception.Message)
        if (Test-Path $tmpPath) { try { [System.IO.File]::Delete($tmpPath) } catch { } }
    }
}

function Save-WatermarksFromDetailTasks {
    <#
    .SYNOPSIS
        Reads a completed DetailTasks.csv and persists per-task watermarks.
    .DESCRIPTION
        Called at end of a successful export. Skips tasks in Error state. Uses
        OriginalExpectedCount as the aggregate count and ExpectedCount (overwritten
        to actual exported count) as the exported count.
    #>
    param(
        [Parameter(Mandatory)][string]$ScriptRoot,
        [string]$TenantPrefix,
        [Parameter(Mandatory)][array]$DetailTasks,
        [switch]$WasFullRun
    )
    $tasks = @{}
    foreach ($t in $DetailTasks) {
        if ($t.Status -eq "Error") { continue }
        $location = if ($t.Location) { $t.Location } else { "" }
        $key = "{0}|{1}|{2}|{3}" -f $t.TagType, $t.TagName, $t.Workload, $location
        $aggregate = if ($t.OriginalExpectedCount) { ($t.OriginalExpectedCount -as [long]) } else { ($t.ExpectedCount -as [long]) }
        $exported  = ($t.ExpectedCount -as [long])
        $tasks[$key] = @{
            AggregateCount = $aggregate
            ExportedCount  = $exported
            LastRunAt      = (Get-Date).ToString("o")
        }
    }
    Write-Watermarks -ScriptRoot $ScriptRoot -TenantPrefix $TenantPrefix -Tasks $tasks -WasFullRun:$WasFullRun
}

function Write-AggregateDeltaReport {
    <#
    .SYNOPSIS
        Compares current aggregate task counts against saved watermarks and writes
        _Coordination/AggregateDelta.json. Pure observability — no behavior change.
    .DESCRIPTION
        For each aggregate task in the current run, classify against the watermark:
          - "new"        — no prior watermark
          - "unchanged"  — count within 1% of prior
          - "grown"      — current > prior + 1%
          - "shrunk"     — current < prior - 1%
        Written to <ExportDir>/_Coordination/AggregateDelta.json.
    #>
    param(
        [Parameter(Mandatory)][string]$ExportDir,
        [Parameter(Mandatory)][hashtable]$Watermarks,
        [Parameter(Mandatory)][array]$AggregateTasks
    )
    $coordDir = Get-CoordinationDir $ExportDir
    if (-not (Test-Path $coordDir)) { return }

    $report = @{
        scan_time      = (Get-Date).ToString("o")
        tenant_prefix  = $Watermarks.TenantPrefix
        last_run_at    = $Watermarks.LastRunAt
        last_full_at   = $Watermarks.LastFullRunAt
        summary        = @{ new = 0; unchanged = 0; grown = 0; shrunk = 0 }
        tasks          = @()
    }
    $tolerance = 0.01
    foreach ($task in $AggregateTasks) {
        $location = if ($task.Location) { $task.Location } else { "" }
        $key = "{0}|{1}|{2}|{3}" -f $task.TagType, $task.TagName, $task.Workload, $location
        $currentCount = ($task.ExpectedCount -as [long])
        $watermark = $Watermarks.Tasks[$key]
        $classification = "new"
        $priorCount = $null
        if ($watermark) {
            $priorCount = $watermark.AggregateCount
            if ($priorCount -eq 0 -and $currentCount -eq 0) {
                $classification = "unchanged"
            }
            elseif ($priorCount -eq 0) {
                $classification = "grown"
            }
            else {
                $delta = [Math]::Abs($currentCount - $priorCount) / [double]$priorCount
                if ($delta -le $tolerance) { $classification = "unchanged" }
                elseif ($currentCount -gt $priorCount) { $classification = "grown" }
                else { $classification = "shrunk" }
            }
        }
        $report.summary[$classification]++
        $report.tasks += @{
            tag_type        = $task.TagType
            tag_name        = $task.TagName
            workload        = $task.Workload
            location        = $location
            current_count   = $currentCount
            prior_count     = $priorCount
            classification  = $classification
        }
    }

    $reportPath = Join-Path $coordDir "AggregateDelta.json"
    try {
        $report | ConvertTo-Json -Depth 6 | Set-Content -Path $reportPath -Encoding UTF8
        Write-ExportLog -Message ("Aggregate delta written: new={0} unchanged={1} grown={2} shrunk={3} -> {4}" -f $report.summary.new, $report.summary.unchanged, $report.summary.grown, $report.summary.shrunk, $reportPath) -Level Info
    }
    catch {
        Write-Warning ("Could not write aggregate delta report: {0}" -f $_.Exception.Message)
    }
}

#endregion
