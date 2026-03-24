#region Content Explorer - Adaptive Paging & Telemetry

function New-ContentExplorerTelemetry {
    <#
    .SYNOPSIS
        Creates a new telemetry tracking object for a Content Explorer export task.
    .PARAMETER TagType
        The classifier tag type.
    .PARAMETER TagName
        The classifier tag name.
    .PARAMETER Workload
        The workload being exported.
    .OUTPUTS
        Hashtable with telemetry tracking fields.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$TagType,

        [Parameter(Mandatory)]
        [string]$TagName,

        [Parameter(Mandatory)]
        [string]$Workload
    )

    return @{
        TagType       = $TagType
        TagName       = $TagName
        Workload      = $Workload
        Location      = ""
        LocationType  = ""
        PageSize      = 0
        RecordCount   = 0
        PageCount     = 0
        TotalTimeMs   = 0
        Status        = "Pending"
        StartedTime   = (Get-Date).ToString("o")
        CompletedTime = $null
        Hostname      = $env:COMPUTERNAME
        PID           = $PID
    }
}

function Get-AdaptivePageSize {
    <#
    .SYNOPSIS
        Selects optimal page size based on volume and location distribution.
    .DESCRIPTION
        Uses volume-based selection for high-volume tasks (25,000+ records) and
        distribution-based selection for lower volumes. Small locations (<100 items)
        are detected to determine if the workload is dominated by tiny locations.

        Volume-based (25k+):
          - 500 if 90%+ locations are small
          - 2000 if median location >500 items
          - 1000 default (best throughput baseline)

        Distribution-based (<25k):
          - Exchange: 500
          - SharePoint: 500
          - OneDrive: 1000
          - Teams: 500

        Bounds (applied after selection):
          - Floor: 100 (minimum page size)
          - Ceiling: 2x total expected count (no larger than needed)
    .PARAMETER Task
        Task hashtable with Locations array (each with Name, ExpectedCount).
    .PARAMETER Workload
        The workload type.
    .PARAMETER TelemetryDatabasePath
        Path to telemetry database (for future use with historical analysis).
    .OUTPUTS
        Integer page size (clamped to [100, max(100, 2 * totalExpected)]).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Task,

        [Parameter(Mandatory)]
        [string]$Workload,

        [string]$TelemetryDatabasePath
    )

    $smallThreshold = 100
    $highVolumeThreshold = 25000
    $locations = @($Task.Locations)

    # Calculate total expected
    $totalExpected = 0
    foreach ($loc in $locations) {
        $count = $loc.ExpectedCount -as [int]
        if ($count) { $totalExpected += $count }
    }

    # No data - return default (clamped to floor)
    if ($locations.Count -eq 0 -or $totalExpected -eq 0) {
        return 1000
    }

    # Classify small vs large locations
    $smallLocations = @($locations | Where-Object { ($_.ExpectedCount -as [int]) -lt $smallThreshold })
    $smallRatio = $smallLocations.Count / $locations.Count

    # Calculate median location size
    $sortedCounts = @($locations | ForEach-Object { $_.ExpectedCount -as [int] } | Where-Object { $_ -gt 0 } | Sort-Object)
    $medianCount = 0
    if ($sortedCounts.Count -gt 0) {
        $midIndex = [Math]::Floor($sortedCounts.Count / 2)
        if ($sortedCounts.Count % 2 -eq 0 -and $sortedCounts.Count -gt 1) {
            $medianCount = [Math]::Round(($sortedCounts[$midIndex - 1] + $sortedCounts[$midIndex]) / 2)
        }
        else {
            $medianCount = $sortedCounts[$midIndex]
        }
    }

    # Select base page size from algorithm
    $selectedSize = 1000  # default
    if ($totalExpected -ge $highVolumeThreshold) {
        # Volume-based selection for high-volume tasks
        if ($smallRatio -ge 0.9) {
            $selectedSize = 500
        }
        elseif ($medianCount -gt 500) {
            $selectedSize = 2000
        }
        else {
            $selectedSize = 1000
        }
    }
    else {
        # Distribution-based selection for lower volumes
        $selectedSize = switch ($Workload) {
            "Exchange"   { 500 }
            "SharePoint" { 500 }
            "OneDrive"   { 1000 }
            "Teams"      { 500 }
            default      { 1000 }
        }
    }

    # Apply bounds: floor 100, ceiling 2x total expected
    $maxPageSize = [Math]::Max(100, 2 * $totalExpected)
    $clamped = [Math]::Max(100, [Math]::Min($selectedSize, $maxPageSize))
    return $clamped
}

function Save-ContentExplorerTelemetry {
    <#
    .SYNOPSIS
        Writes a telemetry entry as a JSONL line to the database file.
    .DESCRIPTION
        Appends a single JSON line to the telemetry JSONL database.
        Creates the directory and file if they do not exist.
    .PARAMETER Telemetry
        The telemetry hashtable to save.
    .PARAMETER DatabasePath
        Path to the JSONL telemetry database file.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Telemetry,

        [Parameter(Mandatory)]
        [string]$DatabasePath
    )

    try {
        $dbDir = Split-Path $DatabasePath -Parent
        if ($dbDir -and -not (Test-Path $dbDir)) {
            New-Item -ItemType Directory -Force -Path $dbDir | Out-Null
        }

        $serializable = ConvertTo-SerializableObject -InputObject $Telemetry
        $jsonLine = $serializable | ConvertTo-Json -Depth 10 -Compress
        [System.IO.File]::AppendAllText($DatabasePath, $jsonLine + [Environment]::NewLine, [System.Text.Encoding]::UTF8)
    }
    catch {
        Write-ExportLog -Message ("Failed to save telemetry: " + $_.Exception.Message) -Level Warning
    }
}

function Get-ContentExplorerTelemetryStats {
    <#
    .SYNOPSIS
        Reads the telemetry database and returns summary statistics.
    .DESCRIPTION
        Parses the JSONL telemetry file and aggregates statistics by workload
        including total records, pages, average time, and page size distribution.
    .PARAMETER DatabasePath
        Path to the JSONL telemetry database file.
    .OUTPUTS
        Hashtable with:
          TotalEntries   - Number of telemetry entries
          ByWorkload     - Hashtable with per-workload stats
          ByPageSize     - Hashtable with per-page-size stats
          AvgTimePerPage - Average time per page in milliseconds
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$DatabasePath
    )

    $stats = @{
        TotalEntries   = 0
        ByWorkload     = @{}
        ByPageSize     = @{}
        AvgTimePerPage = 0
    }

    if (-not (Test-Path $DatabasePath)) {
        return $stats
    }

    try {
        $lines = Get-Content -Path $DatabasePath -Encoding UTF8 -ErrorAction Stop
    }
    catch {
        Write-ExportLog -Message ("Failed to read telemetry database: " + $_.Exception.Message) -Level Warning
        return $stats
    }

    $totalTime = 0
    $totalPages = 0

    foreach ($line in $lines) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }

        try {
            $entry = $line | ConvertFrom-Json -ErrorAction Stop
        }
        catch {
            Write-Verbose "Skipping malformed telemetry line: $($_.Exception.Message)"
            continue
        }
        if ($null -eq $entry) { continue }

        $stats.TotalEntries++

        # By workload
        $wl = $entry.Workload
        if ($wl) {
            if (-not $stats.ByWorkload.ContainsKey($wl)) {
                $stats.ByWorkload[$wl] = @{
                    Entries      = 0
                    TotalRecords = 0
                    TotalPages   = 0
                    TotalTimeMs  = 0
                    Completed    = 0
                    Failed       = 0
                }
            }
            $wlStats = $stats.ByWorkload[$wl]
            $wlStats.Entries++
            $wlStats.TotalRecords += ($entry.RecordCount -as [int])
            $wlStats.TotalPages += ($entry.PageCount -as [int])
            $wlStats.TotalTimeMs += ($entry.TotalTimeMs -as [int])
            if ($entry.Status -eq "Completed") { $wlStats.Completed++ }
            elseif ($entry.Status -in @("Failed", "PartialFailure")) { $wlStats.Failed++ }
        }

        # By page size
        $ps = $entry.PageSize -as [string]
        if ($ps) {
            if (-not $stats.ByPageSize.ContainsKey($ps)) {
                $stats.ByPageSize[$ps] = @{
                    Entries      = 0
                    TotalRecords = 0
                    TotalPages   = 0
                    TotalTimeMs  = 0
                }
            }
            $psStats = $stats.ByPageSize[$ps]
            $psStats.Entries++
            $psStats.TotalRecords += ($entry.RecordCount -as [int])
            $psStats.TotalPages += ($entry.PageCount -as [int])
            $psStats.TotalTimeMs += ($entry.TotalTimeMs -as [int])
        }

        $totalTime += ($entry.TotalTimeMs -as [int])
        $totalPages += ($entry.PageCount -as [int])
    }

    if ($totalPages -gt 0) {
        $stats.AvgTimePerPage = [Math]::Round($totalTime / $totalPages, 0)
    }

    return $stats
}

#endregion

