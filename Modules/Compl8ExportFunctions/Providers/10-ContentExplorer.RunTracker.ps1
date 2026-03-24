#region Content Explorer - Run Tracker (State Persistence)

function Get-ContentExplorerRunTracker {
    <#
    .SYNOPSIS
        Loads or creates a Content Explorer run tracker for resumable exports.
    .DESCRIPTION
        If the tracker file exists, loads it. Otherwise creates a new tracker
        with default values. The tracker persists completed tasks, output files,
        SIT mapping, and export statistics.
    .PARAMETER TrackerPath
        Path to the RunTracker.json file.
    .OUTPUTS
        Hashtable with tracker state.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$TrackerPath
    )

    if (Test-Path $TrackerPath) {
        try {
            $content = Get-Content -Raw -Path $TrackerPath -ErrorAction Stop
            $tracker = $content | ConvertFrom-Json -AsHashtable -ErrorAction Stop
            if ($tracker) {
                Write-ExportLog -Message "Loaded run tracker: $TrackerPath" -Level Info
                # Ensure required keys exist
                if (-not $tracker.ContainsKey('CompletedTasks')) { $tracker.CompletedTasks = @() }
                if (-not $tracker.ContainsKey('OutputFiles')) { $tracker.OutputFiles = @() }
                if (-not $tracker.ContainsKey('SitMapping')) { $tracker.SitMapping = @{} }
                if (-not $tracker.ContainsKey('TotalExported')) { $tracker.TotalExported = 0 }
                if (-not $tracker.ContainsKey('TotalDeduplicated')) { $tracker.TotalDeduplicated = 0 }
                if (-not $tracker.ContainsKey('Status')) { $tracker.Status = "InProgress" }
                if (-not $tracker.ContainsKey('TaskMetrics')) { $tracker.TaskMetrics = @() }
                return $tracker
            }
        }
        catch {
            Write-ExportLog -Message ("Failed to load run tracker, creating new: " + $_.Exception.Message) -Level Warning
        }
    }

    # Create new tracker
    $tracker = @{
        CompletedTasks    = @()
        OutputFiles       = @()
        SitMapping        = @{}
        TotalExported     = 0
        TotalDeduplicated = 0
        Status            = "InProgress"
        TaskMetrics       = @()
        CreatedTime       = (Get-Date).ToString("o")
        LastUpdated       = (Get-Date).ToString("o")
    }

    return $tracker
}

function Save-ContentExplorerRunTracker {
    <#
    .SYNOPSIS
        Saves the Content Explorer run tracker using atomic write.
    .DESCRIPTION
        Writes the tracker to a temporary file first, then renames to the target
        path. This prevents corruption if the process is interrupted mid-write.
    .PARAMETER Tracker
        The tracker hashtable to save.
    .PARAMETER TrackerPath
        Path to the RunTracker.json file.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Tracker,

        [Parameter(Mandatory)]
        [string]$TrackerPath
    )

    $Tracker.LastUpdated = (Get-Date).ToString("o")

    try {
        $serializableTracker = ConvertTo-SerializableObject -InputObject $Tracker
        $json = $serializableTracker | ConvertTo-Json -Depth 20

        # Atomic write: write to temp file then rename
        $tempPath = $TrackerPath + ".tmp." + [System.IO.Path]::GetRandomFileName()
        Set-Content -Path $tempPath -Value $json -Encoding UTF8 -ErrorAction Stop

        # Rename (atomic on NTFS)
        if (Test-Path $TrackerPath) {
            [System.IO.File]::Delete($TrackerPath)
        }
        [System.IO.File]::Move($tempPath, $TrackerPath)
    }
    catch {
        Write-ExportLog -Message ("Failed to save run tracker: " + $_.Exception.Message) -Level Warning
        # Clean up temp file on failure
        if ($tempPath -and (Test-Path $tempPath -ErrorAction SilentlyContinue)) {
            Remove-Item -Path $tempPath -Force -ErrorAction SilentlyContinue
        }
    }
}

#endregion

