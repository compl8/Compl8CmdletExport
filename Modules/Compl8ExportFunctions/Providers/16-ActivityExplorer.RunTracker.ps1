#region Activity Explorer Run Tracker Functions

function Get-ActivityExplorerRunTracker {
    <#
    .SYNOPSIS
        Loads or creates an Activity Explorer run tracker.
    .DESCRIPTION
        If a tracker file exists at the specified path, loads and returns it.
        Otherwise creates a new tracker with default values.
    .PARAMETER TrackerPath
        Path to the RunTracker.json file.
    .OUTPUTS
        Hashtable containing the run tracker state.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$TrackerPath
    )

    if (Test-Path $TrackerPath) {
        try {
            $content = Get-Content -Raw -Path $TrackerPath -ErrorAction Stop
            $loaded = $content | ConvertFrom-Json -ErrorAction Stop
            if ($null -eq $loaded) {
                Write-ExportLog -Message "  Run tracker file parsed as null, creating new tracker" -Level Warning
            }
            else {
                # Convert PSCustomObject to hashtable for easier manipulation
                $tracker = @{
                    CompletedPages  = if ($null -ne $loaded.CompletedPages) { [int]$loaded.CompletedPages } else { 0 }
                    TotalRecords    = if ($null -ne $loaded.TotalRecords) { [int]$loaded.TotalRecords } else { 0 }
                    LastWaterMark   = $loaded.LastWaterMark
                    LastPageTime    = $loaded.LastPageTime
                    Status          = if ($loaded.Status) { $loaded.Status } else { "InProgress" }
                    PartialErrors   = if ($loaded.PartialErrors) { @($loaded.PartialErrors) } else { @() }
                    StartTime       = if ($loaded.StartTime) { $loaded.StartTime } else { (Get-Date).ToString("o") }
                    PartialFailure  = if ($null -ne $loaded.PartialFailure) { [bool]$loaded.PartialFailure } else { $false }
                }

                Write-ExportLog -Message ("  Loaded existing run tracker: {0} pages, {1} records" -f $tracker.CompletedPages, $tracker.TotalRecords) -Level Info
                return $tracker
            }
        }
        catch {
            Write-ExportLog -Message ("  Failed to load run tracker, creating new: {0}" -f $_.Exception.Message) -Level Warning
        }
    }

    # Create new tracker
    $tracker = @{
        CompletedPages  = 0
        TotalRecords    = 0
        LastWaterMark   = $null
        LastPageTime    = $null
        Status          = "InProgress"
        PartialErrors   = @()
        StartTime       = (Get-Date).ToString("o")
        PartialFailure  = $false
    }

    return $tracker
}

function Save-ActivityExplorerRunTracker {
    <#
    .SYNOPSIS
        Saves the Activity Explorer run tracker state atomically.
    .DESCRIPTION
        Writes the tracker to a temporary file first, then renames it to the
        final path. This prevents corruption if the process is interrupted mid-write.
    .PARAMETER Tracker
        The run tracker hashtable to save.
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

    # Update the save timestamp
    $Tracker['LastSaveTime'] = (Get-Date).ToString("o")

    # Atomic write: temp file then rename
    $tempPath = $TrackerPath + ".tmp"

    try {
        $serializableTracker = ConvertTo-SerializableObject -InputObject $Tracker
        $json = $serializableTracker | ConvertTo-Json -Depth 10
        Set-Content -Path $tempPath -Value $json -Encoding UTF8 -ErrorAction Stop

        # Rename (atomic on NTFS)
        if (Test-Path $TrackerPath) {
            Remove-Item -Path $TrackerPath -Force -ErrorAction Stop
        }
        Rename-Item -Path $tempPath -NewName (Split-Path $TrackerPath -Leaf) -Force -ErrorAction Stop
    }
    catch {
        # Clean up temp file on failure
        if (Test-Path $tempPath) {
            Remove-Item -Path $tempPath -Force -ErrorAction SilentlyContinue
        }
        Write-ExportLog -Message ("  Failed to save run tracker: {0}" -f $_.Exception.Message) -Level Warning
    }
}

#endregion

