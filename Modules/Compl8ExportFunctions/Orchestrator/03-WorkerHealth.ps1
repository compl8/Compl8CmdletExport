#region Worker Health Monitoring

function Test-WorkerAlive {
    <#
    .SYNOPSIS
        Checks if a worker process is still running and its script loop is active.
    .DESCRIPTION
        First checks if the OS process exists. If WorkerDir is provided, uses a
        two-tier approach:
        1. If currenttask file exists, the worker is actively processing — always alive
           (API calls can block for 15+ minutes per page, during which nothing updates)
        2. Otherwise, checks staleness of Progress.log and output files.
           A stale worker with no active task means the script loop has exited
           (e.g. crashed between iterations, or pwsh -NoExit keeping process alive).
    #>
    param(
        [Parameter(Mandatory)][int]$WorkerPID,
        [string]$WorkerDir
    )
    try {
        $proc = Get-Process -Id $WorkerPID -ErrorAction SilentlyContinue
        if ($null -eq $proc -or $proc.HasExited) { return $false }
    }
    catch { return $false }

    # Process is alive — if WorkerDir provided, check for active task or staleness
    if ($WorkerDir) {
        # If currenttask exists, worker is actively processing a task.
        # API calls (Export-ContentExplorerData, Export-ActivityExplorerData) can block
        # for 15+ minutes per page. No progress updates happen during this time.
        # Trust the process — it's working.
        # Use try/catch instead of Test-Path to avoid TOCTOU race.
        $currentTaskPath = Join-Path $WorkerDir "currenttask"
        try {
            $null = [System.IO.File]::GetAttributes($currentTaskPath)
            return $true  # currenttask exists — worker is busy
        }
        catch [System.IO.FileNotFoundException], [System.IO.DirectoryNotFoundException] {
            # No active task — fall through to staleness checks
        }

        # No active task — check staleness of Progress.log
        $progPath = Join-Path $WorkerDir "Progress.log"
        $lastWrite = [datetime]::MinValue
        try {
            $progTime = [System.IO.File]::GetLastWriteTime($progPath)
            if ($progTime.Year -gt 1601 -and $progTime -gt $lastWrite) { $lastWrite = $progTime }
        }
        catch { <# Progress.log inaccessible — skip #> }
        if ($lastWrite -ne [datetime]::MinValue) {
            $staleness = (Get-Date) - $lastWrite
            if ($staleness.TotalMinutes -gt 15) { return $false }
        }
        # No files yet = worker just started, treat as alive
    }
    return $true
}

function Get-WorkerState {
    <#
    .SYNOPSIS
        Returns the current state of a worker: Idle, Busy, WaitingForTask, or Dead.
    #>
    param(
        [Parameter(Mandatory)][string]$WorkerDir,
        [Parameter(Mandatory)][int]$WorkerPID
    )
    if (-not (Test-WorkerAlive -WorkerPID $WorkerPID -WorkerDir $WorkerDir)) {
        return "Dead"
    }
    # Use try/catch instead of Test-Path to avoid TOCTOU race on file existence checks.
    $hasCurrent = $false
    $hasNext = $false
    try {
        $null = [System.IO.File]::GetAttributes((Join-Path $WorkerDir "currenttask"))
        $hasCurrent = $true
    }
    catch [System.IO.FileNotFoundException], [System.IO.DirectoryNotFoundException] { }
    if ($hasCurrent) { return "Busy" }
    try {
        $null = [System.IO.File]::GetAttributes((Join-Path $WorkerDir "nexttask"))
        $hasNext = $true
    }
    catch [System.IO.FileNotFoundException], [System.IO.DirectoryNotFoundException] { }
    if ($hasNext) { return "WaitingForTask" }
    return "Idle"
}

#endregion

