#region Worker Health Monitoring

function Test-WorkerAlive {
    <#
    .SYNOPSIS
        Checks if a worker process is still running and its script loop is active.
    .DESCRIPTION
        First checks if the OS process exists. If WorkerDir is provided, uses a
        two-tier approach:
        1. If currenttask file exists, the worker is processing a task. A single
           API page can block 15+ minutes, so the worker is trusted until a
           generous lease (LeaseMinutes) elapses with NO sign of life. The most
           recent of (currenttask mtime, Progress.log mtime) is the heartbeat:
           Progress.log is written per page, and currenttask mtime marks task
           pickup. If neither has advanced within the lease, the worker is hung
           (e.g. a network partition with no socket timeout) and is declared dead
           so the orchestrator can reclaim the task.
        2. With no active task, a stale Progress.log means the script loop exited
           (e.g. crashed between iterations, or pwsh -NoExit keeping process alive).
    .PARAMETER LeaseMinutes
        Maximum minutes a worker may hold an active task with no progress before
        it is declared hung. Default 30 (double the ~15-min worst-case page).
    #>
    param(
        [Parameter(Mandatory)][int]$WorkerPID,
        [string]$WorkerDir,
        [int]$LeaseMinutes = 30
    )
    try {
        $proc = Get-Process -Id $WorkerPID -ErrorAction SilentlyContinue
        if ($null -eq $proc -or $proc.HasExited) { return $false }
    }
    catch { return $false }

    # Process is alive — if WorkerDir provided, check for active task or staleness
    if ($WorkerDir) {
        # Use try/catch instead of Test-Path to avoid TOCTOU race.
        $currentTaskPath = Join-Path $WorkerDir "currenttask"
        $progPath = Join-Path $WorkerDir "Progress.log"
        $hasCurrentTask = $false
        $lastSignOfLife = [datetime]::MinValue

        try {
            $null = [System.IO.File]::GetAttributes($currentTaskPath)   # throws if file absent
            $hasCurrentTask = $true
            $ctTime = [System.IO.File]::GetLastWriteTime($currentTaskPath)
            if ($ctTime.Year -gt 1601 -and $ctTime -gt $lastSignOfLife) { $lastSignOfLife = $ctTime }
        }
        catch [System.IO.FileNotFoundException], [System.IO.DirectoryNotFoundException] {
            # No active task — fall through to staleness checks
        }

        # Progress.log mtime acts as the per-page heartbeat for both branches.
        try {
            $progTime = [System.IO.File]::GetLastWriteTime($progPath)
            if ($progTime.Year -gt 1601 -and $progTime -gt $lastSignOfLife) { $lastSignOfLife = $progTime }
        }
        catch { <# Progress.log inaccessible — skip #> }

        if ($hasCurrentTask) {
            # Active task: trust the worker until the lease expires with no progress.
            if ($lastSignOfLife -ne [datetime]::MinValue) {
                $idle = (Get-Date) - $lastSignOfLife
                if ($idle.TotalMinutes -gt $LeaseMinutes) { return $false }
            }
            return $true
        }

        # No active task — a stale Progress.log means the loop has exited.
        if ($lastSignOfLife -ne [datetime]::MinValue) {
            $staleness = (Get-Date) - $lastSignOfLife
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

