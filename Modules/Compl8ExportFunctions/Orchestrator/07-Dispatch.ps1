#region Dispatch Loop Engine

function Invoke-DispatchLoop {
    <#
    .SYNOPSIS
        Generic orchestrator dispatch loop for multi-terminal export coordination.
    .DESCRIPTION
        Replaces per-phase while($true) loops with a single engine that uses
        scriptblock callbacks for phase-specific behavior. Supports continuous
        pipeline mode where completing one task type generates tasks of another type.

        All shared state must be passed through the $Context hashtable because
        PowerShell scriptblocks invoked with & do not have access to the
        defining scope's variables.
    #>
    [CmdletBinding()]
    param(
        # ========== Core ==========
        [Parameter(Mandatory)]
        [string]$ExportDir,

        [Parameter(Mandatory)]
        [System.Collections.ArrayList]$Tasks,

        [Parameter(Mandatory)]
        [System.Collections.ArrayList]$WorkerProcesses,

        [hashtable]$Context = @{},

        # ========== Mandatory Callbacks ==========
        [Parameter(Mandatory)]
        [scriptblock]$OnScanCompletions,
        # Signature: { param($ExportDir, $WorkerDirs, $Context) }
        # Must return: @{ CompletedTasks = @(@{...}, ...); ErrorTasks = @(@{...}, ...) }

        [Parameter(Mandatory)]
        [scriptblock]$OnMatchTask,
        # Signature: { param($CompletionOrErrorData, $Tasks, $Context) }
        # Must return: the task object from $Tasks that matches, or $null

        [Parameter(Mandatory)]
        [scriptblock]$OnDispatchTask,
        # Signature: { param($Worker, $NextPendingTask, $Context) }
        # Must return: $true if task was sent successfully

        [Parameter(Mandatory)]
        [scriptblock]$OnShowDashboard,
        # Signature: { param($LoopState, $Context) }

        # ========== Optional Callbacks ==========
        [scriptblock]$OnCompletionGeneratesTasks = $null,
        # Signature: { param($CompletedTask, $CompletionData, $Context) }
        # Must return: @() array of new task hashtables to add to $Tasks

        [scriptblock]$OnCheckComplete = $null,
        # Signature: { param($Tasks, $LoopState, $Context) }
        # Must return: $true to break the loop
        # Default behavior if null: break when no Pending or InProgress tasks

        [scriptblock]$OnAllWorkersDead = $null,
        # Signature: { param($Tasks, $PendingCount, $Context) }
        # Called when all workers are dead but pending/in-progress tasks remain

        [scriptblock]$OnIterationComplete = $null,
        # Signature: { param($Tasks, $LoopState, $Context) }
        # Called at end of each iteration for CSV writes, phase updates, etc.

        # ========== Tuning ==========
        [int]$SleepSeconds = 2,
        [int]$MaxRecentActivity = 20,
        [int]$MaxRecentErrors = 10
    )

    # Initialize loop state
    $loopState = @{
        Iteration        = 0
        StartTime        = Get-Date
        LastActivityTime = Get-Date
        ElapsedTime      = [TimeSpan]::Zero
        RecentActivity   = [System.Collections.ArrayList]::new()
        RecentErrors     = [System.Collections.ArrayList]::new()
        CompletedCount   = 0
        TotalCount       = $Tasks.Count
        WorkerProcesses  = $WorkerProcesses
    }

    while ($true) {
        try {
            $loopState.Iteration++
            $loopState.ElapsedTime = (Get-Date) - $loopState.StartTime

            # Step 1: Scan completions
            $workerDirs = @($WorkerProcesses | ForEach-Object { $_.WorkerDir })
            $scanResult = & $OnScanCompletions $ExportDir $workerDirs $Context

            # Step 2: Process completed tasks
            foreach ($completion in @($scanResult.CompletedTasks)) {
                $matchedTask = & $OnMatchTask $completion $Tasks $Context
                if ($null -eq $matchedTask) { continue }

                $matchedTask.Status = "Completed"
                # Copy known metadata fields from completion to task
                foreach ($key in @('RecordCount', 'PageCount', 'ExpectedCount', 'TotalAvailable')) {
                    if ($completion.ContainsKey($key)) { $matchedTask.$key = $completion[$key] }
                }

                [void]$loopState.RecentActivity.Add(@{
                    Time    = (Get-Date -Format "HH:mm:ss")
                    Message = if ($completion.ContainsKey('Message')) { $completion.Message } else { "Task completed" }
                })
                $loopState.LastActivityTime = Get-Date

                # Continuous pipeline: completion generates new tasks
                if ($null -ne $OnCompletionGeneratesTasks) {
                    $newTasks = @(& $OnCompletionGeneratesTasks $matchedTask $completion $Context)
                    foreach ($nt in $newTasks) {
                        [void]$Tasks.Add($nt)
                    }
                }
            }

            # Step 3: Process error tasks
            foreach ($err in @($scanResult.ErrorTasks)) {
                $matchedTask = & $OnMatchTask $err $Tasks $Context
                if ($null -eq $matchedTask) { continue }

                $matchedTask.Status = "Error"
                if ($err.ContainsKey('ErrorMessage')) { $matchedTask.ErrorMessage = $err.ErrorMessage }

                [void]$loopState.RecentErrors.Add(@{
                    Time    = (Get-Date -Format "HH:mm:ss")
                    Message = if ($err.ContainsKey('Message')) { $err.Message } else { $err.ErrorMessage }
                })
            }

            # Step 4: Worker health check + task reclamation
            foreach ($task in $Tasks) {
                if ($task.Status -eq "InProgress" -and $task.AssignedPID) {
                    $assignedPID = $task.AssignedPID -as [int]
                    if ($assignedPID -le 0) { continue }
                    $wInfo = $WorkerProcesses | Where-Object { $_.PID -eq $assignedPID } | Select-Object -First 1
                    $wDir = if ($wInfo) { $wInfo.WorkerDir } else { $null }
                    if (-not (Test-WorkerAlive -WorkerPID $assignedPID -WorkerDir $wDir)) {
                        $task.Status = "Pending"
                        $task.AssignedPID = 0
                        [void]$loopState.RecentActivity.Add(@{
                            Time    = (Get-Date -Format "HH:mm:ss")
                            Message = "Reclaimed task from dead worker PID $assignedPID"
                        })
                    }
                }
            }

            # Step 5: Adaptive worker scaling (opt-in via $env:COMPL8_ADAPTIVE_WORKERS=1).
            # Park workers when there's slack, unpark them when the queue grows.
            # Parked workers stay alive (session warm, no re-auth) but skip dispatch.
            if ($env:COMPL8_ADAPTIVE_WORKERS -eq "1") {
                $aliveWorkers = @($WorkerProcesses | Where-Object { Test-WorkerAlive -WorkerPID $_.PID -WorkerDir $_.WorkerDir })
                $pendingCount = @($Tasks | Where-Object { $_.Status -eq "Pending" }).Count
                # Target ~1 active worker per 4 pending tasks, bounded by alive count.
                $target = [Math]::Max(1, [Math]::Min($aliveWorkers.Count, [Math]::Ceiling($pendingCount / 4)))
                # Pick the lowest-PID workers as active; park the rest.
                $sortedWorkers = @($aliveWorkers | Sort-Object PID)
                for ($i = 0; $i -lt $sortedWorkers.Count; $i++) {
                    Set-WorkerParked -WorkerDir $sortedWorkers[$i].WorkerDir -Parked:($i -ge $target)
                }
            }

            # Step 6: Dispatch pending tasks to idle, non-parked workers
            foreach ($w in $WorkerProcesses) {
                if (Test-WorkerParked -WorkerDir $w.WorkerDir) { continue }
                $state = Get-WorkerState -WorkerDir $w.WorkerDir -WorkerPID $w.PID
                if ($state -eq "Idle") {
                    $nextTask = $Tasks | Where-Object { $_.Status -eq "Pending" } | Select-Object -First 1
                    if ($null -ne $nextTask) {
                        $sent = & $OnDispatchTask $w $nextTask $Context
                        if ($sent) {
                            $nextTask.Status = "InProgress"
                            $nextTask.AssignedPID = $w.PID
                            $loopState.LastActivityTime = Get-Date
                        }
                    }
                }
            }

            # Step 6: Build loop state for dashboard
            $loopState.CompletedCount = @($Tasks | Where-Object { $_.Status -in @("Completed", "Error") }).Count
            $loopState.TotalCount = $Tasks.Count
            $loopState.WorkerProcesses = $WorkerProcesses

            # Trim recent lists
            while ($loopState.RecentActivity.Count -gt $MaxRecentActivity) { $loopState.RecentActivity.RemoveAt(0) }
            while ($loopState.RecentErrors.Count -gt $MaxRecentErrors) { $loopState.RecentErrors.RemoveAt(0) }

            # Show dashboard
            & $OnShowDashboard $loopState $Context

            # Step 7: Check completion
            $pendingOrInProgress = @($Tasks | Where-Object { $_.Status -in @("Pending", "InProgress") }).Count
            if ($null -ne $OnCheckComplete) {
                if (& $OnCheckComplete $Tasks $loopState $Context) { break }
            }
            else {
                if ($pendingOrInProgress -eq 0) { break }
            }

            # Step 8: All workers dead check
            $aliveCount = @($WorkerProcesses | Where-Object {
                Test-WorkerAlive -WorkerPID $_.PID -WorkerDir $_.WorkerDir
            }).Count
            $pendingCount = @($Tasks | Where-Object { $_.Status -eq "Pending" }).Count
            if ($aliveCount -eq 0 -and ($pendingCount -gt 0 -or @($Tasks | Where-Object { $_.Status -eq "InProgress" }).Count -gt 0)) {
                if ($null -ne $OnAllWorkersDead) {
                    & $OnAllWorkersDead $Tasks $pendingCount $Context
                }
                else {
                    Write-ExportLog -Message "All workers dead with $pendingCount pending tasks - saving state for resume" -Level Error
                }
                break
            }

            # Step 9: Iteration complete callback
            if ($null -ne $OnIterationComplete) {
                & $OnIterationComplete $Tasks $loopState $Context
            }

            # Step 10: Flush keyboard buffer + sleep
            try { while ([Console]::KeyAvailable) { $null = [Console]::ReadKey($true) } } catch { }
            Start-Sleep -Seconds $SleepSeconds
        }
        catch {
            Write-ExportLog -Message "Dispatch loop iteration error: $($_.Exception.Message)" -Level Error
            Start-Sleep -Seconds $SleepSeconds
        }
    }

    # Return loop state for caller to inspect
    return $loopState
}

#endregion

