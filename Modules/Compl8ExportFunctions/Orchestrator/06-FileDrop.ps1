#region File-Drop Coordination

function Get-ExportRunSigningKey {
    <#
    .SYNOPSIS
        Gets (or creates) the per-run signing key used for file-drop message integrity.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ExportDir,
        [switch]$CreateIfMissing
    )

    $coordDir = Get-CoordinationDir $ExportDir
    if ($CreateIfMissing -and -not (Test-Path $coordDir)) { New-Item -ItemType Directory -Force -Path $coordDir | Out-Null }
    $keyPath = Join-Path $coordDir "RunSigningKey.txt"

    if (Test-Path $keyPath) {
        return ([System.IO.File]::ReadAllText($keyPath)).Trim()
    }

    if (-not $CreateIfMissing) {
        return $null
    }

    $keyBytes = [byte[]]::new(32)
    [System.Security.Cryptography.RandomNumberGenerator]::Fill($keyBytes)
    $newKey = [Convert]::ToBase64String($keyBytes)

    try {
        $fs = [System.IO.File]::Open($keyPath, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
        try {
            $bytes = [System.Text.Encoding]::UTF8.GetBytes($newKey)
            $fs.Write($bytes, 0, $bytes.Length)
        }
        finally {
            $fs.Dispose()
        }
    }
    catch [System.IO.IOException] {
        # Another process created it first.
    }

    if (Test-Path $keyPath) {
        return ([System.IO.File]::ReadAllText($keyPath)).Trim()
    }

    throw "Failed to create or load signing key at $keyPath"
}

function Get-MessageSignature {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$PayloadJson,
        [Parameter(Mandatory)][string]$SigningKey
    )

    $keyBytes = [System.Text.Encoding]::UTF8.GetBytes($SigningKey)
    $payloadBytes = [System.Text.Encoding]::UTF8.GetBytes($PayloadJson)
    $hmac = [System.Security.Cryptography.HMACSHA256]::new($keyBytes)
    try {
        $hash = $hmac.ComputeHash($payloadBytes)
        return ([Convert]::ToHexString($hash)).ToLowerInvariant()
    }
    finally {
        $hmac.Dispose()
    }
}

function ConvertTo-SignedEnvelopeJson {
    <#
    .SYNOPSIS
        Wraps a payload in a signed envelope JSON document.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Payload,
        [string]$SigningKey
    )

    if ([string]::IsNullOrWhiteSpace($SigningKey)) {
        return ($Payload | ConvertTo-Json -Depth 20 -Compress)
    }

    $payloadJson = $Payload | ConvertTo-Json -Depth 20 -Compress
    $signature = Get-MessageSignature -PayloadJson $payloadJson -SigningKey $SigningKey
    return (@{
        Payload   = $Payload
        Signature = $signature
    } | ConvertTo-Json -Depth 22 -Compress)
}

function ConvertFrom-SignedEnvelopeJson {
    <#
    .SYNOPSIS
        Parses and validates a signed envelope JSON document.
    .OUTPUTS
        Hashtable containing the verified payload.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Json,
        [string]$SigningKey,
        [switch]$RequireSignature,
        [string]$Context = "message"
    )

    $parsed = $Json | ConvertFrom-Json -AsHashtable -ErrorAction Stop
    $hasEnvelope = ($parsed -is [hashtable]) -and $parsed.ContainsKey('Payload') -and $parsed.ContainsKey('Signature')

    if ($hasEnvelope) {
        if ([string]::IsNullOrWhiteSpace($SigningKey)) {
            throw "Signed $Context cannot be validated because the signing key is missing."
        }

        $payloadJson = $parsed.Payload | ConvertTo-Json -Depth 20 -Compress
        $expectedSignature = Get-MessageSignature -PayloadJson $payloadJson -SigningKey $SigningKey
        if ($expectedSignature -ne $parsed.Signature) {
            throw "Signature verification failed for $Context."
        }

        if ($parsed.Payload -is [hashtable]) {
            return $parsed.Payload
        }
        return (($parsed.Payload | ConvertTo-Json -Depth 20 -Compress) | ConvertFrom-Json -AsHashtable -ErrorAction Stop)
    }

    if ($RequireSignature -and -not [string]::IsNullOrWhiteSpace($SigningKey)) {
        throw "Unsigned $Context rejected."
    }

    if ($parsed -is [hashtable]) {
        return $parsed
    }
    return (($parsed | ConvertTo-Json -Depth 20 -Compress) | ConvertFrom-Json -AsHashtable -ErrorAction Stop)
}

function Test-WorkerTaskSchema {
    <#
    .SYNOPSIS
        Validates task payload shape and allowed values before worker execution.
    .OUTPUTS
        Boolean. Throws on invalid payload.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$TaskData
    )

    if (-not $TaskData.ContainsKey('Phase')) {
        throw "Task missing required field 'Phase'."
    }

    $phase = [string]$TaskData.Phase
    $allowedPhases = @('Planning', 'Aggregate', 'Detail', 'AEExport')
    if ($phase -notin $allowedPhases) {
        throw "Invalid task phase '$phase'."
    }

    if ($phase -in @('Planning', 'Aggregate', 'Detail')) {
        if ([string]::IsNullOrWhiteSpace([string]$TaskData.TagType)) { throw "Task missing TagType." }
        if ([string]::IsNullOrWhiteSpace([string]$TaskData.TagName)) { throw "Task missing TagName." }
        if ([string]::IsNullOrWhiteSpace([string]$TaskData.Workload)) { throw "Task missing Workload." }

        $allowedTagTypes = @('Retention', 'SensitiveInformationType', 'Sensitivity', 'TrainableClassifier')
        if ([string]$TaskData.TagType -notin $allowedTagTypes) {
            throw "Invalid TagType '$($TaskData.TagType)'."
        }

        $allowedWorkloads = @('Exchange', 'EXO', 'OneDrive', 'ODB', 'SharePoint', 'SPO', 'Teams')
        if ([string]$TaskData.Workload -notin $allowedWorkloads) {
            throw "Invalid Workload '$($TaskData.Workload)'."
        }

        if ($phase -eq 'Detail') {
            $allowedLocationTypes = @('', 'SiteUrl', 'UPN', 'WorkloadFallback')
            $locationType = if ($TaskData.ContainsKey('LocationType') -and $TaskData.LocationType) { [string]$TaskData.LocationType } else { '' }
            if ($locationType -notin $allowedLocationTypes) {
                throw "Invalid LocationType '$locationType'."
            }
        }
    }

    if ($phase -eq 'AEExport') {
        if ([string]::IsNullOrWhiteSpace([string]$TaskData.Day)) { throw "AE task missing Day." }
        if ([string]::IsNullOrWhiteSpace([string]$TaskData.StartTime)) { throw "AE task missing StartTime." }
        if ([string]::IsNullOrWhiteSpace([string]$TaskData.EndTime)) { throw "AE task missing EndTime." }
    }

    return $true
}

function Send-WorkerTask {
    <#
    .SYNOPSIS
        Assigns a task to a worker by writing nexttask file. Skips if nexttask already exists.
    .OUTPUTS
        $true if task was assigned, $false if worker already has a pending task.
    #>
    param(
        [Parameter(Mandatory)][string]$WorkerDir,
        [Parameter(Mandatory)][hashtable]$TaskData,
        [Parameter(Mandatory)][string]$ExportDir
    )
    $nextTaskPath = Join-Path $WorkerDir "nexttask"
    $signingKey = Get-ExportRunSigningKey -ExportDir $ExportDir -CreateIfMissing
    # Use existence check via file open rather than Test-Path to avoid TOCTOU race.
    # If the file already exists, the worker hasn't consumed the previous task yet.
    try {
        # CreateNew fails if file exists — atomic check-and-create
        $fs = [System.IO.File]::Open($nextTaskPath, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
        try {
            $json = ConvertTo-SignedEnvelopeJson -Payload $TaskData -SigningKey $signingKey
            $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
            $fs.Write($bytes, 0, $bytes.Length)
        }
        finally {
            $fs.Dispose()
        }
        return $true
    }
    catch [System.IO.IOException] {
        # File already exists (worker hasn't consumed previous task) or write failed
        return $false
    }
}

function Receive-WorkerTask {
    <#
    .SYNOPSIS
        Worker reads and acknowledges a task by renaming nexttask to currenttask.
    .OUTPUTS
        Task hashtable, or $null if no task available.
    #>
    param(
        [Parameter(Mandatory)][string]$WorkerDir,
        [Parameter(Mandatory)][string]$ExportDir
    )
    $nextTaskPath = Join-Path $WorkerDir "nexttask"
    $currentTaskPath = Join-Path $WorkerDir "currenttask"
    $signingKey = Get-ExportRunSigningKey -ExportDir $ExportDir
    $requireSignature = -not [string]::IsNullOrWhiteSpace($signingKey)
    try {
        $json = [System.IO.File]::ReadAllText($nextTaskPath)
    }
    catch [System.IO.FileNotFoundException], [System.IO.DirectoryNotFoundException] {
        # No task available — normal in concurrent operation
        return $null
    }
    catch {
        Write-Verbose "Receive-WorkerTask: Unexpected error reading nexttask: $($_.Exception.Message)"
        return $null
    }

    try {
        $taskData = ConvertFrom-SignedEnvelopeJson -Json $json -SigningKey $signingKey -RequireSignature:$requireSignature -Context "worker task"
        if ($null -eq $taskData) {
            throw "Task payload is null"
        }
        Test-WorkerTaskSchema -TaskData $taskData | Out-Null
    }
    catch {
        # Quarantine malformed or tampered task payloads so worker does not get stuck on the same file.
        $invalidPath = Join-Path $WorkerDir ("invalidtask-{0}-{1}.json" -f (Get-Date -Format "yyyyMMdd-HHmmssfff"), $PID)
        try {
            [System.IO.File]::Move($nextTaskPath, $invalidPath, $true)
        }
        catch {
            try { [System.IO.File]::Delete($nextTaskPath) } catch { }
        }
        Write-Warning "Receive-WorkerTask: Rejected invalid nexttask payload: $($_.Exception.Message)"
        return $null
    }

    try {
        # Rename to acknowledge after successful parse/validation.
        [System.IO.File]::Move($nextTaskPath, $currentTaskPath, $true)
    }
    catch [System.IO.FileNotFoundException], [System.IO.DirectoryNotFoundException] {
        # nexttask vanished between parse and rename — another process consumed it
        return $null
    }
    catch {
        Write-Verbose "Receive-WorkerTask: Could not rename nexttask to currenttask: $($_.Exception.Message)"
        return $null
    }

    return $taskData
}

function Complete-WorkerTask {
    <#
    .SYNOPSIS
        Worker signals task completion by deleting currenttask file.
    #>
    param([Parameter(Mandatory)][string]$WorkerDir)
    $currentTaskPath = Join-Path $WorkerDir "currenttask"
    try {
        [System.IO.File]::Delete($currentTaskPath)
    }
    catch [System.IO.FileNotFoundException], [System.IO.DirectoryNotFoundException] {
        # File already removed — normal in concurrent operation
    }
    catch {
        Write-Verbose "Complete-WorkerTask: Could not delete currenttask: $($_.Exception.Message)"
    }
}

#endregion

