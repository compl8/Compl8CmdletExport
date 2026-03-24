#region Export Directory Path Helpers

function Get-LegacySafeDirectoryName {
    <#
    .SYNOPSIS
        Returns the legacy filesystem-safe name used by older export runs.
    #>
    param([Parameter(Mandatory)][string]$Name)

    $safe = $Name -replace '[\\/:*?"<>|]', '_'
    $safe = $safe -replace '[\x00-\x1f]', ''
    $safe = $safe.Trim('. ')
    if ($safe -match '^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])$') { $safe = "_$safe" }
    if ($safe.Length -gt 200) { $safe = $safe.Substring(0, 200) }
    if (-not $safe) { $safe = "_unnamed" }
    return $safe
}

function Get-DeterministicNameHash {
    <#
    .SYNOPSIS
        Returns a short deterministic hash for disambiguating sanitized names.
    #>
    param([Parameter(Mandatory)][string]$Name)

    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($Name)
        $hash = $sha.ComputeHash($bytes)
        return ([Convert]::ToHexString($hash)).ToLowerInvariant().Substring(0, 10)
    }
    finally {
        $sha.Dispose()
    }
}

function ConvertTo-SafeDirectoryName {
    <#
    .SYNOPSIS
        Converts a string to a safe directory name and adds a stable hash when
        sanitization or truncation would otherwise create collisions.
    #>
    param([Parameter(Mandatory)][string]$Name)

    $legacySafe = Get-LegacySafeDirectoryName -Name $Name
    if ($legacySafe -eq $Name) {
        return $legacySafe
    }

    $hashSuffix = "~" + (Get-DeterministicNameHash -Name $Name)
    $maxBaseLength = 200 - $hashSuffix.Length
    $baseName = $legacySafe
    if ($baseName.Length -gt $maxBaseLength) {
        $baseName = $baseName.Substring(0, $maxBaseLength).TrimEnd('. ')
    }
    if (-not $baseName) {
        $baseName = "_"
    }

    return $baseName + $hashSuffix
}

function Get-CoordinationDir {
    <#
    .SYNOPSIS
        Returns the _Coordination/ subdirectory path for an export directory.
    #>
    param([Parameter(Mandatory)][string]$ExportDir)
    return (Join-Path $ExportDir "_Coordination")
}

function Get-CompletionsDir {
    <#
    .SYNOPSIS
        Returns the _Coordination/Completions/ subdirectory path.
    #>
    param([Parameter(Mandatory)][string]$ExportDir)
    return (Join-Path $ExportDir "_Coordination" "Completions")
}

function Get-WorkerCoordDir {
    <#
    .SYNOPSIS
        Returns the _Coordination/Workers/PID/ subdirectory path for a specific worker.
    #>
    param(
        [Parameter(Mandatory)][string]$ExportDir,
        [Parameter(Mandatory)][string]$WorkerPID
    )
    return (Join-Path $ExportDir "_Coordination" "Workers" $WorkerPID)
}

function Get-LogsDir {
    <#
    .SYNOPSIS
        Returns the _Logs/ subdirectory path for an export directory.
    #>
    param([Parameter(Mandatory)][string]$ExportDir)
    return (Join-Path $ExportDir "_Logs")
}

function Get-CEDataDir {
    <#
    .SYNOPSIS
        Returns the Data/ContentExplorer/ subdirectory path.
    #>
    param([Parameter(Mandatory)][string]$ExportDir)
    return (Join-Path $ExportDir "Data" "ContentExplorer")
}

function Get-AEDataDir {
    <#
    .SYNOPSIS
        Returns the Data/ActivityExplorer/ subdirectory path.
    #>
    param([Parameter(Mandatory)][string]$ExportDir)
    return (Join-Path $ExportDir "Data" "ActivityExplorer")
}

function Get-CEClassifierDir {
    <#
    .SYNOPSIS
        Returns the Data/ContentExplorer/TagType/TagName/ subdirectory path.
        Tag names are sanitized to be filesystem-safe.
    #>
    param(
        [Parameter(Mandatory)][string]$ExportDir,
        [Parameter(Mandatory)][string]$TagType,
        [Parameter(Mandatory)][string]$TagName
    )
    $safeTagType = ConvertTo-SafeDirectoryName -Name $TagType
    $safeTagName = ConvertTo-SafeDirectoryName -Name $TagName
    $resolvedPath = Join-Path $ExportDir "Data" "ContentExplorer" $safeTagType $safeTagName

    # Preserve resumability for older exports created before hashed safe names.
    $legacyTagType = Get-LegacySafeDirectoryName -Name $TagType
    $legacyTagName = Get-LegacySafeDirectoryName -Name $TagName
    $legacyPath = Join-Path $ExportDir "Data" "ContentExplorer" $legacyTagType $legacyTagName
    if ($legacyPath -ne $resolvedPath -and (Test-Path $legacyPath) -and -not (Test-Path $resolvedPath)) {
        return $legacyPath
    }

    return $resolvedPath
}

function Get-AEDayDir {
    <#
    .SYNOPSIS
        Returns the Data/ActivityExplorer/YYYY-MM-DD/ subdirectory path.
    #>
    param(
        [Parameter(Mandatory)][string]$ExportDir,
        [Parameter(Mandatory)][string]$Day
    )
    return (Join-Path $ExportDir "Data" "ActivityExplorer" $Day)
}

function Initialize-ExportDirectories {
    <#
    .SYNOPSIS
        Creates the standard export directory structure upfront.
    .PARAMETER ExportDir
        Root export directory.
    .PARAMETER ExportType
        'ContentExplorer', 'ActivityExplorer', or 'Full'.
    #>
    param(
        [Parameter(Mandatory)][string]$ExportDir,
        [Parameter(Mandatory)][ValidateSet('ContentExplorer','ActivityExplorer','Full')][string]$ExportType
    )

    $dirs = @(
        (Get-CoordinationDir $ExportDir),
        (Get-CompletionsDir $ExportDir),
        (Join-Path (Get-CoordinationDir $ExportDir) "Workers"),
        (Get-LogsDir $ExportDir)
    )

    if ($ExportType -in @('ContentExplorer', 'Full')) {
        $dirs += (Get-CEDataDir $ExportDir)
    }
    if ($ExportType -in @('ActivityExplorer', 'Full')) {
        $dirs += (Get-AEDataDir $ExportDir)
    }

    foreach ($dir in $dirs) {
        if (-not (Test-Path $dir)) {
            New-Item -ItemType Directory -Force -Path $dir | Out-Null
        }
    }
}

function Write-CEManifest {
    <#
    .SYNOPSIS
        Writes a _manifest.json summary for Content Explorer data.
    .DESCRIPTION
        Scans Data/ContentExplorer/ for _task-*.json files and aggregates
        into a top-level manifest with tag types, classifiers, and record counts.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ExportDir
    )

    $ceDataDir = Get-CEDataDir $ExportDir
    if (-not (Test-Path $ceDataDir)) { return }

    # Scan for _task-*.json summaries
    $taskFiles = @(Get-ChildItem -Path $ceDataDir -Recurse -Filter "_task-*.json" -ErrorAction SilentlyContinue)
    if ($taskFiles.Count -eq 0) { return }

    $tagTypes = @{}
    $totalRecords = [long]0
    $totalPages = 0

    foreach ($taskFile in $taskFiles) {
        try {
            $task = Get-Content -Path $taskFile.FullName -Raw -ErrorAction Stop | ConvertFrom-Json
            $tt = $task.TagType
            $tn = $task.TagName

            if (-not $tagTypes.ContainsKey($tt)) {
                $tagTypes[$tt] = @{}
            }
            if (-not $tagTypes[$tt].ContainsKey($tn)) {
                $tagTypes[$tt][$tn] = @{ Workloads = @{}; TotalRecords = [long]0; TotalPages = 0 }
            }

            $wl = $task.Workload
            $count = ($task.ActualCount -as [long])
            $pages = ($task.Pages -as [int])

            $tagTypes[$tt][$tn].Workloads[$wl] = @{ Records = $count; Pages = $pages; Status = $task.Status }
            $tagTypes[$tt][$tn].TotalRecords += $count
            $tagTypes[$tt][$tn].TotalPages += $pages
            $totalRecords += $count
            $totalPages += $pages
        }
        catch {
            Write-Verbose "Skipping malformed task file: $($taskFile.FullName)"
        }
    }

    # Build manifest
    $classifiers = @()
    foreach ($tt in $tagTypes.Keys | Sort-Object) {
        foreach ($tn in $tagTypes[$tt].Keys | Sort-Object) {
            $entry = $tagTypes[$tt][$tn]
            $classifiers += @{
                TagType      = $tt
                TagName      = $tn
                TotalRecords = $entry.TotalRecords
                TotalPages   = $entry.TotalPages
                Workloads    = $entry.Workloads
            }
        }
    }

    $manifest = @{
        ExportType     = "ContentExplorer"
        ExportDate     = (Get-Date).ToString("o")
        TagTypes       = @($tagTypes.Keys | Sort-Object)
        ClassifierCount = $classifiers.Count
        TotalRecords   = $totalRecords
        TotalPages     = $totalPages
        Classifiers    = @($classifiers)
    }

    $manifestPath = Join-Path $ceDataDir "_manifest.json"
    try {
        $manifest | ConvertTo-Json -Depth 10 | Set-Content -Path $manifestPath -Encoding UTF8
        Write-ExportLog -Message ("CE manifest written: {0} classifiers, {1} records, {2} pages" -f $classifiers.Count, $totalRecords, $totalPages) -Level Info
    }
    catch {
        Write-ExportLog -Message ("Failed to write CE manifest: " + $_.Exception.Message) -Level Warning
    }
}

function Write-AEManifest {
    <#
    .SYNOPSIS
        Writes a _manifest.json summary for Activity Explorer data.
    .DESCRIPTION
        Scans Data/ActivityExplorer/ for day directories and Page-*.json files,
        aggregates into a top-level manifest with days, record counts, and page counts.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ExportDir
    )

    $aeDataDir = Get-AEDataDir $ExportDir
    if (-not (Test-Path $aeDataDir)) { return }

    # Scan day directories
    $dayDirs = @(Get-ChildItem -Path $aeDataDir -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '^\d{4}-\d{2}-\d{2}$' } |
        Sort-Object Name)

    $days = @()
    $totalRecords = [long]0
    $totalPages = 0

    foreach ($dayDir in $dayDirs) {
        $pageFiles = @(Get-ChildItem -Path $dayDir.FullName -Filter "Page-*.json" -ErrorAction SilentlyContinue)
        $dayRecords = [long]0

        foreach ($pf in $pageFiles) {
            try {
                # Read only the first few lines to extract RecordCount without parsing entire file
                $head = Get-Content -Path $pf.FullName -TotalCount 10 -ErrorAction Stop
                $match = ($head -join "`n") | Select-String -Pattern '"RecordCount"\s*:\s*(\d+)'
                if ($match) {
                    $dayRecords += ($match.Matches[0].Groups[1].Value -as [long])
                }
            }
            catch {
                Write-Verbose "Skipping malformed page file: $($pf.FullName)"
            }
        }

        $days += @{
            Day         = $dayDir.Name
            RecordCount = $dayRecords
            PageCount   = $pageFiles.Count
        }
        $totalRecords += $dayRecords
        $totalPages += $pageFiles.Count
    }

    $manifest = @{
        ExportType   = "ActivityExplorer"
        ExportDate   = (Get-Date).ToString("o")
        DaysExported = $days.Count
        TotalRecords = $totalRecords
        TotalPages   = $totalPages
        Days         = @($days)
    }

    $manifestPath = Join-Path $aeDataDir "_manifest.json"
    try {
        $manifest | ConvertTo-Json -Depth 10 | Set-Content -Path $manifestPath -Encoding UTF8
        Write-ExportLog -Message ("AE manifest written: {0} days, {1} records, {2} pages" -f $days.Count, $totalRecords, $totalPages) -Level Info
    }
    catch {
        Write-ExportLog -Message ("Failed to write AE manifest: " + $_.Exception.Message) -Level Warning
    }
}

#endregion

