#region Content Explorer - Aggregate CSV Discovery & Reuse

function Find-RecentAggregateCsv {
    <#
    .SYNOPSIS
        Scans Output directory for recent Content Explorer aggregate CSV files.
    .DESCRIPTION
        Looks for ContentExplorer-Aggregates.csv files in Export-* subdirectories.
        Reads AggregateMetadata.json (if present) for tenant matching.
        Returns results sorted by newest first.
    .PARAMETER OutputDirectory
        Base output directory to scan for Export-* subfolders.
    .PARAMETER MaxAgeDays
        Maximum age in days for aggregate files to be considered recent. Default: 30.
    .PARAMETER TenantId
        Optional tenant ID filter. When provided, only returns aggregates from
        matching tenants (or those without tenant metadata).
    .OUTPUTS
        Array of objects with: Path, FolderName, RecordCount, AgeHours, AgeDays,
        TenantDomain, TenantId. Sorted newest first.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$OutputDirectory,

        [int]$MaxAgeDays = 30,

        [string]$TenantId
    )

    $results = [System.Collections.ArrayList]::new()

    if (-not (Test-Path $OutputDirectory)) {
        return @()
    }

    $exportFolders = Get-ChildItem -Path $OutputDirectory -Directory -Filter "Export-*" -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending

    foreach ($folder in $exportFolders) {
        $csvPath = Join-Path (Get-CoordinationDir $folder.FullName) "ContentExplorer-Aggregates.csv"
        if (-not (Test-Path $csvPath)) { continue }

        $csvFile = Get-Item $csvPath
        $age = (Get-Date) - $csvFile.LastWriteTime
        if ($age.TotalDays -gt $MaxAgeDays) { continue }

        # Count data records (exclude header and error rows)
        $recordCount = 0
        try {
            $lines = Get-Content -Path $csvPath -Encoding UTF8 -ErrorAction Stop
            # Skip header line, count non-empty data lines
            $recordCount = @($lines | Select-Object -Skip 1 | Where-Object { $_ -match '\S' }).Count
        }
        catch {
            Write-Verbose "Failed to read aggregate CSV at $csvPath : $($_.Exception.Message)"
            continue
        }

        # Read tenant metadata if available
        $tenantDomain = $null
        $tenantIdValue = $null
        $metadataPath = Join-Path (Get-CoordinationDir $folder.FullName) "AggregateMetadata.json"
        if (Test-Path $metadataPath) {
            try {
                $metadata = Get-Content -Raw -Path $metadataPath -ErrorAction Stop | ConvertFrom-Json
                if ($null -ne $metadata) {
                    $tenantDomain = $metadata.TenantDomain
                    $tenantIdValue = $metadata.TenantId
                }
            }
            catch {
                Write-Verbose "Failed to read metadata at $metadataPath : $($_.Exception.Message)"
            }
        }

        # Apply tenant filter if specified
        if ($TenantId -and (-not $tenantIdValue -or $tenantIdValue -ne $TenantId)) {
            continue
        }

        $entry = [PSCustomObject]@{
            Path         = $csvPath
            FolderName   = $folder.Name
            RecordCount  = $recordCount
            AgeHours     = [Math]::Round($age.TotalHours, 1)
            AgeDays      = [Math]::Round($age.TotalDays, 1)
            TenantDomain = $tenantDomain
            TenantId     = $tenantIdValue
        }
        [void]$results.Add($entry)
    }

    # Already sorted by newest first via folder enumeration
    return @($results)
}

function Save-AggregateMetadata {
    <#
    .SYNOPSIS
        Saves tenant info alongside aggregate CSV for future reuse matching.
    .DESCRIPTION
        Writes AggregateMetadata.json to the export directory containing tenant
        domain, tenant ID, and timestamp information.
    .PARAMETER ExportRunDirectory
        The export run directory where the metadata file will be written.
    .PARAMETER TenantInfo
        Hashtable with TenantDomain and TenantId keys.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ExportRunDirectory,

        [Parameter(Mandatory)]
        [hashtable]$TenantInfo
    )

    $coordDir = Get-CoordinationDir $ExportRunDirectory
    if (-not (Test-Path $coordDir)) { New-Item -ItemType Directory -Force -Path $coordDir | Out-Null }
    $metadataPath = Join-Path $coordDir "AggregateMetadata.json"

    $metadata = [ordered]@{
        TenantDomain = $TenantInfo.TenantDomain
        TenantId     = $TenantInfo.TenantId
        CreatedTime  = (Get-Date).ToString("o")
        ExportFolder = Split-Path $ExportRunDirectory -Leaf
    }

    try {
        $json = $metadata | ConvertTo-Json -Depth 20
        Set-Content -Path $metadataPath -Value $json -Encoding UTF8
        Write-ExportLog -Message "Saved aggregate metadata: $metadataPath" -Level Info
    }
    catch {
        Write-ExportLog -Message ("Failed to save aggregate metadata: " + $_.Exception.Message) -Level Warning
    }
}

function Save-ExportSettings {
    <#
    .SYNOPSIS
        Saves an ExportSettings.json manifest to the export directory.
    .DESCRIPTION
        Writes a settings manifest that captures configuration at export start time.
        On resume, this manifest is reloaded so settings remain consistent even if
        config files have changed on disk.
    .PARAMETER ExportRunDirectory
        The export run directory where ExportSettings.json will be written.
    .PARAMETER ExportType
        The export type: "ContentExplorer" or "ActivityExplorer".
    .PARAMETER Settings
        Hashtable of key-value settings to persist in the manifest.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ExportRunDirectory,

        [Parameter(Mandatory)]
        [ValidateSet("ContentExplorer", "ActivityExplorer")]
        [string]$ExportType,

        [Parameter(Mandatory)]
        [hashtable]$Settings
    )

    $coordDir = Get-CoordinationDir $ExportRunDirectory
    if (-not (Test-Path $coordDir)) { New-Item -ItemType Directory -Force -Path $coordDir | Out-Null }
    $settingsPath = Join-Path $coordDir "ExportSettings.json"

    $manifest = [ordered]@{
        ExportType  = $ExportType
        CreatedTime = (Get-Date).ToUniversalTime().ToString("o")
    }

    # Merge all settings into the manifest
    foreach ($key in $Settings.Keys) {
        $manifest[$key] = $Settings[$key]
    }

    try {
        $tempPath = "$settingsPath.tmp"
        $json = $manifest | ConvertTo-Json -Depth 20
        Set-Content -Path $tempPath -Value $json -Encoding UTF8
        Move-Item -Path $tempPath -Destination $settingsPath -Force
        Write-ExportLog -Message "Saved export settings manifest: $settingsPath" -Level Info
    }
    catch {
        Write-ExportLog -Message ("Failed to save export settings: " + $_.Exception.Message) -Level Warning
    }
}

function Get-ExportSettings {
    <#
    .SYNOPSIS
        Reads the ExportSettings.json manifest from the export directory.
    .DESCRIPTION
        Returns the parsed settings object, or $null if no manifest exists.
        Callers should fall back to config-file behavior when $null is returned
        (backward compatibility with exports created before the manifest feature).
    .PARAMETER ExportRunDirectory
        The export run directory to read ExportSettings.json from.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ExportRunDirectory
    )

    $settingsPath = Join-Path (Get-CoordinationDir $ExportRunDirectory) "ExportSettings.json"

    try {
        $content = Get-Content -Raw -Path $settingsPath -ErrorAction Stop
        $settings = ConvertFrom-Json -InputObject $content -ErrorAction Stop
        return $settings
    }
    catch [System.Management.Automation.ItemNotFoundException] {
        return $null
    }
    catch {
        Write-ExportLog -Message ("Failed to read export settings: " + $_.Exception.Message) -Level Warning
        return $null
    }
}

function Resolve-CEPageSize {
    <#
    .SYNOPSIS
        Resolves Content Explorer page size from saved manifest or config file, with fallback default.
    .PARAMETER ExportRunDirectory
        The export run directory to check for ExportSettings.json.
    .PARAMETER ConfigPath
        Path to ContentExplorerClassifiers.json config file.
    .PARAMETER FallbackPageSize
        Default page size if neither manifest nor config provides one.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ExportRunDirectory,

        [string]$ConfigPath,

        [int]$FallbackPageSize = 100
    )

    $savedSettings = Get-ExportSettings -ExportRunDirectory $ExportRunDirectory
    if ($savedSettings) {
        Write-ExportLog -Message "Loaded export settings from ExportSettings.json" -Level Info
    }
    else {
        Write-ExportLog -Message "No ExportSettings.json found - using current config files" -Level Warning
    }

    $ceConfig = if ($ConfigPath) { Read-JsonConfig -Path $ConfigPath } else { $null }
    $ceSettings = Get-ContentExplorerSettings -ConfigObject $ceConfig -SavedSettings $savedSettings -DefaultPageSize $FallbackPageSize
    $cePageSize = $ceSettings.PageSize

    return [PSCustomObject]@{
        PageSize      = $cePageSize
        Settings      = $ceSettings
        SavedSettings = $savedSettings
        CeConfig      = $ceConfig
    }
}

function Resolve-AEFilters {
    <#
    .SYNOPSIS
        Resolves Activity Explorer filters from saved manifest or config file.
    .PARAMETER ExportRunDirectory
        The export run directory to check for ExportSettings.json.
    .PARAMETER ConfigPath
        Path to ActivityExplorerSelector.json config file.
    .PARAMETER LogDetails
        When set, logs filter details via Write-ExportLog.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ExportRunDirectory,

        [string]$ConfigPath,

        [switch]$LogDetails
    )

    $savedSettings = Get-ExportSettings -ExportRunDirectory $ExportRunDirectory
    if ($savedSettings -and $savedSettings.SelectorConfig) {
        $filters = Get-ActivityExplorerFilters -ConfigObject $savedSettings.SelectorConfig -LogDetails:$LogDetails
        Write-ExportLog -Message "Using saved filter settings from ExportSettings.json" -Level Info
    }
    elseif ($ConfigPath) {
        $filters = Get-ActivityExplorerFilters -ConfigPath $ConfigPath -LogDetails:$LogDetails
        Write-ExportLog -Message "No saved settings found - using current config file" -Level Warning
    }
    else {
        Write-ExportLog -Message "No saved settings or config path - exporting all activities" -Level Warning
        $filters = $null
    }

    return $filters
}

function Get-TagNamesFromAggregateCsv {
    <#
    .SYNOPSIS
        Extracts unique tag names from an aggregate CSV file.
    .DESCRIPTION
        Reads the ContentExplorer-Aggregates.csv, parses the TagName column,
        and returns unique values optionally filtered by TagType.
    .PARAMETER CsvPath
        Path to the ContentExplorer-Aggregates.csv file.
    .PARAMETER TagType
        Optional filter to return only tag names of the specified type.
    .OUTPUTS
        Array of unique tag name strings.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$CsvPath,

        [string]$TagType
    )

    if (-not (Test-Path $CsvPath)) {
        Write-ExportLog -Message "Aggregate CSV not found: $CsvPath" -Level Warning
        return @()
    }

    try {
        $csvData = Import-Csv -Path $CsvPath -Encoding UTF8 -ErrorAction Stop

        if ($TagType) {
            $filtered = @($csvData | Where-Object { $_.TagType -eq $TagType })
        }
        else {
            $filtered = @($csvData)
        }

        $tagNames = @($filtered | ForEach-Object { $_.TagName } | Where-Object { $_ } | Select-Object -Unique | Sort-Object)
        return $tagNames
    }
    catch {
        Write-ExportLog -Message ("Failed to read aggregate CSV: " + $_.Exception.Message) -Level Error
        return @()
    }
}

function Import-AggregateDataFromCsv {
    <#
    .SYNOPSIS
        Loads aggregate data from CSV into a structured task data format.
    .DESCRIPTION
        Reads ContentExplorer-Aggregates.csv and builds a hashtable keyed by
        "TagType|TagName|Workload" with location arrays and total counts.
        Error rows (Location=ERROR) are tracked separately.
    .PARAMETER CsvPath
        Path to the ContentExplorer-Aggregates.csv file.
    .PARAMETER TagType
        Optional filter for a specific tag type.
    .PARAMETER TagNames
        Optional array of tag names to include.
    .PARAMETER Workloads
        Optional array of workloads to include.
    .OUTPUTS
        Hashtable with:
          TaskData  - Hashtable keyed by "TagType|TagName|Workload", each with
                      TagType, TagName, Workload, Locations array, TotalCount, HasError
          HasErrors - Boolean indicating if any error rows were found
          ErrorTasks - Array of task keys that had errors
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$CsvPath,

        [string]$TagType,

        [string[]]$TagNames,

        [string[]]$Workloads
    )

    $result = @{
        TaskData   = @{}
        HasErrors  = $false
        ErrorTasks = @()
    }

    if (-not (Test-Path $CsvPath)) {
        Write-ExportLog -Message "Aggregate CSV not found: $CsvPath" -Level Warning
        return $result
    }

    try {
        $csvData = Import-Csv -Path $CsvPath -Encoding UTF8 -ErrorAction Stop
    }
    catch {
        Write-ExportLog -Message ("Failed to read aggregate CSV: " + $_.Exception.Message) -Level Error
        return $result
    }

    foreach ($row in $csvData) {
        # Apply filters
        if ($TagType -and $row.TagType -ne $TagType) { continue }
        if ($TagNames -and $TagNames.Count -gt 0 -and $row.TagName -notin $TagNames) { continue }
        if ($Workloads -and $Workloads.Count -gt 0 -and $row.Workload -notin $Workloads) { continue }

        $taskKey = "{0}|{1}|{2}" -f $row.TagType, $row.TagName, $row.Workload

        # Initialize task entry if needed
        if (-not $result.TaskData.ContainsKey($taskKey)) {
            $result.TaskData[$taskKey] = @{
                TagType    = $row.TagType
                TagName    = $row.TagName
                Workload   = $row.Workload
                Locations  = [System.Collections.ArrayList]::new()
                TotalCount = 0
                HasError   = $false
            }
        }

        $taskEntry = $result.TaskData[$taskKey]

        # Check for error rows
        if ($row.Location -eq "ERROR") {
            $taskEntry.HasError = $true
            $result.HasErrors = $true
            if ($taskKey -notin $result.ErrorTasks) {
                $result.ErrorTasks += $taskKey
            }
            continue
        }

        # _FILECOUNT row: probed file count from detail API (more accurate than match count)
        if ($row.Location -eq "_FILECOUNT") {
            $fc = $row.Count -as [int]
            if ($fc -and $fc -gt 0) {
                $taskEntry.FileCount = $fc
            }
            continue
        }

        # Skip NONE marker rows (zero-result tasks)
        if ($row.Location -eq "NONE") { continue }

        # Add location data (these counts are match counts from aggregate API)
        $count = 0
        if ($row.Count) {
            $count = $row.Count -as [int]
            if ($null -eq $count) { $count = 0 }
        }

        [void]$taskEntry.Locations.Add(@{
            Name          = $row.Location
            ExpectedCount = $count
            ExportedCount = 0
        })
        $taskEntry.MatchCount = ($taskEntry.MatchCount -as [int]) + $count
    }

    # Finalize TotalCount: prefer FileCount (actual files) over MatchCount (aggregate matches)
    foreach ($taskKey in $result.TaskData.Keys) {
        $entry = $result.TaskData[$taskKey]
        if ($entry.FileCount -and $entry.FileCount -gt 0) {
            $entry.TotalCount = $entry.FileCount
        } else {
            $entry.TotalCount = $entry.MatchCount -as [int]
        }
    }

    return $result
}

#endregion

