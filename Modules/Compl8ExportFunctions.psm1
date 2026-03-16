#Requires -Version 7.0
<#
.SYNOPSIS
    PowerShell module for Compl8 Cmdlet Export functions.

.DESCRIPTION
    This module provides functions for connecting to Security & Compliance PowerShell,
    exporting compliance configuration data, and handling pagination for large datasets.

    Based on current Microsoft Learn documentation:
    - https://learn.microsoft.com/powershell/module/exchange/connect-ippssession
    - https://learn.microsoft.com/powershell/module/exchange/export-contentexplorerdata
    - https://learn.microsoft.com/powershell/module/exchange/export-activityexplorerdata

.NOTES
    Version: 1.0.0
    Requires: PowerShell 7+, ExchangeOnlineManagement module 3.2.0+
#>

#region Script Variables
$script:LogFile = $null
$script:SessionStartTime = $null
$script:ExportStats = @{
    ItemsExported = @{}
    Errors = [System.Collections.ArrayList]::new()
    Warnings = [System.Collections.ArrayList]::new()
}
$script:DashboardLineCount = 0
$script:AEDashboardLineCount = 0
#endregion

#region Logging Functions

function Initialize-ExportLog {
    <#
    .SYNOPSIS
        Initializes a log file for the export session.
    .PARAMETER LogDirectory
        Directory where log files will be created.
    .PARAMETER Prefix
        Prefix for the log file name.
    .OUTPUTS
        String path to the created log file.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$LogDirectory,

        [string]$Prefix = "Compl8Export"
    )

    if (-not (Test-Path $LogDirectory)) {
        New-Item -ItemType Directory -Force -Path $LogDirectory | Out-Null
    }

    $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $script:LogFile = Join-Path $LogDirectory "$Prefix-$timestamp.log"
    $script:SessionStartTime = Get-Date

    $header = @"
================================================================================
Compl8 Cmdlet Export
Started: $($script:SessionStartTime.ToString("yyyy-MM-dd HH:mm:ss"))
PowerShell Version: $($PSVersionTable.PSVersion)
================================================================================

"@
    Set-Content -Path $script:LogFile -Value $header
    return $script:LogFile
}

function Write-ExportLog {
    <#
    .SYNOPSIS
        Writes a log message to console and log file.
    .PARAMETER Message
        The message to log.
    .PARAMETER Level
        Log level: Info, Warning, Error, Success.
    .PARAMETER NoNewline
        Don't add newline (for status updates).
    .PARAMETER LogOnly
        Write to log file only (skip console output). Useful when a dashboard is managing console display.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet("Info", "Warning", "Error", "Success")]
        [string]$Level = "Info",

        [switch]$NoNewline,

        [switch]$LogOnly
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"

    # Console output (unless LogOnly)
    if (-not $LogOnly) {
        $color = switch ($Level) {
            "Info"    { "White" }
            "Warning" { "Yellow" }
            "Error"   { "Red" }
            "Success" { "Green" }
        }

        if ($NoNewline) {
            Write-Host $Message -ForegroundColor $color -NoNewline
        }
        else {
            Write-Host $Message -ForegroundColor $color
        }
    }

    # Write to log file
    if ($script:LogFile -and (Test-Path $script:LogFile -ErrorAction SilentlyContinue)) {
        Add-Content -Path $script:LogFile -Value $logEntry
    }

    # Track errors and warnings
    if ($Level -eq "Error") {
        [void]$script:ExportStats.Errors.Add(@{ Timestamp = $timestamp; Message = $Message })
    }
    elseif ($Level -eq "Warning") {
        [void]$script:ExportStats.Warnings.Add(@{ Timestamp = $timestamp; Message = $Message })
    }
}

function Get-ExportStatistics {
    <#
    .SYNOPSIS
        Returns statistics from the current export session.
    #>
    [CmdletBinding()]
    param()

    return @{
        StartTime = $script:SessionStartTime
        EndTime = Get-Date
        Duration = if ($script:SessionStartTime) { (Get-Date) - $script:SessionStartTime } else { $null }
        ItemsExported = $script:ExportStats.ItemsExported.Clone()
        ErrorCount = $script:ExportStats.Errors.Count
        WarningCount = $script:ExportStats.Warnings.Count
        Errors = $script:ExportStats.Errors.ToArray()
        Warnings = $script:ExportStats.Warnings.ToArray()
        LogFile = $script:LogFile
    }
}

function Add-ExportCount {
    <#
    .SYNOPSIS
        Tracks the count of exported items by category.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Category,

        [Parameter(Mandatory)]
        [int]$Count
    )

    if ($script:ExportStats.ItemsExported.ContainsKey($Category)) {
        $script:ExportStats.ItemsExported[$Category] += $Count
    }
    else {
        $script:ExportStats.ItemsExported[$Category] = $Count
    }
}

#endregion

#region Connection Functions

function Test-ExportPrerequisites {
    <#
    .SYNOPSIS
        Verifies PowerShell version and required modules.
    .PARAMETER RequiredModules
        Array of required module names.
    .OUTPUTS
        Boolean indicating if prerequisites are met.
    #>
    [CmdletBinding()]
    param(
        [string[]]$RequiredModules = @("ExchangeOnlineManagement")
    )

    $passed = $true

    # Check PowerShell version (7+ required for modern features)
    Write-ExportLog -Message "Checking PowerShell version... " -Level Info -NoNewline
    if ($PSVersionTable.PSVersion.Major -ge 7) {
        Write-ExportLog -Message "OK (v$($PSVersionTable.PSVersion))" -Level Success
    }
    else {
        Write-ExportLog -Message "FAILED (v$($PSVersionTable.PSVersion), need 7+)" -Level Error
        $passed = $false
    }

    # Check required modules
    foreach ($moduleName in $RequiredModules) {
        Write-ExportLog -Message "Checking module '$moduleName'... " -Level Info -NoNewline
        $module = Get-Module -ListAvailable -Name $moduleName | Sort-Object Version -Descending | Select-Object -First 1

        if ($module) {
            $version = $module.Version
            # ExchangeOnlineManagement 3.2.0+ required for REST API mode
            if ($moduleName -eq "ExchangeOnlineManagement" -and $version -lt [version]"3.2.0") {
                Write-ExportLog -Message "OUTDATED (v$version, need 3.2.0+)" -Level Warning
                Write-ExportLog -Message "  Update with: Install-Module ExchangeOnlineManagement -Force" -Level Info
            }
            else {
                Write-ExportLog -Message "OK (v$version)" -Level Success
            }
        }
        else {
            Write-ExportLog -Message "NOT FOUND" -Level Error
            Write-ExportLog -Message "  Install with: Install-Module $moduleName -Scope CurrentUser" -Level Info
            $passed = $false
        }
    }

    return $passed
}

function Connect-Compl8Compliance {
    <#
    .SYNOPSIS
        Connects to Security & Compliance PowerShell.

    .DESCRIPTION
        Establishes a connection to Microsoft Purview Security & Compliance PowerShell
        using modern authentication (REST API mode, no WinRM required).

        Reference: https://learn.microsoft.com/powershell/exchange/connect-to-scc-powershell

    .PARAMETER UserPrincipalName
        UPN for interactive authentication (optional, enables pre-filled username).

    .PARAMETER AppId
        Application ID for certificate-based authentication.

    .PARAMETER CertificateThumbprint
        Certificate thumbprint for app-based auth (Windows only).

    .PARAMETER Certificate
        X509Certificate2 object for app-based auth (cross-platform).

    .PARAMETER Organization
        Organization domain (e.g., contoso.onmicrosoft.com) for app-based auth.

    .PARAMETER LogOnly
        Suppress console output (log file only). Used by keepalive to avoid flashing messages on the dashboard.

    .OUTPUTS
        Boolean indicating connection success.
    #>
    [CmdletBinding(DefaultParameterSetName = 'Interactive')]
    param(
        [Parameter(ParameterSetName = 'Interactive')]
        [string]$UserPrincipalName,

        [Parameter(ParameterSetName = 'Certificate', Mandatory)]
        [string]$AppId,

        [Parameter(ParameterSetName = 'Certificate')]
        [string]$CertificateThumbprint,

        [Parameter(ParameterSetName = 'Certificate')]
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate,

        [Parameter(ParameterSetName = 'Certificate', Mandatory)]
        [string]$Organization,

        [Parameter(DontShow)]
        [switch]$LogOnly
    )

    try {
        Import-Module ExchangeOnlineManagement -ErrorAction Stop

        # Check for existing connection
        $existingConnection = Get-ConnectionInformation -ErrorAction SilentlyContinue |
            Where-Object { $_.ConnectionUri -like "*compliance*" }

        if ($existingConnection) {
            Write-ExportLog -Message "Already connected to Security & Compliance PowerShell" -Level Info -LogOnly:$LogOnly
            return $true
        }

        Write-ExportLog -Message "Connecting to Security & Compliance PowerShell..." -Level Info -LogOnly:$LogOnly

        $connectParams = @{
            ErrorAction = 'Stop'
        }

        if ($PSCmdlet.ParameterSetName -eq 'Certificate') {
            # Certificate-based authentication (unattended)
            $connectParams['AppId'] = $AppId
            $connectParams['Organization'] = $Organization

            if ($CertificateThumbprint) {
                $connectParams['CertificateThumbprint'] = $CertificateThumbprint
            }
            elseif ($Certificate) {
                $connectParams['Certificate'] = $Certificate
            }
            else {
                throw "Either CertificateThumbprint or Certificate parameter is required for app-based auth"
            }

            Write-ExportLog -Message "  Using certificate-based authentication" -Level Info -LogOnly:$LogOnly
        }
        else {
            # Interactive authentication
            if ($UserPrincipalName) {
                $connectParams['UserPrincipalName'] = $UserPrincipalName
                Write-ExportLog -Message "  Using interactive auth for: $UserPrincipalName" -Level Info -LogOnly:$LogOnly
            }
            else {
                Write-ExportLog -Message "  Using interactive authentication (browser)" -Level Info -LogOnly:$LogOnly
            }
        }

        Connect-IPPSSession @connectParams

        Write-ExportLog -Message "Connected successfully" -Level Success -LogOnly:$LogOnly
        return $true
    }
    catch {
        Write-ExportLog -Message "Connection failed: $($_.Exception.Message)" -Level Error -LogOnly:$LogOnly
        return $false
    }
}

function Disconnect-Compl8Compliance {
    <#
    .SYNOPSIS
        Disconnects from Security & Compliance PowerShell.
    .PARAMETER LogOnly
        Suppress console output (log file only). Used by keepalive to avoid flashing messages on the dashboard.
    #>
    [CmdletBinding()]
    param(
        [switch]$LogOnly
    )

    try {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue -InformationAction SilentlyContinue
        Write-ExportLog -Message "Disconnected from Security & Compliance PowerShell" -Level Info -LogOnly:$LogOnly
    }
    catch {
        # Disconnect errors are expected when no session exists or session already closed
        Write-Verbose "Disconnect-ExchangeOnline error (non-critical): $($_.Exception.Message)"
    }
}

#endregion

#region Export Directory Path Helpers

function ConvertTo-SafeDirectoryName {
    <#
    .SYNOPSIS
        Converts a string to a safe directory name by replacing invalid characters.
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
    return (Join-Path $ExportDir "Data" "ContentExplorer" $safeTagType $safeTagName)
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

#region Configuration Functions

function Read-JsonConfig {
    <#
    .SYNOPSIS
        Reads and parses a JSON configuration file.
    .PARAMETER Path
        Path to the JSON file.
    .PARAMETER Required
        Throw error if file doesn't exist.
    .OUTPUTS
        PSCustomObject from parsed JSON.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [switch]$Required
    )

    if (-not (Test-Path $Path)) {
        if ($Required) {
            throw "Required configuration file not found: $Path"
        }
        Write-ExportLog -Message "Config file not found: $Path (using defaults)" -Level Warning
        return $null
    }

    try {
        $content = Get-Content -Raw -Path $Path -ErrorAction Stop
        $config = ConvertFrom-Json -InputObject $content -ErrorAction Stop
        return $config
    }
    catch {
        Write-ExportLog -Message "Failed to parse config file: $($_.Exception.Message)" -Level Error
        throw
    }
}

function Get-EnabledItems {
    <#
    .SYNOPSIS
        Returns items from a config object where value is "True".
    .PARAMETER Config
        PSCustomObject with True/False string values.
    .OUTPUTS
        Array of enabled property names.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [PSCustomObject]$Config
    )

    $enabled = @()
    foreach ($prop in $Config.PSObject.Properties) {
        if ($prop.Value -eq "True") {
            $enabled += $prop.Name
        }
    }
    return $enabled
}

#endregion

#region Export Helper Functions

function Protect-CsvFormulaValue {
    <#
    .SYNOPSIS
        Protects CSV cell values against spreadsheet formula injection.
    .DESCRIPTION
        Prefixes values that start with formula trigger characters (=, +, -, @),
        including values with leading whitespace/tab/newline before the trigger.
    #>
    [CmdletBinding()]
    param(
        [AllowNull()]
        [Parameter()]
        $Value
    )

    if ($null -eq $Value) {
        return $null
    }

    if ($Value -isnot [string]) {
        return $Value
    }

    if ($Value -match '^[\s\t\r\n]*[=\+\-@]') {
        return "'" + $Value
    }

    return $Value
}

function ConvertTo-SafeCsvRecord {
    <#
    .SYNOPSIS
        Converts an object to a CSV-safe object by sanitizing string property values.
    #>
    [CmdletBinding()]
    param(
        [AllowNull()]
        [Parameter()]
        $InputObject
    )

    if ($null -eq $InputObject) {
        return $null
    }

    if ($InputObject -is [System.Collections.IDictionary]) {
        $safeDict = [ordered]@{}
        foreach ($key in $InputObject.Keys) {
            $safeDict[[string]$key] = Protect-CsvFormulaValue -Value $InputObject[$key]
        }
        return [PSCustomObject]$safeDict
    }

    if ($InputObject -is [PSCustomObject] -or $InputObject -isnot [string]) {
        $props = $InputObject.PSObject.Properties
        if ($props -and $props.Count -gt 0) {
            $safeObj = [ordered]@{}
            foreach ($prop in $props) {
                if ($prop.MemberType -match 'Property$') {
                    $safeObj[$prop.Name] = Protect-CsvFormulaValue -Value $prop.Value
                }
            }
            return [PSCustomObject]$safeObj
        }
    }

    return Protect-CsvFormulaValue -Value $InputObject
}

function ConvertTo-SerializableObject {
    <#
    .SYNOPSIS
        Recursively converts hashtables and dictionaries to PSCustomObjects for JSON serialization.
    .DESCRIPTION
        PowerShell's ConvertTo-Json cannot serialize hashtables with non-string keys or
        certain dictionary types. This function converts all such types recursively.
    .PARAMETER InputObject
        The object to convert. Can be null.
    #>
    [CmdletBinding()]
    param(
        [AllowNull()]
        [Parameter()]
        $InputObject
    )

    # Handle null early
    if ($null -eq $InputObject) {
        return $null
    }

    # Get the type for checking
    $type = $InputObject.GetType()

    # Handle hashtables and ordered dictionaries
    if ($InputObject -is [System.Collections.Hashtable] -or
        $InputObject -is [System.Collections.Specialized.OrderedDictionary]) {
        $result = [ordered]@{}
        foreach ($key in $InputObject.Keys) {
            $value = $InputObject[$key]
            if ($null -eq $value) {
                $result[[string]$key] = $null
            }
            else {
                $result[[string]$key] = ConvertTo-SerializableObject -InputObject $value
            }
        }
        return [PSCustomObject]$result
    }

    # Handle generic dictionaries (Dictionary<TKey, TValue>)
    if ($type.IsGenericType -and $type.GetGenericTypeDefinition().Name -like "Dictionary*") {
        $result = [ordered]@{}
        foreach ($key in $InputObject.Keys) {
            $value = $InputObject[$key]
            if ($null -eq $value) {
                $result[[string]$key] = $null
            }
            else {
                $result[[string]$key] = ConvertTo-SerializableObject -InputObject $value
            }
        }
        return [PSCustomObject]$result
    }

    # Handle IDictionary interface
    if ($InputObject -is [System.Collections.IDictionary]) {
        $result = [ordered]@{}
        foreach ($key in $InputObject.Keys) {
            $value = $InputObject[$key]
            if ($null -eq $value) {
                $result[[string]$key] = $null
            }
            else {
                $result[[string]$key] = ConvertTo-SerializableObject -InputObject $value
            }
        }
        return [PSCustomObject]$result
    }

    # Handle arrays
    if ($InputObject -is [System.Array]) {
        $result = [System.Collections.ArrayList]::new()
        foreach ($item in $InputObject) {
            if ($null -eq $item) {
                [void]$result.Add($null)
            }
            else {
                [void]$result.Add((ConvertTo-SerializableObject -InputObject $item))
            }
        }
        return @($result)
    }

    # Handle ArrayList and other IList collections
    if ($InputObject -is [System.Collections.IList]) {
        $result = [System.Collections.ArrayList]::new()
        foreach ($item in $InputObject) {
            if ($null -eq $item) {
                [void]$result.Add($null)
            }
            else {
                [void]$result.Add((ConvertTo-SerializableObject -InputObject $item))
            }
        }
        return @($result)
    }

    # Handle PSCustomObject - need to recurse into properties
    if ($InputObject -is [System.Management.Automation.PSCustomObject]) {
        $result = [ordered]@{}
        foreach ($prop in $InputObject.PSObject.Properties) {
            if ($null -eq $prop.Value) {
                $result[$prop.Name] = $null
            }
            else {
                $result[$prop.Name] = ConvertTo-SerializableObject -InputObject $prop.Value
            }
        }
        return [PSCustomObject]$result
    }

    # Return primitives and other objects as-is
    return $InputObject
}

function Export-ToJsonFile {
    <#
    .SYNOPSIS
        Exports data to a JSON file.
    .PARAMETER Data
        Data to export.
    .PARAMETER Path
        Output file path.
    .PARAMETER Depth
        JSON serialization depth (default: 10).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Data,

        [Parameter(Mandatory)]
        [string]$Path,

        [int]$Depth = 20
    )

    try {
        # Convert hashtables to PSCustomObjects for proper JSON serialization
        $serializableData = ConvertTo-SerializableObject -InputObject $Data
        $json = $serializableData | ConvertTo-Json -Depth $Depth
        Set-Content -Path $Path -Value $json -Encoding UTF8
        Write-ExportLog -Message "Exported to: $Path" -Level Info
        return $true
    }
    catch {
        Write-ExportLog -Message "Failed to export JSON: $($_.Exception.Message)" -Level Error
        return $false
    }
}

function Export-ToCsvFile {
    <#
    .SYNOPSIS
        Exports data to a CSV file.
    .PARAMETER Data
        Data to export.
    .PARAMETER Path
        Output file path.
    .PARAMETER Append
        Append to existing file.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Data,

        [Parameter(Mandatory)]
        [string]$Path,

        [switch]$Append
    )

    try {
        $safeRows = @($Data | ForEach-Object { ConvertTo-SafeCsvRecord -InputObject $_ })
        $params = @{
            Path = $Path
            NoTypeInformation = $true
            Encoding = 'UTF8'
        }

        if ($Append -and (Test-Path $Path)) {
            $params['Append'] = $true
        }

        $safeRows | Export-Csv @params
        Write-ExportLog -Message "Exported to: $Path ($($safeRows.Count) rows)" -Level Info
        return $true
    }
    catch {
        Write-ExportLog -Message "Failed to export CSV: $($_.Exception.Message)" -Level Error
        return $false
    }
}

#endregion

#region Content Explorer Functions

function Export-ContentExplorer {
    <#
    .SYNOPSIS
        Exports Content Explorer data with automatic pagination.

    .DESCRIPTION
        Exports data classification file details from Microsoft Purview Content Explorer.
        Handles pagination automatically for large datasets.

        Reference: https://learn.microsoft.com/powershell/module/exchange/export-contentexplorerdata

    .PARAMETER TagType
        Type of classifier. Valid values: Retention, SensitiveInformationType, Sensitivity, TrainableClassifier

    .PARAMETER TagName
        Name of the label/classifier to export.

    .PARAMETER Workload
        Location to export from. Valid values: Exchange, EXO, OneDrive, ODB, SharePoint, SPO, Teams

    .PARAMETER PageSize
        Number of records per page (1-10000). Default: 5000

    .PARAMETER ConfidenceLevel
        Filter by confidence level: low, medium, high

    .PARAMETER Aggregate
        Return folder-level aggregates instead of item details.

    .OUTPUTS
        Array of exported records.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet("Retention", "SensitiveInformationType", "Sensitivity", "TrainableClassifier")]
        [string]$TagType,

        [Parameter(Mandatory)]
        [string]$TagName,

        [ValidateSet("Exchange", "EXO", "OneDrive", "ODB", "SharePoint", "SPO", "Teams")]
        [string]$Workload,

        [ValidateRange(1, 10000)]
        [int]$PageSize = 5000,

        [ValidateSet("low", "medium", "high")]
        [string]$ConfidenceLevel,

        [switch]$Aggregate
    )

    $allRecords = [System.Collections.ArrayList]::new()
    $pageNumber = 0
    $totalExported = 0

    # Build parameters
    $exportParams = @{
        TagType = $TagType
        TagName = $TagName
        PageSize = $PageSize
        ErrorAction = 'Stop'
    }

    if ($Workload) { $exportParams['Workload'] = $Workload }
    if ($ConfidenceLevel) { $exportParams['ConfidenceLevel'] = $ConfidenceLevel }
    if ($Aggregate) { $exportParams['Aggregate'] = $true }

    $workloadDisplay = if ($Workload) { $Workload } else { "All" }
    Write-ExportLog -Message "Exporting Content Explorer: $TagType/$TagName (Workload: $workloadDisplay)" -Level Info

    try {
        # Initial query
        $result = Export-ContentExplorerData @exportParams

        # Check if we got results (result[0] contains metadata)
        if ($null -eq $result -or $result.Count -eq 0) {
            Write-ExportLog -Message "  No results returned" -Level Warning
            return @()
        }

        $totalCount = $result[0].TotalCount
        if ($totalCount -eq 0) {
            Write-ExportLog -Message "  No matching records found" -Level Warning
            return @()
        }

        Write-ExportLog -Message "  Total matches: $totalCount" -Level Info

        # Process pages
        do {
            $pageNumber++
            $recordsInPage = $result[0].RecordsReturned

            # Records are in array items 1 onwards (item 0 is metadata)
            if ($recordsInPage -gt 0) {
                $pageRecords = $result[1..$recordsInPage]
                [void]$allRecords.AddRange(@($pageRecords))
                $totalExported += $recordsInPage
            }

            Write-ExportLog -Message "  Page $pageNumber`: $recordsInPage records (Total: $totalExported)" -Level Info

            # Check for more pages
            if ($result[0].MorePagesAvailable -eq $true -or $result[0].MorePagesAvailable -eq "True") {
                $exportParams['PageCookie'] = $result[0].PageCookie
                $result = Export-ContentExplorerData @exportParams
            }
            else {
                break
            }
        } while ($true)

        Add-ExportCount -Category "ContentExplorer_$TagType" -Count $totalExported
        Write-ExportLog -Message "  Completed: $totalExported records exported" -Level Success

        return $allRecords
    }
    catch {
        Write-ExportLog -Message "  Export failed: $($_.Exception.Message)" -Level Error
        return @()
    }
}

#endregion

#region Activity Explorer Functions

function Export-ActivityExplorer {
    <#
    .SYNOPSIS
        Exports Activity Explorer data with automatic pagination.

    .DESCRIPTION
        Exports activity data from Microsoft Purview Activity Explorer.
        Handles pagination automatically for large datasets.

        Reference: https://learn.microsoft.com/powershell/module/exchange/export-activityexplorerdata

    .PARAMETER StartTime
        Start of the date range to export.

    .PARAMETER EndTime
        End of the date range to export.

    .PARAMETER OutputFormat
        Output format: Json or Csv. Default: Json

    .PARAMETER PageSize
        Records per page (1-5000). Default: 5000

    .PARAMETER Filters
        Hashtable of filters. Keys are filter names, values are arrays of values.
        Example: @{ Activity = @("DLPRuleMatch"); Workload = @("Exchange", "SharePoint") }

    .OUTPUTS
        Array of activity records.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [datetime]$StartTime,

        [Parameter(Mandatory)]
        [datetime]$EndTime,

        [ValidateSet("Json", "Csv")]
        [string]$OutputFormat = "Json",

        [ValidateRange(1, 5000)]
        [int]$PageSize = 5000,

        [hashtable]$Filters
    )

    $allRecords = [System.Collections.ArrayList]::new()
    $pageNumber = 0
    $totalExported = 0

    # Format dates for the cmdlet
    $startStr = $StartTime.ToString("MM/dd/yyyy HH:mm:ss")
    $endStr = $EndTime.ToString("MM/dd/yyyy HH:mm:ss")

    # Build parameters
    $exportParams = @{
        StartTime = $startStr
        EndTime = $endStr
        OutputFormat = $OutputFormat
        PageSize = $PageSize
        ErrorAction = 'Stop'
    }

    # Add filters (up to 5 supported by the API)
    if ($Filters -and $Filters.Count -gt 0) {
        if ($Filters.Count -gt 5) {
            Write-Warning "Activity Explorer API supports max 5 filters per category. $($Filters.Count) items enabled — only the first 5 will be used. Disable some items in the config file."
        }
        $filterIndex = 1
        foreach ($filterName in $Filters.Keys) {
            if ($filterIndex -le 5) {
                $filterValues = @($filterName) + @($Filters[$filterName])
                $exportParams["Filter$filterIndex"] = $filterValues
                $filterIndex++
            }
        }
    }

    Write-ExportLog -Message "Exporting Activity Explorer: $startStr to $endStr" -Level Info
    if ($Filters) {
        Write-ExportLog -Message "  Filters: $($Filters.Keys -join ', ')" -Level Info
    }

    try {
        # Initial query
        Write-ExportLog -Message "  Calling Export-ActivityExplorerData..." -Level Info
        $result = Export-ActivityExplorerData @exportParams

        if ($null -eq $result) {
            Write-ExportLog -Message "  Result is null - no data returned from cmdlet" -Level Warning
            return @()
        }

        # Log result object properties for debugging
        Write-ExportLog -Message "  Result type: $($result.GetType().Name)" -Level Info
        Write-ExportLog -Message "  Result properties: $($result.PSObject.Properties.Name -join ', ')" -Level Info

        $totalCount = $result.TotalResultCount
        Write-ExportLog -Message "  TotalResultCount: $totalCount" -Level Info

        if ($null -eq $totalCount -or $totalCount -eq 0 -or $totalCount -eq "0") {
            Write-ExportLog -Message "  No matching activities found (TotalResultCount is 0 or null)" -Level Warning
            return @()
        }

        Write-ExportLog -Message "  Total activities to retrieve: $totalCount" -Level Info

        # Process pages
        do {
            $pageNumber++

            # Parse ResultData (JSON string)
            if ($result.ResultData) {
                Write-ExportLog -Message "  ResultData length: $($result.ResultData.Length) chars" -Level Info
                $pageRecords = $result.ResultData | ConvertFrom-Json
                if ($pageRecords) {
                    $recordCount = @($pageRecords).Count
                    [void]$allRecords.AddRange(@($pageRecords))
                    $totalExported += $recordCount
                    Write-ExportLog -Message "  Page $pageNumber`: $recordCount records (Total: $totalExported)" -Level Info
                }
            }
            else {
                Write-ExportLog -Message "  Page $pageNumber`: ResultData is empty or null" -Level Warning
            }

            # Check for more pages (LastPage property)
            Write-ExportLog -Message "  LastPage: $($result.LastPage), WaterMark: $($result.WaterMark)" -Level Info
            if ($result.LastPage -eq $false) {
                # PageCookie valid for 120 seconds per documentation
                $exportParams['PageCookie'] = $result.WaterMark
                $result = Export-ActivityExplorerData @exportParams
            }
            else {
                break
            }
        } while ($true)

        Add-ExportCount -Category "ActivityExplorer" -Count $totalExported
        Write-ExportLog -Message "  Completed: $totalExported records exported" -Level Success

        return $allRecords
    }
    catch {
        Write-ExportLog -Message "  Export failed: $($_.Exception.Message)" -Level Error
        return @()
    }
}

#endregion

#region DLP and Policy Functions

function Export-DlpPolicies {
    <#
    .SYNOPSIS
        Exports DLP compliance policies and rules.

    .DESCRIPTION
        Retrieves all DLP policies and their associated rules from Microsoft Purview.

    .OUTPUTS
        Hashtable with Policies and Rules arrays.
    #>
    [CmdletBinding()]
    param()

    $result = @{
        Policies = @()
        Rules = @()
    }

    Write-ExportLog -Message "Exporting DLP Policies..." -Level Info

    try {
        # Get DLP Policies
        Write-ExportLog -Message "  Retrieving DLP policies..." -Level Info
        $policies = Get-DlpCompliancePolicy -ErrorAction Stop
        $result.Policies = $policies
        $policyCount = @($policies).Count
        Add-ExportCount -Category "DlpPolicies" -Count $policyCount
        Write-ExportLog -Message "  Found $policyCount DLP policies" -Level Success

        # Get DLP Rules
        Write-ExportLog -Message "  Retrieving DLP rules..." -Level Info
        $rules = Get-DlpComplianceRule -ErrorAction Stop
        $result.Rules = $rules
        $ruleCount = @($rules).Count
        Add-ExportCount -Category "DlpRules" -Count $ruleCount
        Write-ExportLog -Message "  Found $ruleCount DLP rules" -Level Success
    }
    catch {
        Write-ExportLog -Message "  Failed: $($_.Exception.Message)" -Level Error
    }

    return $result
}

function Export-SensitiveInfoTypes {
    <#
    .SYNOPSIS
        Exports Sensitive Information Type definitions.

    .OUTPUTS
        Array of SIT definitions.
    #>
    [CmdletBinding()]
    param()

    Write-ExportLog -Message "Exporting Sensitive Information Types..." -Level Info

    try {
        Write-ExportLog -Message "  Retrieving sensitive information types..." -Level Info
        $sits = Get-DlpSensitiveInformationType -ErrorAction Stop
        $sitCount = @($sits).Count
        Add-ExportCount -Category "SensitiveInfoTypes" -Count $sitCount
        Write-ExportLog -Message "  Found $sitCount sensitive information types" -Level Success
        return $sits
    }
    catch {
        Write-ExportLog -Message "  Failed: $($_.Exception.Message)" -Level Error
        return @()
    }
}

#endregion

#region Label Functions

function Export-SensitivityLabels {
    <#
    .SYNOPSIS
        Exports sensitivity labels and label policies.

    .OUTPUTS
        Hashtable with Labels and Policies arrays.
    #>
    [CmdletBinding()]
    param()

    $result = @{
        Labels = @()
        Policies = @()
    }

    Write-ExportLog -Message "Exporting Sensitivity Labels..." -Level Info

    try {
        # Get Labels
        Write-ExportLog -Message "  Retrieving sensitivity labels..." -Level Info
        $labels = Get-Label -ErrorAction Stop
        $result.Labels = $labels
        $labelCount = @($labels).Count
        Add-ExportCount -Category "SensitivityLabels" -Count $labelCount
        Write-ExportLog -Message "  Found $labelCount sensitivity labels" -Level Success

        # Get Label Policies
        Write-ExportLog -Message "  Retrieving label policies..." -Level Info
        $policies = Get-LabelPolicy -ErrorAction Stop
        $result.Policies = $policies
        $policyCount = @($policies).Count
        Add-ExportCount -Category "LabelPolicies" -Count $policyCount
        Write-ExportLog -Message "  Found $policyCount label policies" -Level Success
    }
    catch {
        Write-ExportLog -Message "  Failed: $($_.Exception.Message)" -Level Error
    }

    return $result
}

function Export-RetentionLabels {
    <#
    .SYNOPSIS
        Exports retention labels and policies.

    .OUTPUTS
        Hashtable with Labels, Policies, and Rules arrays.
    #>
    [CmdletBinding()]
    param()

    $result = @{
        Labels = @()
        Policies = @()
        Rules = @()
    }

    Write-ExportLog -Message "Exporting Retention Labels..." -Level Info

    try {
        # Get Retention Labels (Compliance Tags)
        Write-ExportLog -Message "  Retrieving retention labels (compliance tags)..." -Level Info
        $labels = Get-ComplianceTag -ErrorAction Stop
        $result.Labels = $labels
        $labelCount = @($labels).Count
        Add-ExportCount -Category "RetentionLabels" -Count $labelCount
        Write-ExportLog -Message "  Found $labelCount retention labels" -Level Success

        # Get Retention Policies
        Write-ExportLog -Message "  Retrieving retention policies..." -Level Info
        $policies = Get-RetentionCompliancePolicy -ErrorAction Stop
        $result.Policies = $policies
        $policyCount = @($policies).Count
        Add-ExportCount -Category "RetentionPolicies" -Count $policyCount
        Write-ExportLog -Message "  Found $policyCount retention policies" -Level Success

        # Get Retention Rules
        Write-ExportLog -Message "  Retrieving retention rules..." -Level Info
        $rules = Get-RetentionComplianceRule -ErrorAction Stop
        $result.Rules = $rules
        $ruleCount = @($rules).Count
        Add-ExportCount -Category "RetentionRules" -Count $ruleCount
        Write-ExportLog -Message "  Found $ruleCount retention rules" -Level Success
    }
    catch {
        Write-ExportLog -Message "  Failed: $($_.Exception.Message)" -Level Error
    }

    return $result
}

#endregion

#region eDiscovery Functions

function Export-eDiscoveryCases {
    <#
    .SYNOPSIS
        Exports eDiscovery cases and searches.

    .OUTPUTS
        Hashtable with Cases and Searches arrays.
    #>
    [CmdletBinding()]
    param()

    $result = @{
        Cases = @()
        Searches = @()
    }

    Write-ExportLog -Message "Exporting eDiscovery Cases..." -Level Info

    try {
        # Get Compliance Cases
        Write-ExportLog -Message "  Retrieving compliance cases..." -Level Info
        $cases = Get-ComplianceCase -ErrorAction Stop
        $result.Cases = $cases
        $caseCount = @($cases).Count
        Add-ExportCount -Category "eDiscoveryCases" -Count $caseCount
        Write-ExportLog -Message "  Found $caseCount compliance cases" -Level Success

        # Get Compliance Searches
        Write-ExportLog -Message "  Retrieving compliance searches..." -Level Info
        $searches = Get-ComplianceSearch -ErrorAction Stop
        $result.Searches = $searches
        $searchCount = @($searches).Count
        Add-ExportCount -Category "ComplianceSearches" -Count $searchCount
        Write-ExportLog -Message "  Found $searchCount compliance searches" -Level Success
    }
    catch {
        Write-ExportLog -Message "  Failed: $($_.Exception.Message)" -Level Error
    }

    return $result
}

#endregion

#region RBAC Functions

function Export-RbacConfiguration {
    <#
    .SYNOPSIS
        Exports RBAC role groups and assignments.

    .OUTPUTS
        Hashtable with RoleGroups and Members arrays.
    #>
    [CmdletBinding()]
    param()

    $result = @{
        RoleGroups = @()
        Members = @()
    }

    Write-ExportLog -Message "Exporting RBAC Configuration..." -Level Info

    try {
        # Get Role Groups
        Write-ExportLog -Message "  Retrieving role groups..." -Level Info
        $roleGroups = Get-RoleGroup -ErrorAction Stop
        $result.RoleGroups = $roleGroups
        $totalGroups = @($roleGroups).Count
        Add-ExportCount -Category "RoleGroups" -Count $totalGroups
        Write-ExportLog -Message "  Found $totalGroups role groups" -Level Success

        # Get Members for each role group with progress
        Write-ExportLog -Message "  Retrieving members for each role group..." -Level Info
        $allMembers = @()
        $currentGroup = 0

        foreach ($rg in $roleGroups) {
            $currentGroup++
            $percentComplete = [math]::Round(($currentGroup / $totalGroups) * 100, 0)
            Write-ExportLog -Message "    [$currentGroup/$totalGroups] ($percentComplete%) Processing: $($rg.Name)" -Level Info

            try {
                $members = Get-RoleGroupMember -Identity $rg.Name -ErrorAction SilentlyContinue
                if ($members) {
                    $memberCount = @($members).Count
                    foreach ($member in $members) {
                        $allMembers += [PSCustomObject]@{
                            RoleGroup = $rg.Name
                            MemberName = $member.Name
                            MemberType = $member.RecipientType
                        }
                    }
                    if ($memberCount -gt 0) {
                        Write-ExportLog -Message "      -> $memberCount members" -Level Info
                    }
                }
            }
            catch {
                Write-ExportLog -Message "      -> Failed to read members: $($_.Exception.Message)" -Level Warning
            }
        }
        $result.Members = $allMembers
        Add-ExportCount -Category "RoleGroupMembers" -Count $allMembers.Count
        Write-ExportLog -Message "  Completed: $($allMembers.Count) total role group members" -Level Success
    }
    catch {
        Write-ExportLog -Message "  Failed: $($_.Exception.Message)" -Level Error
    }

    return $result
}

#endregion

#region Configuration Validation

function Test-ExportConfiguration {
    <#
    .SYNOPSIS
        Validates that required configuration files exist for the chosen export mode.

    .DESCRIPTION
        Checks for the presence of JSON configuration files needed by the selected
        export mode before establishing a connection to Security & Compliance PowerShell.
        This prevents wasted authentication time when config files are missing.

        For FullExport mode, Content Explorer and Activity Explorer config files are
        checked unless the corresponding -NoContent or -NoActivity switches are set.

    .PARAMETER ExportMode
        The export mode to validate. Valid values: Full, DLP, Labels, ContentExplorer,
        ActivityExplorer, eDiscovery, RBAC.

    .PARAMETER ScriptRoot
        Root directory of the export script, used to locate ConfigFiles subfolder.

    .PARAMETER NoContent
        When set, skips validation of Content Explorer config files (used with FullExport).

    .PARAMETER NoActivity
        When set, skips validation of Activity Explorer config files (used with FullExport).

    .OUTPUTS
        Boolean. Returns $true if all required config files exist, $false otherwise.

    .EXAMPLE
        Test-ExportConfiguration -ExportMode "ContentExplorer" -ScriptRoot $PSScriptRoot

    .EXAMPLE
        Test-ExportConfiguration -ExportMode "Full" -ScriptRoot $PSScriptRoot -NoActivity
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet("Full", "DLP", "Labels", "ContentExplorer", "ActivityExplorer", "eDiscovery", "RBAC")]
        [string]$ExportMode,

        [Parameter(Mandatory)]
        [string]$ScriptRoot,

        [switch]$NoContent,

        [switch]$NoActivity
    )

    $configDir = Join-Path $ScriptRoot "ConfigFiles"
    $allValid = $true

    # Build the list of required config files based on export mode
    $requiredConfigs = @()

    switch ($ExportMode) {
        "ContentExplorer" {
            $requiredConfigs += @{
                Name = "ContentExplorerClassifiers.json"
                Description = "Content Explorer classifier configuration"
            }
        }
        "ActivityExplorer" {
            $requiredConfigs += @{
                Name = "ActivityExplorerSelector.json"
                Description = "Activity Explorer activity/workload filter"
            }
        }
        "Full" {
            if (-not $NoContent) {
                $requiredConfigs += @{
                    Name = "ContentExplorerClassifiers.json"
                    Description = "Content Explorer classifier configuration"
                }
            }
            if (-not $NoActivity) {
                $requiredConfigs += @{
                    Name = "ActivityExplorerSelector.json"
                    Description = "Activity Explorer activity/workload filter"
                }
            }
        }
        # DLP, Labels, eDiscovery, RBAC modes do not require additional config files
    }

    # Validate each required config file
    foreach ($config in $requiredConfigs) {
        $configPath = Join-Path $configDir $config.Name
        if (Test-Path $configPath) {
            Write-ExportLog -Message "  Config OK: $($config.Name)" -Level Success
        }
        else {
            Write-ExportLog -Message "  Config MISSING: $($config.Name) ($($config.Description))" -Level Error
            Write-ExportLog -Message "    Expected at: $configPath" -Level Info
            $allValid = $false
        }
    }

    if ($allValid) {
        Write-ExportLog -Message "Configuration validation passed" -Level Success
    }
    else {
        Write-ExportLog -Message "Configuration validation failed - missing required files" -Level Error
    }

    return $allValid
}

#endregion

#region Error Handling Functions

function Write-ExportErrorLog {
    <#
    .SYNOPSIS
        Writes detailed error information to an error log file using atomic NTFS append.

    .DESCRIPTION
        Appends a structured error entry to the specified log file. Each entry includes
        a timestamp, operation context, full exception details, HTTP error classification,
        inner exception chain, and optional additional context data.

        Uses Add-Content for atomic NTFS append, making it safe for multiple terminals
        to write to the same log file concurrently.

    .PARAMETER ErrorLogPath
        Full path to the error log file. Created if it does not exist.

    .PARAMETER Context
        Description of the operation that failed (e.g., "Worker Aggregate", "Detail Export").

    .PARAMETER TaskKey
        Identifier for the task that failed (e.g., "SensitiveInformationType|Credit Card|Exchange").

    .PARAMETER ErrorRecord
        The PowerShell ErrorRecord object from the catch block.

    .PARAMETER AdditionalData
        Optional hashtable of extra context to include in the log entry
        (e.g., retry count, page number, page cookie).

    .EXAMPLE
        Write-ExportErrorLog -ErrorLogPath $logPath -Context "Worker Aggregate" `
            -TaskKey "SIT|Credit Card|Exchange" -ErrorRecord $_

    .EXAMPLE
        Write-ExportErrorLog -ErrorLogPath $logPath -Context "Detail Export" `
            -TaskKey "Sensitivity|Confidential|SharePoint" -ErrorRecord $_ `
            -AdditionalData @{ RetryCount = 3; PageNumber = 12 }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ErrorLogPath,

        [string]$Context = "Unknown",

        [string]$TaskKey,

        [Parameter(Mandatory)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord,

        [hashtable]$AdditionalData
    )

    try {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
        $separator = "=" * 80

        # Build the error entry using a StringBuilder for efficiency
        $sb = [System.Text.StringBuilder]::new(2048)
        [void]$sb.AppendLine($separator)
        [void]$sb.AppendLine("TIMESTAMP:  $timestamp")
        [void]$sb.AppendLine("CONTEXT:    $Context")

        if ($TaskKey) {
            [void]$sb.AppendLine("TASK:       $TaskKey")
        }

        [void]$sb.AppendLine("PID:        $PID")

        # Exception details
        $exception = $ErrorRecord.Exception
        if ($exception) {
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("EXCEPTION:")
            [void]$sb.AppendLine("  Message:  $($exception.Message)")
            [void]$sb.AppendLine("  Type:     $($exception.GetType().FullName)")
        }

        # Stack trace
        if ($ErrorRecord.ScriptStackTrace) {
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("STACK TRACE:")
            [void]$sb.AppendLine($ErrorRecord.ScriptStackTrace)
        }

        # HTTP error classification
        $httpInfo = Get-HttpErrorExplanation -ErrorMessage $exception.Message -ErrorRecord $ErrorRecord
        if ($httpInfo) {
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("HTTP ERROR ANALYSIS:")
            if ($httpInfo.StatusCode) {
                [void]$sb.AppendLine("  Status Code: $($httpInfo.StatusCode)")
            }
            [void]$sb.AppendLine("  Category:    $($httpInfo.Category)")
            [void]$sb.AppendLine("  IsTransient: $($httpInfo.IsTransient)")
            [void]$sb.AppendLine("  Explanation: $($httpInfo.Explanation)")
        }

        # Inner exception chain (up to 3 levels)
        $innerException = $exception.InnerException
        $innerLevel = 1
        while ($innerException -and $innerLevel -le 3) {
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("INNER EXCEPTION (Level $innerLevel):")
            [void]$sb.AppendLine("  Message:  $($innerException.Message)")
            [void]$sb.AppendLine("  Type:     $($innerException.GetType().FullName)")
            $innerException = $innerException.InnerException
            $innerLevel++
        }

        # Additional context data
        if ($AdditionalData -and $AdditionalData.Count -gt 0) {
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("ADDITIONAL DATA:")
            foreach ($key in $AdditionalData.Keys) {
                $value = $AdditionalData[$key]
                # Truncate long values to keep log readable
                $valueStr = if ($null -eq $value) { "(null)" } else { "$value" }
                if ($valueStr.Length -gt 500) {
                    $valueStr = $valueStr.Substring(0, 500) + "... (truncated)"
                }
                [void]$sb.AppendLine("  ${key}: $valueStr")
            }
        }

        [void]$sb.AppendLine("")

        # Ensure parent directory exists
        $logDir = Split-Path $ErrorLogPath -Parent
        if ($logDir -and -not (Test-Path $logDir)) {
            New-Item -ItemType Directory -Force -Path $logDir | Out-Null
        }

        # Atomic NTFS append - safe for concurrent writes from multiple terminals
        Add-Content -Path $ErrorLogPath -Value $sb.ToString() -Encoding UTF8
    }
    catch {
        # If error logging itself fails, write to console as last resort
        Write-Warning "Failed to write to error log '$ErrorLogPath': $($_.Exception.Message)"
    }
}

function Format-ErrorDetail {
    <#
    .SYNOPSIS
        Formats a PowerShell ErrorRecord into a human-readable string.

    .DESCRIPTION
        Extracts the exception message, exception type, and script stack trace from
        an ErrorRecord and formats them into a single readable string suitable for
        log messages or console output.

    .PARAMETER ErrorRecord
        The PowerShell ErrorRecord object to format.

    .OUTPUTS
        String containing the formatted error details.

    .EXAMPLE
        try { ... } catch { $detail = Format-ErrorDetail -ErrorRecord $_; Write-Host $detail }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )

    $parts = [System.Collections.ArrayList]::new()

    # Exception message
    if ($ErrorRecord.Exception) {
        [void]$parts.Add("Message: $($ErrorRecord.Exception.Message)")
        [void]$parts.Add("Type: $($ErrorRecord.Exception.GetType().FullName)")
    }
    else {
        [void]$parts.Add("Message: $($ErrorRecord.ToString())")
    }

    # Stack trace
    if ($ErrorRecord.ScriptStackTrace) {
        [void]$parts.Add("Stack Trace:")
        [void]$parts.Add($ErrorRecord.ScriptStackTrace)
    }

    return $parts -join [Environment]::NewLine
}

function Get-HttpErrorExplanation {
    <#
    .SYNOPSIS
        Classifies HTTP errors from Purview API responses.

    .DESCRIPTION
        Analyzes an error message and/or ErrorRecord to determine the HTTP status code,
        error category, whether the error is transient (retryable), and a human-readable
        explanation. This is used to drive retry logic and error reporting.

        Categories:
        - AuthError:   HTTP 401 (token expired or invalid credentials)
        - Throttle:    HTTP 429 (rate limiting / too many requests)
        - ServerError: HTTP 500, 502, 503, 504 (transient server-side failures)
        - ClientError: HTTP 400, 403, 404 (non-retryable client errors)
        - Network:     WebException, timeout, connection reset (transient)
        - Unknown:     Unrecognized error pattern

    .PARAMETER ErrorMessage
        The exception message string to analyze for HTTP status codes.

    .PARAMETER ErrorRecord
        The full PowerShell ErrorRecord for additional context extraction.

    .OUTPUTS
        Hashtable with keys: StatusCode (int or $null), Category (string),
        IsTransient (bool), Explanation (string).

    .EXAMPLE
        $info = Get-HttpErrorExplanation -ErrorMessage $_.Exception.Message -ErrorRecord $_
        if ($info.IsTransient) { Start-Sleep -Seconds 60; continue }
    #>
    [CmdletBinding()]
    param(
        [string]$ErrorMessage,

        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )

    $result = @{
        StatusCode  = $null
        Category    = "Unknown"
        IsTransient = $false
        Explanation = "Unrecognized error"
    }

    # Combine error message sources for pattern matching
    $fullMessage = $ErrorMessage
    if ($ErrorRecord -and $ErrorRecord.Exception) {
        $fullMessage = "$ErrorMessage $($ErrorRecord.Exception.Message)"
        # Check inner exception messages too
        $inner = $ErrorRecord.Exception.InnerException
        if ($inner) {
            $fullMessage = "$fullMessage $($inner.Message)"
        }
    }

    if ([string]::IsNullOrWhiteSpace($fullMessage)) {
        return $result
    }

    # Try to extract HTTP status code from the message
    # Common patterns: "(401)", "Status code: 401", "HttpStatusCode: 401", "Response status code does not indicate success: 401"
    $statusCode = $null
    if ($fullMessage -match '\b(4\d{2}|5\d{2})\b') {
        $statusCode = [int]$Matches[1]
    }

    # Classify by HTTP status code first
    if ($statusCode) {
        $result.StatusCode = $statusCode

        switch ($statusCode) {
            401 {
                $result.Category = "AuthError"
                $result.IsTransient = $false
                $result.Explanation = "Authentication expired or invalid. Token refresh or re-authentication required."
            }
            403 {
                $result.Category = "ClientError"
                $result.IsTransient = $false
                $result.Explanation = "Access denied. Insufficient permissions for the requested operation."
            }
            404 {
                $result.Category = "ClientError"
                $result.IsTransient = $false
                $result.Explanation = "Resource not found. The requested item or endpoint does not exist."
            }
            429 {
                $result.Category = "Throttle"
                $result.IsTransient = $true
                $result.Explanation = "Request throttled. Too many requests - wait before retrying."
            }
            400 {
                $result.Category = "ClientError"
                $result.IsTransient = $false
                $result.Explanation = "Bad request. Check parameters and date ranges."
            }
            500 {
                $result.Category = "ServerError"
                $result.IsTransient = $true
                $result.Explanation = "Internal server error. Transient failure on the server side."
            }
            502 {
                $result.Category = "ServerError"
                $result.IsTransient = $true
                $result.Explanation = "Bad gateway. Upstream server returned an invalid response."
            }
            503 {
                $result.Category = "ServerError"
                $result.IsTransient = $true
                $result.Explanation = "Service unavailable. Server is temporarily overloaded or under maintenance."
            }
            504 {
                $result.Category = "ServerError"
                $result.IsTransient = $true
                $result.Explanation = "Gateway timeout. Server did not respond in time."
            }
            default {
                if ($statusCode -ge 500) {
                    $result.Category = "ServerError"
                    $result.IsTransient = $true
                    $result.Explanation = "Server error (HTTP $statusCode). Transient failure."
                }
                elseif ($statusCode -ge 400) {
                    $result.Category = "ClientError"
                    $result.IsTransient = $false
                    $result.Explanation = "Client error (HTTP $statusCode). Review request parameters."
                }
            }
        }

        return $result
    }

    # No status code found - classify by message patterns
    $messageLower = $fullMessage.ToLower()

    # Network / connectivity errors
    if ($messageLower -match 'webexception|web exception|socketexception|socket exception') {
        $result.Category = "Network"
        $result.IsTransient = $true
        $result.Explanation = "Network error (WebException). Check connectivity and retry."
        return $result
    }

    if ($messageLower -match 'timeout|timed\s*out|operation\s*was\s*canceled') {
        $result.Category = "Network"
        $result.IsTransient = $true
        $result.Explanation = "Request timed out. The server did not respond within the allowed time."
        return $result
    }

    if ($messageLower -match 'connection\s*(was\s*)?reset|connection\s*(was\s*)?closed|connection\s*(was\s*)?refused') {
        $result.Category = "Network"
        $result.IsTransient = $true
        $result.Explanation = "Connection was reset or refused. Network disruption detected."
        return $result
    }

    if ($messageLower -match 'ssl|tls|certificate|secure\s*channel') {
        $result.Category = "Network"
        $result.IsTransient = $true
        $result.Explanation = "SSL/TLS error. Secure channel could not be established."
        return $result
    }

    # Authentication patterns without HTTP status codes
    if ($messageLower -match 'unauthorized|authentication\s*failed|token\s*expired|access\s*token') {
        $result.Category = "AuthError"
        $result.IsTransient = $false
        $result.Explanation = "Authentication error detected in message. Re-authentication may be required."
        return $result
    }

    # Throttling patterns without HTTP status codes
    if ($messageLower -match 'throttl|rate\s*limit|too\s*many\s*requests') {
        $result.Category = "Throttle"
        $result.IsTransient = $true
        $result.Explanation = "Throttling detected in message. Wait before retrying."
        return $result
    }

    # Generic transient server-side patterns
    if ($messageLower -match 'service\s*unavailable|server\s*(is\s*)?busy|temporarily\s*unavailable|internal\s*server\s*error') {
        $result.Category = "ServerError"
        $result.IsTransient = $true
        $result.Explanation = "Server-side transient error. Retry after delay."
        return $result
    }

    # Role definition not ready (EXO session initialization race condition)
    if ($messageLower -match 'not present in the role definition') {
        $result.Category = "ServerError"
        $result.IsTransient = $true
        $result.Explanation = "Role definition not yet provisioned. Backend session still initializing - retry after delay."
        return $result
    }

    # AggregateException from EXO REST module (wraps internal connection/HTTP failures)
    if ($messageLower -match 'aggregateexception|one or more errors occurred') {
        $result.Category = "ServerError"
        $result.IsTransient = $true
        $result.Explanation = "AggregateException from EXO module. Likely transient server-side or connection failure."
        return $result
    }

    # Fall through - unknown error
    $result.Explanation = "Unrecognized error: $($ErrorMessage.Substring(0, [Math]::Min($ErrorMessage.Length, 200)))"
    return $result
}

#endregion

#region Auth Recovery

function Invoke-WithAuthRecovery {
    <#
    .SYNOPSIS
        Wraps a scriptblock with automatic authentication token expiry recovery.

    .DESCRIPTION
        Executes the provided scriptblock and monitors for HTTP 401 (authentication expired)
        errors. On detecting a 401:

        - Certificate auth: Silently disconnects and reconnects using the stored auth
          parameters. No user interaction required.
        - Interactive auth (non-worker): Prompts the user before opening a browser login
          window. User can press Enter to re-authenticate or Q to abort.
        - Worker mode: Throws the error immediately to let the worker handle it (workers
          cannot prompt for interactive auth; they save progress and exit).

        On successful re-authentication, the scriptblock is retried once.

    .PARAMETER ScriptBlock
        The scriptblock to execute with auth recovery protection.

    .PARAMETER AuthParams
        Hashtable of authentication parameters (same format as Connect-Compl8Compliance).
        If it contains AppId and CertificateThumbprint, certificate auth is assumed.

    .PARAMETER Context
        Description of the operation for logging purposes.

    .PARAMETER IsWorkerMode
        When $true, indicates this is a spawned worker terminal that cannot prompt
        for interactive authentication. Auth failures will throw immediately.

    .OUTPUTS
        The return value of the scriptblock on success.

    .EXAMPLE
        $result = Invoke-WithAuthRecovery -ScriptBlock {
            Export-ContentExplorerData -TagType "SensitiveInformationType" -TagName "Credit Card" -PageSize 1000
        } -AuthParams $script:AuthParams -Context "Content Explorer export"

    .EXAMPLE
        $data = Invoke-WithAuthRecovery -ScriptBlock { Get-DlpCompliancePolicy } `
            -AuthParams $authParams -Context "DLP Policy retrieval"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [scriptblock]$ScriptBlock,

        [hashtable]$AuthParams,

        [string]$Context = "API call",

        [bool]$IsWorkerMode = $false
    )

    try {
        # Execute the scriptblock
        $result = & $ScriptBlock
        return $result
    }
    catch {
        # Check if this is an auth error
        $errorInfo = Get-HttpErrorExplanation -ErrorMessage $_.Exception.Message -ErrorRecord $_

        if ($errorInfo.Category -ne "AuthError") {
            # Not an auth error - rethrow as-is
            throw
        }

        Write-ExportLog -Message "AUTH EXPIRED during $Context - attempting recovery..." -Level Warning

        # Worker mode cannot do interactive auth - throw to let worker save progress and exit
        if ($IsWorkerMode) {
            $isCertAuth = $AuthParams -and $AuthParams.ContainsKey('AppId') -and
                         ($AuthParams.ContainsKey('CertificateThumbprint') -or $AuthParams.ContainsKey('Certificate'))

            if (-not $isCertAuth) {
                Write-ExportLog -Message "Worker cannot re-authenticate interactively. Throwing to save progress." -Level Error
                throw
            }
        }

        # Determine auth mode
        $isCertAuth = $AuthParams -and $AuthParams.ContainsKey('AppId') -and
                     ($AuthParams.ContainsKey('CertificateThumbprint') -or $AuthParams.ContainsKey('Certificate'))

        if ($isCertAuth) {
            # Certificate auth - reconnect silently
            Write-ExportLog -Message "  Reconnecting with certificate authentication..." -Level Info
            try {
                Disconnect-Compl8Compliance
                $reconnected = Connect-Compl8Compliance @AuthParams
                if (-not $reconnected) {
                    Write-ExportLog -Message "  Certificate re-authentication returned false" -Level Error
                    throw $_
                }
                Write-ExportLog -Message "  Re-authentication successful - retrying $Context" -Level Success
            }
            catch {
                Write-ExportLog -Message "  Certificate re-authentication failed: $($_.Exception.Message)" -Level Error
                throw
            }
        }
        else {
            # Interactive auth - prompt user (non-worker only)
            Write-ExportLog -Message "  Authentication token has expired." -Level Warning
            Write-Host ""
            Write-Host "Press ENTER to open browser for re-authentication, or Q to abort: " -ForegroundColor Yellow -NoNewline
            $userInput = Read-Host

            if ($userInput -and $userInput.Trim().ToUpper() -eq 'Q') {
                Write-ExportLog -Message "  User chose to abort re-authentication" -Level Error
                throw $_
            }

            try {
                Disconnect-Compl8Compliance
                $connectParams = if ($AuthParams -and $AuthParams.Count -gt 0) { $AuthParams } else { @{} }
                $reconnected = Connect-Compl8Compliance @connectParams
                if (-not $reconnected) {
                    Write-ExportLog -Message "  Interactive re-authentication returned false" -Level Error
                    throw $_
                }
                Write-ExportLog -Message "  Re-authentication successful - retrying $Context" -Level Success
            }
            catch {
                Write-ExportLog -Message "  Interactive re-authentication failed: $($_.Exception.Message)" -Level Error
                throw
            }
        }

        # Retry the scriptblock once after successful re-authentication
        try {
            $retryResult = & $ScriptBlock
            return $retryResult
        }
        catch {
            Write-ExportLog -Message "  Retry after re-authentication failed: $($_.Exception.Message)" -Level Error
            throw
        }
    }
}

#endregion

#region Utility Functions

function Format-TimeSpan {
    <#
    .SYNOPSIS
        Formats a duration in seconds into a human-readable time string.

    .DESCRIPTION
        Converts a number of seconds into a concise display string using hours,
        minutes, and seconds components. Omits zero-value leading components
        for brevity (e.g., does not show "0h" for sub-hour durations).

    .PARAMETER Seconds
        The duration in seconds to format. Accepts fractional values.

    .OUTPUTS
        String in format "Xh Ym Zs", "Ym Zs", or "Zs" depending on magnitude.

    .EXAMPLE
        Format-TimeSpan -Seconds 3725.4
        # Returns "1h 2m 5s"

    .EXAMPLE
        Format-TimeSpan -Seconds 330
        # Returns "5m 30s"

    .EXAMPLE
        Format-TimeSpan -Seconds 45.7
        # Returns "45s"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [double]$Seconds
    )

    # Handle negative or zero values
    if ($Seconds -le 0) {
        return "0s"
    }

    $totalSeconds = [int][Math]::Floor($Seconds)
    $hours = [int][Math]::Floor($totalSeconds / 3600)
    $minutes = [int][Math]::Floor(($totalSeconds % 3600) / 60)
    $secs = $totalSeconds % 60

    if ($hours -gt 0) {
        return "${hours}h ${minutes}m ${secs}s"
    }
    elseif ($minutes -gt 0) {
        return "${minutes}m ${secs}s"
    }
    else {
        return "${secs}s"
    }
}

#endregion

#region SIT and Tenant Functions

function Get-SITsToSkip {
    <#
    .SYNOPSIS
        Loads the list of Sensitive Information Types to exclude from Content Explorer exports.

    .DESCRIPTION
        Reads the SITstoSkip.json configuration file and returns the names of all SITs
        marked as "True" (meaning they should be skipped). This is used to filter out
        country-specific or noisy SITs that generate large volumes of irrelevant data.

        Uses Read-JsonConfig and Get-EnabledItems internally.

    .PARAMETER ConfigPath
        Full path to the SITstoSkip.json file. Defaults to ConfigFiles/SITstoSkip.json
        relative to the module's parent directory.

    .OUTPUTS
        String array of SIT names to skip. Returns empty array if config file is missing.

    .EXAMPLE
        $sitsToSkip = Get-SITsToSkip

    .EXAMPLE
        $sitsToSkip = Get-SITsToSkip -ConfigPath "C:\Config\SITstoSkip.json"
    #>
    [CmdletBinding()]
    param(
        [string]$ConfigPath
    )

    # Default path: module parent directory / ConfigFiles / SITstoSkip.json
    if (-not $ConfigPath) {
        $moduleRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
        # If running from the module file directly, PSScriptRoot is the Modules folder
        # so the parent is the project root
        $moduleParent = Split-Path $PSScriptRoot -Parent
        $ConfigPath = Join-Path $moduleParent "ConfigFiles" "SITstoSkip.json"
    }

    $config = Read-JsonConfig -Path $ConfigPath
    if (-not $config) {
        Write-ExportLog -Message "SITs-to-skip config not found at: $ConfigPath (no SITs will be skipped)" -Level Warning
        return @()
    }

    $skipList = Get-EnabledItems -Config $config
    if ($skipList.Count -gt 0) {
        Write-ExportLog -Message "Loaded $($skipList.Count) SITs to skip from config" -Level Info
    }

    return $skipList
}

function Get-Compl8TenantInfo {
    <#
    .SYNOPSIS
        Gets the current tenant domain and ID from the active Security & Compliance connection.

    .DESCRIPTION
        Queries the ExchangeOnlineManagement connection information to extract the tenant
        domain and tenant ID. This is used for aggregate caching (to match aggregates
        to the correct tenant) and for logging purposes.

        Returns $null if no active compliance connection is found.

    .OUTPUTS
        Hashtable with TenantDomain (string) and TenantId (string), or $null if not connected.

    .EXAMPLE
        $tenant = Get-Compl8TenantInfo
        if ($tenant) { Write-Host "Connected to: $($tenant.TenantDomain)" }
    #>
    [CmdletBinding()]
    param()

    try {
        $connectionInfo = Get-ConnectionInformation -ErrorAction SilentlyContinue |
            Where-Object { $_.ConnectionUri -like "*compliance*" } |
            Select-Object -First 1

        if (-not $connectionInfo) {
            Write-ExportLog -Message "No active compliance connection found" -Level Warning
            return $null
        }

        # Extract tenant info from the connection
        # ConnectionUri is typically like: https://nam06b.ps.compliance.protection.outlook.com/PowerShell-LiveId
        # TenantId and Organization/UserPrincipalName provide tenant identity
        $tenantDomain = $null
        $tenantId = $null

        # Try to get TenantId from connection properties
        if ($connectionInfo.PSObject.Properties['TenantId']) {
            $tenantId = $connectionInfo.TenantId
        }

        # Try to extract domain from UserPrincipalName or Organization
        if ($connectionInfo.PSObject.Properties['UserPrincipalName'] -and $connectionInfo.UserPrincipalName) {
            $upn = $connectionInfo.UserPrincipalName
            if ($upn -match '@(.+)$') {
                $tenantDomain = $Matches[1]
            }
        }

        if (-not $tenantDomain -and $connectionInfo.PSObject.Properties['Organization']) {
            $tenantDomain = $connectionInfo.Organization
        }

        if (-not $tenantDomain -and -not $tenantId) {
            Write-ExportLog -Message "Could not extract tenant info from connection" -Level Warning
            return $null
        }

        return @{
            TenantDomain = $tenantDomain
            TenantId     = $tenantId
        }
    }
    catch {
        Write-ExportLog -Message "Failed to get tenant info: $($_.Exception.Message)" -Level Warning
        return $null
    }
}

function Get-SitGuidMapping {
    <#
    .SYNOPSIS
        Creates a lookup hashtable mapping Sensitive Information Type GUIDs to display names.

    .DESCRIPTION
        Calls Get-DlpSensitiveInformationType to retrieve all SIT definitions from the
        tenant, then builds a hashtable keyed by the SIT's Id (GUID) with the Name as
        the value. This mapping is used to resolve GUIDs that appear in Content Explorer
        results back to human-readable SIT names.

        The mapping is typically built once per export session and cached in the run tracker.

    .OUTPUTS
        Hashtable where keys are SIT GUIDs (string) and values are SIT display names (string).
        Returns empty hashtable on failure.

    .EXAMPLE
        $sitMapping = Get-SitGuidMapping
        $displayName = $sitMapping[$someGuid]  # Resolves GUID to friendly name
    #>
    [CmdletBinding()]
    param()

    Write-ExportLog -Message "Building SIT GUID-to-Name mapping..." -Level Info

    try {
        $sits = Get-DlpSensitiveInformationType -ErrorAction Stop
        $mapping = @{}

        foreach ($sit in $sits) {
            if ($sit.Id -and $sit.Name) {
                $mapping[$sit.Id.ToString()] = $sit.Name
            }
        }

        $sitCount = $mapping.Count
        Write-ExportLog -Message "  Mapped $sitCount Sensitive Information Types" -Level Success
        return $mapping
    }
    catch {
        Write-ExportLog -Message "  Failed to build SIT mapping: $($_.Exception.Message)" -Level Error
        return @{}
    }
}

#endregion

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

    $metadataPath = Join-Path (Get-CoordinationDir $ExportRunDirectory) "AggregateMetadata.json"

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

    if ($savedSettings -and $savedSettings.PageSize) {
        $cePageSize = $savedSettings.PageSize -as [int]
    }
    else {
        $cePageSize = if ($ceConfig -and $ceConfig._PageSize) { ($ceConfig._PageSize -as [int]) } else { $FallbackPageSize }
    }
    if (-not $cePageSize -or $cePageSize -lt 1) { $cePageSize = $FallbackPageSize }

    return [PSCustomObject]@{
        PageSize      = $cePageSize
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

#region Content Explorer - Work Plan & Aggregate Queries

function New-ContentExplorerWorkPlan {
    <#
    .SYNOPSIS
        Runs aggregate queries to build a Content Explorer export work plan.
    .DESCRIPTION
        For each TagName+Workload combination, calls Export-ContentExplorerData
        with -Aggregate to get location-level counts. Results are written to
        the aggregate CSV for progress tracking and potential reuse.
        Includes retry logic for transient errors.
    .PARAMETER TagType
        The classifier type (e.g. SensitiveInformationType, Sensitivity).
    .PARAMETER TagNames
        Array of tag names to query.
    .PARAMETER Workloads
        Array of workloads to query.
    .PARAMETER AggregateCsvPath
        Path to the aggregate CSV file for writing results.
    .PARAMETER ExportRunDirectory
        Directory for the current export run.
    .OUTPUTS
        Hashtable with:
          Tasks              - Array of task objects with TagType, TagName, Workload,
                               ExpectedCount, ExportedCount, Locations, Status
          TotalExpectedRecords - Sum of all expected record counts
          HasErrors          - Boolean indicating if any queries failed
          ErrorTasks         - Array of "TagName|Workload" strings that failed
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$TagType,

        [Parameter(Mandatory)]
        [string[]]$TagNames,

        [Parameter(Mandatory)]
        [string[]]$Workloads,

        [string]$AggregateCsvPath,

        [string]$ExportRunDirectory
    )

    $workPlan = @{
        Tasks                = @()
        TotalExpectedRecords = 0
        HasErrors            = $false
        ErrorTasks           = @()
    }

    $maxRetries = 3

    foreach ($tagName in $TagNames) {
        foreach ($workload in $Workloads) {
            Write-ExportLog -Message ("    Aggregate: " + $tagName + " / " + $workload) -Level Info

            $allAggregates = [System.Collections.ArrayList]::new()
            $pageCookie = $null
            $aggError = $null
            $aggSuccess = $false

            try {
                # Paginated aggregate query loop
                do {
                    $aggParams = @{
                        TagType     = $TagType
                        TagName     = $tagName
                        Workload    = $workload
                        PageSize    = 5000
                        Aggregate   = $true
                        ErrorAction = 'Stop'
                    }
                    if ($pageCookie) { $aggParams['PageCookie'] = $pageCookie }

                    $aggResult = Invoke-RetryWithBackoff -ScriptBlock {
                        Export-ContentExplorerData @aggParams
                    } -MaxRetries $maxRetries -Context ("Aggregate: " + $tagName + "/" + $workload)

                    # Process aggregate page results
                    if ($null -eq $aggResult -or $aggResult.Count -eq 0) { break }

                    $metadata = $aggResult[0]
                    if ($metadata.RecordsReturned -gt 0) {
                        $pageRecords = $aggResult[1..$metadata.RecordsReturned]
                        foreach ($rec in $pageRecords) {
                            [void]$allAggregates.Add($rec)
                        }
                    }

                    # Check for more pages
                    if ($metadata.MorePagesAvailable -eq $true -or $metadata.MorePagesAvailable -eq "True") {
                        $pageCookie = $metadata.PageCookie
                    }
                    else {
                        break
                    }
                } while ($true)

                $aggSuccess = $true
            }
            catch {
                $aggError = $_.Exception.Message
                Write-ExportLog -Message ("      AGGREGATE FAILED: " + $aggError) -Level Error
                $workPlan.HasErrors = $true
                $workPlan.ErrorTasks += ($tagName + "|" + $workload)
            }

            # Build task from aggregate results
            $locations = [System.Collections.ArrayList]::new()
            $totalCount = 0

            if ($aggSuccess) {
                foreach ($agg in $allAggregates) {
                    [void]$locations.Add(@{
                        Name          = $agg.Name
                        ExpectedCount = [int]$agg.Count
                        ExportedCount = 0
                    })
                    $totalCount += [int]$agg.Count
                }
                $displayCount = $totalCount.ToString('N0')
                $locationCount = $locations.Count
                Write-ExportLog -Message ("      -> " + $displayCount + " items in " + $locationCount + " locations") -Level Success
            }

            # Write to aggregate CSV (atomic append)
            if ($AggregateCsvPath) {
                $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")

                if ($aggSuccess) {
                    $csvLines = [System.Collections.ArrayList]::new()
                    foreach ($agg in $allAggregates) {
                        $locationName = $agg.Name -replace '"', '""'
                        if ($locationName -match '[,"]') {
                            $locationName = '"' + $locationName + '"'
                        }
                        $line = $timestamp + "," + $TagType + "," + $tagName + "," + $workload + "," + $locationName + "," + $agg.Count
                        [void]$csvLines.Add($line)
                    }

                    if ($csvLines.Count -gt 0) {
                        $csvContent = $csvLines -join [Environment]::NewLine
                        try {
                            [System.IO.File]::AppendAllText($AggregateCsvPath, $csvContent + [Environment]::NewLine, [System.Text.Encoding]::UTF8)
                        }
                        catch {
                            Write-ExportLog -Message ("      Failed to append to aggregate CSV: " + $_.Exception.Message) -Level Warning
                        }
                    }
                }
                else {
                    # Write error row
                    $escapedError = $aggError -replace '"', '""'
                    $errorLine = $timestamp + "," + $TagType + "," + $tagName + "," + $workload + ',ERROR,0,"' + $escapedError + '"'
                    try {
                        [System.IO.File]::AppendAllText($AggregateCsvPath, $errorLine + [Environment]::NewLine, [System.Text.Encoding]::UTF8)
                    }
                    catch {
                        Write-ExportLog -Message ("      Failed to write error to aggregate CSV: " + $_.Exception.Message) -Level Warning
                    }
                }
            }

            # Add task to work plan
            $task = @{
                TagType       = $TagType
                TagName       = $tagName
                Workload      = $workload
                ExpectedCount = $totalCount
                ExportedCount = 0
                Locations     = @($locations)
                Status        = if ($aggSuccess) { "Pending" } else { "Error" }
                PageMetrics   = @()
                ResponseTimes = @()
            }

            if ($aggError) {
                $task.AggregateError = $aggError
            }

            $workPlan.Tasks += $task
            $workPlan.TotalExpectedRecords += $totalCount
        }
    }

    return $workPlan
}

#endregion

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

#region Content Explorer - Detail Export with Pagination

function Export-ContentExplorerWithProgress {
    <#
    .SYNOPSIS
        Main detail export with pagination, progress tracking, and retries.
    .DESCRIPTION
        Calls Export-ContentExplorerData in a pagination loop with:
        - Retry logic for transient and non-transient errors
        - PageCookie anomaly detection (new cookie but no records)
        - Progress logging to a tailable file
        - Adaptive page size selection
        - Partial error tracking in the Task object
    .PARAMETER Task
        Hashtable with TagType, TagName, Workload, ExpectedCount, Locations.
        Modified in place: ExportedCount, Status, TotalPages, TotalTimeMs, PartialErrors.
    .PARAMETER PageSize
        Base page size for queries. May be overridden by adaptive sizing.
    .PARAMETER ProgressLogPath
        Path to progress log file (tailable).
    .PARAMETER Telemetry
        Telemetry tracking object from New-ContentExplorerTelemetry.
    .PARAMETER TelemetryDatabasePath
        Path to the JSONL telemetry database file.
    .PARAMETER AdaptivePageSize
        If set, calls Get-AdaptivePageSize to select optimal page size.
    .PARAMETER OutputDirectory
        Directory where per-page JSON files are written. Each page creates a separate file
        named {Workload}-{NNN}.json (or {Workload}-{LocationHash}-{NNN}.json for location-filtered tasks).
        Each file contains: {PageNumber, ExportTimestamp, TagType, TagName, Workload, RecordCount, Records}.
    .OUTPUTS
        Exported record count (int).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Task,

        [int]$PageSize = 1000,

        [string]$ProgressLogPath,

        $Telemetry,

        [string]$TelemetryDatabasePath,

        [switch]$AdaptivePageSize,

        [Parameter(Mandatory)]
        [string]$OutputDirectory,

        [string]$SiteUrl,

        [string]$UserPrincipalName
    )

    # Early parameter validation
    if ([string]::IsNullOrWhiteSpace($Task.TagName)) {
        throw "TagName cannot be null or empty"
    }
    if ([string]::IsNullOrWhiteSpace($Task.TagType)) {
        throw "TagType cannot be null or empty"
    }
    if ([string]::IsNullOrWhiteSpace($Task.Workload)) {
        throw "Workload cannot be null or empty"
    }

    $tagType = $Task.TagType
    $tagName = $Task.TagName
    $workload = $Task.Workload
    $expectedCount = if ($Task.ExpectedCount) { $Task.ExpectedCount -as [int] } else { 0 }

    # Select page size
    $effectivePageSize = $PageSize
    if ($AdaptivePageSize -and $Task.Locations -and $Task.Locations.Count -gt 0) {
        try {
            $effectivePageSize = Get-AdaptivePageSize -Task $Task -Workload $workload -TelemetryDatabasePath $TelemetryDatabasePath
            Write-ExportLog -Message ("      Adaptive page size: " + $effectivePageSize + " (workload: " + $workload + ")") -Level Info
        }
        catch {
            Write-ExportLog -Message ("      Adaptive page size failed, using default " + $PageSize + ": " + $_.Exception.Message) -Level Warning
            $effectivePageSize = $PageSize
        }
    }

    # Clamp page size: floor 100, ceiling 2x expected count
    if ($expectedCount -gt 0) {
        $maxPageSize = [Math]::Max(100, 2 * $expectedCount)
        $unclamped = $effectivePageSize
        $effectivePageSize = [Math]::Max(100, [Math]::Min($effectivePageSize, $maxPageSize))
        if ($effectivePageSize -ne $unclamped) {
            Write-ExportLog -Message ("      Page size clamped: {0} -> {1} (expected: {2})" -f $unclamped, $effectivePageSize, $expectedCount) -Level Info
        }
    } else {
        $effectivePageSize = [Math]::Max(100, $effectivePageSize)
    }

    # Initialize task tracking
    $Task.Status = "InProgress"
    $Task.ExportedCount = 0
    $Task.TotalPages = 0
    $Task.TotalTimeMs = 0
    if (-not $Task.ContainsKey('PartialErrors')) { $Task.PartialErrors = @() }

    # Build per-page filename prefix
    $locationSuffix = ""
    if ($SiteUrl) { $locationSuffix = "-" + ([Math]::Abs($SiteUrl.GetHashCode())).ToString("X8") }
    elseif ($UserPrincipalName) { $locationSuffix = "-" + ([Math]::Abs($UserPrincipalName.GetHashCode())).ToString("X8") }
    $pageFilePrefix = "{0}{1}" -f $workload, $locationSuffix

    # Ensure output directory exists
    if (-not (Test-Path $OutputDirectory)) {
        New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
    }

    $pageCookie = $null
    $previousCookie = $null
    $pageNumber = 0
    $maxRetries = 3
    $transientDelaySec = 60
    $nonTransientDelaySec = 5
    $finalAttemptDelaySec = 120
    $emptyPageRetried = $false
    $isRetryingSamePage = $false
    $startTime = Get-Date

    # Log export start
    $locationSuffix = if ($SiteUrl) { " [Site:$SiteUrl]" } elseif ($UserPrincipalName) { " [User:$UserPrincipalName]" } else { "" }
    $logEntry = "[{0}] START {1}/{2}/{3}{4} Expected:{5} PageSize:{6}" -f
        (Get-Date).ToString("HH:mm:ss"), $tagType, $tagName, $workload, $locationSuffix, $expectedCount, $effectivePageSize
    Write-ProgressEntry -LogPath $ProgressLogPath -Message $logEntry

    # Populate telemetry location fields
    if ($Telemetry) {
        if ($SiteUrl) {
            $Telemetry.Location = $SiteUrl
            $Telemetry.LocationType = "SiteUrl"
        } elseif ($UserPrincipalName) {
            $Telemetry.Location = $UserPrincipalName
            $Telemetry.LocationType = "UPN"
        } else {
            $Telemetry.LocationType = "WorkloadFallback"
        }
    }

    # Collect per-page metrics for telemetry
    $pageMetrics = [System.Collections.ArrayList]::new()

    try {
        do {
            if ($isRetryingSamePage) {
                $isRetryingSamePage = $false  # Reset flag; keep same page number for this retry
            }
            else {
                $pageNumber++
            }
            $pageStartTime = Get-Date

            # Build export parameters
            $exportParams = @{
                TagType     = $tagType
                TagName     = $tagName
                Workload    = $workload
                PageSize    = $effectivePageSize
                ErrorAction = 'Stop'
            }
            if ($pageCookie) { $exportParams['PageCookie'] = $pageCookie }
            if ($SiteUrl) { $exportParams['SiteUrl'] = $SiteUrl }
            if ($UserPrincipalName) { $exportParams['UserPrincipalName'] = $UserPrincipalName }

            $pageSuccess = $false
            $retryCount = 0
            $result = $null

            # Retry loop for current page
            while (-not $pageSuccess -and $retryCount -le $maxRetries) {
                try {
                    $result = Export-ContentExplorerData @exportParams
                    $pageSuccess = $true
                }
                catch {
                    # Connection lost - cmdlet not available; no point retrying
                    if ($_.Exception -is [System.Management.Automation.CommandNotFoundException]) {
                        Write-ExportLog -Message ("      Page " + $pageNumber + " FATAL: S&C cmdlet not available - connection lost") -Level Error
                        $Task.PartialErrors += @{
                            Page         = $pageNumber
                            RetryCount   = 0
                            ErrorMessage = $_.Exception.Message
                            IsTransient  = $false
                            Timestamp    = (Get-Date).ToString("o")
                            PageCookie   = $pageCookie
                            Location     = if ($SiteUrl) { $SiteUrl } elseif ($UserPrincipalName) { $UserPrincipalName } else { "" }
                        }
                        throw
                    }

                    $retryCount++
                    $errMsg = $_.Exception.Message
                    $errorInfo = Get-HttpErrorExplanation -ErrorMessage $errMsg -ErrorRecord $_
                    $isTransient = $errorInfo.IsTransient

                    # Track partial error
                    $partialError = @{
                        Page         = $pageNumber
                        RetryCount   = $retryCount
                        ErrorMessage = $errMsg
                        IsTransient  = $isTransient
                        Timestamp    = (Get-Date).ToString("o")
                        PageCookie   = $pageCookie
                        Location     = if ($SiteUrl) { $SiteUrl } elseif ($UserPrincipalName) { $UserPrincipalName } else { "" }
                    }
                    $Task.PartialErrors += $partialError

                    # Auth error - throw immediately for caller to handle (before exhaustion check)
                    if ($errorInfo.Category -eq "AuthError") {
                        Write-ExportLog -Message ("      Page " + $pageNumber + " AUTH ERROR - throwing for caller") -Level Error
                        throw
                    }

                    if ($retryCount -gt $maxRetries) {
                        # All retries exhausted
                        Write-ExportLog -Message ("      Page " + $pageNumber + " FAILED after " + $maxRetries + " retries: " + $errMsg) -Level Error
                        $logMsg = "[{0}] FAIL Page {1} after {2} retries: {3}" -f (Get-Date).ToString("HH:mm:ss"), $pageNumber, $maxRetries, $errMsg
                        Write-ProgressEntry -LogPath $ProgressLogPath -Message $logMsg
                        break
                    }

                    $delay = Get-RetryDelay -AttemptNumber $retryCount -MaxRetries $maxRetries `
                        -IsTransient $isTransient -TransientDelaySeconds $transientDelaySec `
                        -NonTransientDelaySeconds $nonTransientDelaySec -FinalAttemptDelaySeconds $finalAttemptDelaySec

                    $levelPrefix = if ($isTransient) { "Transient error" } else { "Error" }
                    Write-ExportLog -Message ("      " + $levelPrefix + " page " + $pageNumber + " (attempt " + $retryCount + "/" + $maxRetries + ") - retrying in " + $delay + "s") -Level Warning
                    Start-Sleep -Seconds $delay
                }
            }

            # Page failed after all retries
            if (-not $pageSuccess) {
                $Task.Status = "PartialFailure"
                break
            }

            # Process page results
            if ($null -eq $result -or $result.Count -eq 0) {
                break
            }

            $metadata = $result[0]
            $recordsInPage = $metadata.RecordsReturned -as [int]
            if ($null -eq $recordsInPage) { $recordsInPage = 0 }

            # On first page, correct expected count from API metadata if available
            if ($pageNumber -eq 1 -and $null -ne $metadata.TotalCount) {
                $apiTotalCount = $metadata.TotalCount -as [int]
                if ($apiTotalCount -and $apiTotalCount -gt 0 -and $apiTotalCount -ne $expectedCount) {
                    Write-ExportLog -Message ("      Expected count corrected: {0} -> {1} (from API metadata)" -f $expectedCount, $apiTotalCount) -Level Info
                    $expectedCount = $apiTotalCount
                    $Task.ExpectedCount = $apiTotalCount
                }
            }

            if ($recordsInPage -gt 0) {
                $pageRecords = $result[1..$recordsInPage]

                # Add export metadata to each record
                foreach ($record in $pageRecords) {
                    if ($record -is [PSCustomObject]) {
                        $record | Add-Member -NotePropertyName '_ExportTagType' -NotePropertyValue $tagType -Force
                        $record | Add-Member -NotePropertyName '_ExportTagName' -NotePropertyValue $tagName -Force
                    }
                }

                # Write per-page file
                $pageFileName = "{0}-{1:D3}.json" -f $pageFilePrefix, $pageNumber
                $pageFilePath = Join-Path $OutputDirectory $pageFileName

                $pageData = @{
                    PageNumber      = $pageNumber
                    ExportTimestamp  = (Get-Date).ToString("o")
                    TagType         = $tagType
                    TagName         = $tagName
                    Workload        = $workload
                    RecordCount     = $recordsInPage
                    Records         = @($pageRecords)
                }

                try {
                    $serializablePage = ConvertTo-SerializableObject -InputObject $pageData
                    $pageJson = $serializablePage | ConvertTo-Json -Depth 20
                    Set-Content -Path $pageFilePath -Value $pageJson -Encoding UTF8
                }
                catch {
                    Write-ExportLog -Message ("      Page {0}: Failed to save page file: {1}" -f $pageNumber, $_.Exception.Message) -Level Error
                }

                $Task.ExportedCount += $recordsInPage
                $emptyPageRetried = $false
            }

            # Page timing
            $pageElapsed = ((Get-Date) - $pageStartTime).TotalMilliseconds
            $Task.TotalPages = $pageNumber

            # Record per-page metric
            [void]$pageMetrics.Add(@{
                PageNumber   = $pageNumber
                PageTimeMs   = [int]$pageElapsed
                RecordCount  = $recordsInPage
                RetryCount   = $retryCount
                Timestamp    = (Get-Date).ToString("o")
            })

            # Log progress
            $totalExported = $Task.ExportedCount
            $pctStr = if ($expectedCount -gt 0) { [Math]::Round(($totalExported / $expectedCount) * 100, 1).ToString() + "%" } else { "N/A" }
            $logMsg = "[{0}] Page {1}: +{2} records (Total: {3}/{4} = {5}) [{6}ms]" -f
                (Get-Date).ToString("HH:mm:ss"), $pageNumber, $recordsInPage, $totalExported, $expectedCount, $pctStr, [int]$pageElapsed
            Write-ProgressEntry -LogPath $ProgressLogPath -Message $logMsg

            # Check for more pages
            $morePagesAvailable = ($metadata.MorePagesAvailable -eq $true -or $metadata.MorePagesAvailable -eq "True")
            if (-not $morePagesAvailable) {
                break
            }

            # PageCookie tracking
            $newCookie = $metadata.PageCookie

            if ($recordsInPage -eq 0 -and $newCookie -and $newCookie -ne $pageCookie) {
                # New cookie received but no records - anomaly
                if (-not $emptyPageRetried) {
                    # Retry with previous cookie once (same page number, not a new page)
                    Write-ExportLog -Message ("      Page " + $pageNumber + ": New cookie but 0 records - retrying previous cookie (30s wait)") -Level Warning
                    Start-Sleep -Seconds 30
                    $emptyPageRetried = $true
                    $isRetryingSamePage = $true  # Signal loop to skip page increment on next iteration
                    # Keep using current pageCookie (previous cookie)
                    continue
                }
                else {
                    # Already retried, continue with new cookie
                    Write-ExportLog -Message ("      Page " + $pageNumber + ": Still 0 records after retry - continuing with new cookie") -Level Warning
                    $emptyPageRetried = $false
                }
            }

            $previousCookie = $pageCookie
            $pageCookie = $newCookie

        } while ($true)
    }
    catch {
        Write-ExportLog -Message ("      Export exception: " + $_.Exception.Message) -Level Error
        $Task.Status = "Failed"
    }

    # Finalize task
    $totalElapsed = ((Get-Date) - $startTime).TotalMilliseconds
    $Task.TotalTimeMs = [int]$totalElapsed

    if ($Task.Status -eq "InProgress") {
        $Task.Status = "Completed"
    }

    # Log completion
    $totalSec = [Math]::Round($totalElapsed / 1000, 1)
    $logMsg = "[{0}] END {1}/{2}/{3} Records:{4} Pages:{5} Time:{6}s Status:{7}" -f
        (Get-Date).ToString("HH:mm:ss"), $tagType, $tagName, $workload,
        $Task.ExportedCount, $Task.TotalPages, $totalSec, $Task.Status
    Write-ProgressEntry -LogPath $ProgressLogPath -Message $logMsg

    # Save telemetry
    if ($Telemetry) {
        $Telemetry.RecordCount = $Task.ExportedCount
        $Telemetry.PageCount = $Task.TotalPages
        $Telemetry.TotalTimeMs = $Task.TotalTimeMs
        $Telemetry.PageSize = $effectivePageSize
        $Telemetry.Status = $Task.Status
        $Telemetry.CompletedTime = (Get-Date).ToString("o")
        $Telemetry.PageMetrics = @($pageMetrics)

        if ($TelemetryDatabasePath) {
            Save-ContentExplorerTelemetry -Telemetry $Telemetry -DatabasePath $TelemetryDatabasePath
        }
    }

    # Write _task.json summary
    if ($Task.ExportedCount -gt 0) {
        $taskSummary = @{
            TagType        = $tagType
            TagName        = $tagName
            Workload       = $workload
            ExpectedCount  = $expectedCount
            ActualCount    = $Task.ExportedCount
            Pages          = $Task.TotalPages
            Status         = $Task.Status
            TotalTimeMs    = $Task.TotalTimeMs
            ExportDate     = (Get-Date).ToString("o")
            PageFilePrefix = $pageFilePrefix
        }
        if ($Task.PartialErrors.Count -gt 0) {
            $taskSummary.PartialErrors = @($Task.PartialErrors)
        }
        try {
            $taskJsonPath = Join-Path $OutputDirectory ("_task-{0}.json" -f $pageFilePrefix)
            $taskSummary | ConvertTo-Json -Depth 10 | Set-Content -Path $taskJsonPath -Encoding UTF8
        }
        catch {
            Write-ExportLog -Message ("      Failed to write _task summary: " + $_.Exception.Message) -Level Warning
        }
    }

    return $Task.ExportedCount
}

#region Content Explorer - Deduplication

function Remove-DuplicateContentRecordsV2 {
    <#
    .SYNOPSIS
        Deduplicates Content Explorer records with multi-classifier tracking.
    .DESCRIPTION
        Identifies duplicate records using a composite key built from:
        - ContentUri (preferred, if available)
        - Fallback: FileName + FolderPath + LocationUrl
        Each unique record gets a _MatchingClassifiers array listing all
        classifiers that matched (e.g. "SensitiveInformationType:Credit Card Number").
        Records matching multiple classifiers are tracked separately.
    .PARAMETER Records
        Array of Content Explorer records to deduplicate.
    .OUTPUTS
        Hashtable with:
          UniqueRecords    - Array of deduplicated records
          MultiMatchRecords - Array of records matching more than one classifier
          DuplicateCount   - Number of duplicate records removed
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [array]$Records
    )

    $uniqueMap = [System.Collections.Generic.Dictionary[string, object]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )

    $totalInput = $Records.Count
    $duplicateCount = 0

    foreach ($record in $Records) {
        # Build unique key
        $uniqueKey = $null

        if ($record.ContentUri) {
            $uniqueKey = $record.ContentUri
        }
        elseif ($record.FileName -and $record.FolderPath) {
            $locationPart = if ($record.LocationUrl) { $record.LocationUrl } else { "" }
            $uniqueKey = $record.FileName + "|" + $record.FolderPath + "|" + $locationPart
        }
        elseif ($record.FileName -and $record.LocationUrl) {
            $uniqueKey = $record.FileName + "|" + $record.LocationUrl
        }
        elseif ($record.ItemName -and $record.Location) {
            $uniqueKey = $record.ItemName + "|" + $record.Location
        }
        else {
            # Fallback: serialize the record to build a key from all properties
            $keyParts = @()
            foreach ($prop in $record.PSObject.Properties) {
                if ($prop.Name -notlike '_*') {
                    $keyParts += ($prop.Name + "=" + $prop.Value)
                }
            }
            $uniqueKey = $keyParts -join "|"
        }

        # Build classifier label
        $classifierLabel = ""
        if ($record._ExportTagType -and $record._ExportTagName) {
            $classifierLabel = $record._ExportTagType + ":" + $record._ExportTagName
        }

        if ($uniqueMap.ContainsKey($uniqueKey)) {
            # Duplicate found - merge classifier info
            $existing = $uniqueMap[$uniqueKey]
            $duplicateCount++

            # Add classifier if not already present
            if ($classifierLabel -and $existing._MatchingClassifiers -notcontains $classifierLabel) {
                $existing._MatchingClassifiers += $classifierLabel
            }
        }
        else {
            # New unique record
            $record | Add-Member -NotePropertyName '_UniqueKey' -NotePropertyValue $uniqueKey -Force
            $record | Add-Member -NotePropertyName '_MatchingClassifiers' -NotePropertyValue @() -Force

            if ($classifierLabel) {
                $record._MatchingClassifiers = @($classifierLabel)
            }

            $uniqueMap[$uniqueKey] = $record
        }
    }

    # Build output arrays
    $uniqueRecords = [System.Collections.ArrayList]::new()
    $multiMatchRecords = [System.Collections.ArrayList]::new()

    foreach ($entry in $uniqueMap.Values) {
        [void]$uniqueRecords.Add($entry)

        if ($entry._MatchingClassifiers -and $entry._MatchingClassifiers.Count -gt 1) {
            [void]$multiMatchRecords.Add($entry)
        }
    }

    Write-ExportLog -Message ("  Dedup: " + $totalInput + " raw -> " + $uniqueRecords.Count + " unique (" + $duplicateCount + " duplicates, " + $multiMatchRecords.Count + " multi-match)") -Level Info

    return @{
        UniqueRecords    = @($uniqueRecords)
        MultiMatchRecords = @($multiMatchRecords)
        DuplicateCount   = $duplicateCount
    }
}

#endregion

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

#region Content Explorer - Progress Tracking & Display

function Get-ContentExplorerAggregateProgress {
    <#
    .SYNOPSIS
        Calculates aggregate progress for Content Explorer export.
    .DESCRIPTION
        Reads the aggregate CSV to determine total expected records and tasks,
        then merges with completed task counts to calculate progress.
    .PARAMETER AggregateCsvPath
        Path to the ContentExplorer-Aggregates.csv file.
    .PARAMETER CompletedTasks
        Hashtable of completed task keys mapped to their exported record counts.
    .OUTPUTS
        Hashtable with:
          TotalExpected  - Total expected records from all tasks
          TotalExported  - Total records exported so far
          CompletedTasks - Number of completed tasks
          TotalTasks     - Total number of tasks
          Tasks          - Hashtable of task details keyed by "TagType|TagName|Workload"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$AggregateCsvPath,

        [hashtable]$CompletedTasks = @{}
    )

    $progress = @{
        TotalExpected  = 0
        TotalExported  = 0
        CompletedTasks = 0
        TotalTasks     = 0
        Tasks          = @{}
    }

    if (-not (Test-Path $AggregateCsvPath)) {
        return $progress
    }

    try {
        $csvData = Import-Csv -Path $AggregateCsvPath -Encoding UTF8 -ErrorAction Stop
    }
    catch {
        Write-ExportLog -Message ("Failed to read aggregate CSV for progress: " + $_.Exception.Message) -Level Warning
        return $progress
    }

    # Aggregate by task key
    $taskTotals = @{}
    foreach ($row in $csvData) {
        if ($row.Location -eq "ERROR") { continue }

        $taskKey = "{0}|{1}|{2}" -f $row.TagType, $row.TagName, $row.Workload

        if (-not $taskTotals.ContainsKey($taskKey)) {
            $taskTotals[$taskKey] = @{
                TagType       = $row.TagType
                TagName       = $row.TagName
                Workload      = $row.Workload
                ExpectedCount = 0
                ExportedCount = 0
                IsCompleted   = $false
            }
        }

        $count = $row.Count -as [int]
        if ($count) { $taskTotals[$taskKey].ExpectedCount += $count }
    }

    # Merge with completed task data
    foreach ($taskKey in $taskTotals.Keys) {
        $taskInfo = $taskTotals[$taskKey]
        $progress.TotalExpected += $taskInfo.ExpectedCount
        $progress.TotalTasks++

        if ($CompletedTasks.ContainsKey($taskKey)) {
            $exported = $CompletedTasks[$taskKey] -as [int]
            $taskInfo.ExportedCount = $exported
            $taskInfo.IsCompleted = $true
            $progress.TotalExported += $exported
            $progress.CompletedTasks++
        }
    }

    $progress.Tasks = $taskTotals
    return $progress
}

function Write-ContentExplorerProgress {
    <#
    .SYNOPSIS
        Displays formatted Content Explorer export progress.
    .DESCRIPTION
        Shows overall progress and highlights the current task being processed.
    .PARAMETER Progress
        Progress hashtable from Get-ContentExplorerAggregateProgress.
    .PARAMETER CurrentTaskKey
        The "TagType|TagName|Workload" key of the task currently being exported.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Progress,

        [string]$CurrentTaskKey
    )

    $totalTasks = $Progress.TotalTasks
    $completedTasks = $Progress.CompletedTasks
    $totalExpected = $Progress.TotalExpected
    $totalExported = $Progress.TotalExported

    $taskPct = if ($totalTasks -gt 0) { [Math]::Round(($completedTasks / $totalTasks) * 100, 0) } else { 0 }
    $recordPct = if ($totalExpected -gt 0) { [Math]::Round(($totalExported / $totalExpected) * 100, 1) } else { 0 }

    Write-ExportLog -Message ("  Progress: Tasks " + $completedTasks + "/" + $totalTasks + " (" + $taskPct + "%) | Records " + $totalExported.ToString('N0') + "/" + $totalExpected.ToString('N0') + " (" + $recordPct + "%)") -Level Info

    # Show current task info
    if ($CurrentTaskKey -and $Progress.Tasks.ContainsKey($CurrentTaskKey)) {
        $current = $Progress.Tasks[$CurrentTaskKey]
        Write-ExportLog -Message ("  Current: " + $current.TagName + " / " + $current.Workload + " (expected: " + $current.ExpectedCount.ToString('N0') + ")") -Level Info
    }
}

#endregion

#region Activity Explorer Retry Helpers

function Get-PageErrorMessage {
    <#
    .SYNOPSIS
        Extracts a clean error message from an error record.
    .PARAMETER ErrorRecord
        The PowerShell error record to extract the message from.
    .OUTPUTS
        String containing the cleaned error message.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $ErrorRecord
    )

    if ($null -eq $ErrorRecord) {
        return "Unknown error (null error record)"
    }

    # Try the exception message first
    if ($ErrorRecord -is [System.Management.Automation.ErrorRecord]) {
        if ($ErrorRecord.Exception -and $ErrorRecord.Exception.Message) {
            return $ErrorRecord.Exception.Message
        }
        if ($ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message) {
            return $ErrorRecord.ErrorDetails.Message
        }
        return $ErrorRecord.ToString()
    }

    # Handle plain exceptions
    if ($ErrorRecord -is [System.Exception]) {
        return $ErrorRecord.Message
    }

    # Fallback: convert to string
    return "$ErrorRecord"
}

function Test-PageHasContent {
    <#
    .SYNOPSIS
        Checks if an Activity Explorer API result page has actual records.
    .PARAMETER Result
        The API response object from Export-ActivityExplorerData.
    .OUTPUTS
        Boolean indicating whether the page contains records.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Result
    )

    if ($null -eq $Result) {
        return $false
    }

    if ([string]::IsNullOrWhiteSpace($Result.ResultData)) {
        return $false
    }

    try {
        $records = $Result.ResultData | ConvertFrom-Json
        if ($null -eq $records) {
            return $false
        }
        return (@($records).Count -gt 0)
    }
    catch {
        Write-Verbose "Test-PageHasContent: Failed to parse ResultData: $($_.Exception.Message)"
        return $false
    }
}

function Add-PartialError {
    <#
    .SYNOPSIS
        Adds an error entry to the tracker's PartialErrors array.
    .PARAMETER Tracker
        The run tracker hashtable to update.
    .PARAMETER PageNumber
        The page number where the error occurred.
    .PARAMETER ErrorMessage
        The error message to record.
    .PARAMETER ErrorType
        Classification of the error (e.g., "SameCookieRetryExhausted", "ApiError", "FutureDateError").
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Tracker,

        [Parameter(Mandatory)]
        [int]$PageNumber,

        [Parameter(Mandatory)]
        [string]$ErrorMessage,

        [Parameter(Mandatory)]
        [string]$ErrorType
    )

    if (-not $Tracker.ContainsKey('PartialErrors') -or $null -eq $Tracker['PartialErrors']) {
        $Tracker['PartialErrors'] = @()
    }

    $errorEntry = @{
        Timestamp    = (Get-Date).ToString("o")
        PageNumber   = $PageNumber
        ErrorType    = $ErrorType
        ErrorMessage = $ErrorMessage
    }

    $Tracker['PartialErrors'] = @($Tracker['PartialErrors']) + @($errorEntry)
}

function Get-RetryDelay {
    <#
    .SYNOPSIS
        Calculates retry delay based on error classification and attempt number.
    .DESCRIPTION
        Centralizes the retry delay calculation pattern used across Content Explorer
        and Activity Explorer pagination. Delay varies by transient vs non-transient
        errors, with a longer final-attempt delay.
    .PARAMETER AttemptNumber
        Current retry attempt (1-based).
    .PARAMETER MaxRetries
        Maximum number of retries allowed.
    .PARAMETER IsTransient
        Whether the error is transient (server/network). Transient gets longer delays.
    .PARAMETER TransientDelaySeconds
        Base delay for transient errors. Default: 60
    .PARAMETER NonTransientDelaySeconds
        Delay for non-transient errors. Default: 5
    .PARAMETER FinalAttemptDelaySeconds
        Delay before the final retry attempt. Default: 120
    .PARAMETER ScaleByAttempt
        If set, multiplies transient delay by attempt number. Default: $false (flat delay).
    .OUTPUTS
        Integer - number of seconds to wait before retrying.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [int]$AttemptNumber,

        [int]$MaxRetries = 3,

        [bool]$IsTransient = $true,

        [int]$TransientDelaySeconds = 60,

        [int]$NonTransientDelaySeconds = 5,

        [int]$FinalAttemptDelaySeconds = 120,

        [switch]$ScaleByAttempt
    )

    # Final attempt: always use the longer delay regardless of error type
    if ($AttemptNumber -ge $MaxRetries) {
        return $FinalAttemptDelaySeconds
    }

    if ($IsTransient) {
        if ($ScaleByAttempt) {
            return ($TransientDelaySeconds * $AttemptNumber)
        }
        return $TransientDelaySeconds
    }

    return $NonTransientDelaySeconds
}

function Invoke-RetryWithBackoff {
    <#
    .SYNOPSIS
        Generic retry wrapper with backoff delays.
    .DESCRIPTION
        Executes a script block with retry logic. Delay pattern varies based on whether
        the error is transient (longer delays) or non-transient (shorter delays).
        On the final attempt, uses a longer delay to allow recovery.
    .PARAMETER ScriptBlock
        The script block to execute with retry.
    .PARAMETER MaxRetries
        Maximum number of retry attempts. Default: 3
    .PARAMETER InitialDelaySeconds
        Base delay for transient errors (multiplied by attempt number). Default: 60
    .PARAMETER FinalAttemptDelaySeconds
        Delay before the final retry attempt. Default: 120
    .PARAMETER Context
        Description of the operation for logging purposes.
    .OUTPUTS
        Returns the script block's result on success. Throws on final failure.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [scriptblock]$ScriptBlock,

        [int]$MaxRetries = 3,

        [int]$InitialDelaySeconds = 60,

        [int]$FinalAttemptDelaySeconds = 120,

        [string]$Context = "operation"
    )

    $attempt = 0
    $lastError = $null

    while ($attempt -le $MaxRetries) {
        try {
            $result = & $ScriptBlock
            return $result
        }
        catch {
            $lastError = $_
            $attempt++
            $errorMessage = Get-PageErrorMessage -ErrorRecord $_

            # Connection lost - cmdlet not available; no point retrying
            if ($_.Exception -is [System.Management.Automation.CommandNotFoundException]) {
                Write-ExportLog -Message ("  FATAL during {0}: S&C cmdlet not available - connection lost" -f $Context) -Level Error
                throw
            }

            # Check if this is an auth error - do not retry, throw immediately
            $errorInfo = Get-HttpErrorExplanation -ErrorMessage $errorMessage -ErrorRecord $_
            if ($errorInfo.Category -eq "AuthError") {
                Write-ExportLog -Message ("  AUTH ERROR during {0} - not retrying" -f $Context) -Level Error
                throw
            }

            if ($attempt -gt $MaxRetries) {
                # All retries exhausted
                $msg = "  {0}: All {1} retries exhausted. Last error: {2}" -f $Context, $MaxRetries, $errorMessage
                Write-ExportLog -Message $msg -Level Error
                throw
            }

            $isTransient = $errorInfo.IsTransient
            $statusStr = if ($errorInfo.StatusCode) { "HTTP {0}" -f $errorInfo.StatusCode } else { $errorInfo.Category }

            $delay = Get-RetryDelay -AttemptNumber $attempt -MaxRetries $MaxRetries `
                -IsTransient $isTransient -TransientDelaySeconds $InitialDelaySeconds `
                -FinalAttemptDelaySeconds $FinalAttemptDelaySeconds -ScaleByAttempt

            $levelPrefix = if ($isTransient -and $attempt -lt $MaxRetries) { "TRANSIENT " } else { "" }
            $msg = "  {0}: {1}{2} (attempt {3}/{4}) - waiting {5}s" -f $Context, $levelPrefix, $statusStr, $attempt, $MaxRetries, $delay
            Write-ExportLog -Message $msg -Level Warning
            Start-Sleep -Seconds $delay
        }
    }

    # Should not reach here, but safety net
    throw $lastError
}

#endregion

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

#region Activity Explorer Export Functions

function Export-ActivityExplorerWithProgress {
    <#
    .SYNOPSIS
        Resilient per-page Activity Explorer export with progress tracking and resume capability.
    .DESCRIPTION
        Exports Activity Explorer data page by page, saving each page immediately to disk.
        Supports resume from last successful page after failures or interruptions.
        Includes retry logic for transient errors and automatic end-time adjustment for
        "future dates" API errors.

    .PARAMETER StartTime
        Start of the date range to export (UTC).
    .PARAMETER EndTime
        End of the date range to export (UTC).
    .PARAMETER PageSize
        Records per page (1-5000). Default: 5000
    .PARAMETER Filters
        Hashtable of filters (Activity, Workload, etc.).
    .PARAMETER OutputDirectory
        Path to the ActivityExplorer output subfolder.
    .PARAMETER Tracker
        Run tracker hashtable for state persistence.
    .PARAMETER TrackerPath
        File path where the run tracker is saved.
    .PARAMETER ProgressLogPath
        File path for the tailable progress log.
    .PARAMETER Resume
        Switch to enable resume from last successful page.

    .OUTPUTS
        Hashtable with TotalRecords, PageCount, and optionally ResumedFrom.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [datetime]$StartTime,

        [Parameter(Mandatory)]
        [datetime]$EndTime,

        [ValidateRange(1, 5000)]
        [int]$PageSize = 5000,

        [hashtable]$Filters,

        [Parameter(Mandatory)]
        [string]$OutputDirectory,

        [Parameter(Mandatory)]
        [hashtable]$Tracker,

        [Parameter(Mandatory)]
        [string]$TrackerPath,

        [string]$ProgressLogPath,

        [switch]$Resume
    )

    # Early parameter validation
    if ($StartTime -ge $EndTime) {
        throw "StartTime ($($StartTime.ToString('o'))) must be before EndTime ($($EndTime.ToString('o')))"
    }

    # Ensure output directory exists
    if (-not (Test-Path $OutputDirectory)) {
        New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
    }

    # Format dates for the cmdlet (MM/dd/yyyy HH:mm:ss format)
    $startStr = $StartTime.ToString("MM/dd/yyyy HH:mm:ss")
    $currentEndTime = $EndTime
    $endStr = $currentEndTime.ToString("MM/dd/yyyy HH:mm:ss")

    # Build base export parameters
    $exportParams = @{
        StartTime    = $startStr
        EndTime      = $endStr
        OutputFormat = "Json"
        PageSize     = $PageSize
        ErrorAction  = 'Stop'
    }

    # Add filters (up to 5 supported by the API)
    if ($Filters -and $Filters.Count -gt 0) {
        if ($Filters.Count -gt 5) {
            Write-Warning "Activity Explorer API supports max 5 filters per category. $($Filters.Count) items enabled — only the first 5 will be used. Disable some items in the config file."
        }
        $filterIndex = 1
        foreach ($filterName in $Filters.Keys) {
            if ($filterIndex -le 5) {
                $filterValues = @($filterName) + @($Filters[$filterName])
                $exportParams["Filter$filterIndex"] = $filterValues
                $filterIndex++
            }
        }
    }

    # Initialize tracking variables
    $pageNumber = 0
    $totalRecords = 0
    $resumedFrom = $null
    $previousWaterMark = $null

    # Handle resume
    if ($Resume -and $Tracker.LastWaterMark) {
        $pageNumber = $Tracker.CompletedPages
        $totalRecords = $Tracker.TotalRecords
        $exportParams['PageCookie'] = $Tracker.LastWaterMark
        $previousWaterMark = $Tracker.LastWaterMark

        $resumedFrom = @{
            PageNumber  = $pageNumber
            RecordCount = $totalRecords
        }

        $msg = "  RESUMING from page {0} ({1} records already exported)" -f $pageNumber, $totalRecords
        Write-ExportLog -Message $msg -Level Info
        Write-ProgressEntry -Path $ProgressLogPath -Message $msg
    }
    else {
        $msg = "  Starting Activity Explorer export: {0} to {1}" -f $startStr, $endStr
        Write-ExportLog -Message $msg -Level Info
        Write-ProgressEntry -Path $ProgressLogPath -Message $msg
    }

    # Initial query with future-dates retry
    $result = $null
    $futureDateRetries = 0
    $maxFutureDateRetries = 3

    while ($null -eq $result -and $futureDateRetries -le $maxFutureDateRetries) {
        try {
            $result = Export-ActivityExplorerData @exportParams
        }
        catch {
            $errorMsg = Get-PageErrorMessage -ErrorRecord $_
            $isFutureDate = $errorMsg -match "future" -and $errorMsg -match "date"

            if ($isFutureDate -and $futureDateRetries -lt $maxFutureDateRetries) {
                $futureDateRetries++
                $currentEndTime = $currentEndTime.AddHours(-1)
                $endStr = $currentEndTime.ToString("MM/dd/yyyy HH:mm:ss")
                $exportParams['EndTime'] = $endStr

                $msg = "  Future date error - reducing end time by 1 hour (attempt {0}/{1}): {2}" -f $futureDateRetries, $maxFutureDateRetries, $endStr
                Write-ExportLog -Message $msg -Level Warning
                Write-ProgressEntry -Path $ProgressLogPath -Message $msg

                Add-PartialError -Tracker $Tracker -PageNumber 0 -ErrorMessage $errorMsg -ErrorType "FutureDateRetry"
            }
            else {
                # Check for auth error
                $initialErrorInfo = Get-HttpErrorExplanation -ErrorMessage $errorMsg -ErrorRecord $_
                if ($initialErrorInfo.Category -eq "AuthError") {
                    Write-ExportLog -Message "  AUTH ERROR - throwing for caller to handle" -Level Error
                    throw
                }

                # Non-recoverable error on initial query
                $msg = "  Initial query failed: {0}" -f $errorMsg
                Write-ExportLog -Message $msg -Level Error
                Write-ProgressEntry -Path $ProgressLogPath -Message $msg
                Add-PartialError -Tracker $Tracker -PageNumber 0 -ErrorMessage $errorMsg -ErrorType "InitialQueryFailed"

                $Tracker['Status'] = "Failed"
                $Tracker['PartialFailure'] = $true
                Save-ActivityExplorerRunTracker -Tracker $Tracker -TrackerPath $TrackerPath

                return @{
                    TotalRecords = $totalRecords
                    PageCount    = $pageNumber
                    ResumedFrom  = $resumedFrom
                }
            }
        }
    }

    # Check if we got any data
    if ($null -eq $result) {
        Write-ExportLog -Message "  No results returned after future-date retries" -Level Warning
        $Tracker['Status'] = "Completed"
        Save-ActivityExplorerRunTracker -Tracker $Tracker -TrackerPath $TrackerPath
        return @{
            TotalRecords = 0
            PageCount    = 0
            ResumedFrom  = $resumedFrom
        }
    }

    # Log total result count
    $totalAvailable = $result.TotalResultCount
    if ($totalAvailable) {
        $msg = "  Total activities available: {0}" -f $totalAvailable
        Write-ExportLog -Message $msg -Level Info
        Write-ProgressEntry -Path $ProgressLogPath -Message $msg
        $Tracker['TotalAvailable'] = $totalAvailable -as [long]
        Save-ActivityExplorerRunTracker -Tracker $Tracker -TrackerPath $TrackerPath
    }

    # Page processing loop
    $exportStartTime = Get-Date
    $done = $false

    while (-not $done) {
        $pageNumber++
        $pageStartTime = Get-Date

        # Parse records from this page
        $pageRecords = @()
        $hasContent = Test-PageHasContent -Result $result

        if ($hasContent) {
            try {
                $parsed = $result.ResultData | ConvertFrom-Json
                $pageRecords = @($parsed)
            }
            catch {
                $msg = "  Page {0}: Failed to parse ResultData: {1}" -f $pageNumber, $_.Exception.Message
                Write-ExportLog -Message $msg -Level Warning
                Write-ProgressEntry -Path $ProgressLogPath -Message $msg
            }
        }

        $recordCount = $pageRecords.Count

        if ($recordCount -gt 0) {
            # Calculate record time range
            $recordTimeRange = @{
                Earliest = $null
                Latest   = $null
            }

            try {
                $timestamps = @($pageRecords | Where-Object { $_.Happened } | ForEach-Object {
                    if ($_.Happened -is [datetime]) { $_.Happened } else { [DateTime]::Parse($_.Happened, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::RoundtripKind) }
                })
                if ($timestamps.Count -gt 0) {
                    $recordTimeRange.Earliest = ($timestamps | Measure-Object -Minimum).Minimum.ToString("o")
                    $recordTimeRange.Latest = ($timestamps | Measure-Object -Maximum).Maximum.ToString("o")
                }
            }
            catch {
                # Timestamp parsing is optional for progress display - non-critical
                Write-Verbose "Timestamp parsing failed for page $pageNumber : $($_.Exception.Message)"
            }

            # Save page file immediately
            $pageFileName = "Page-{0:D3}.json" -f $pageNumber
            $pageFilePath = Join-Path $OutputDirectory $pageFileName

            $pageData = @{
                PageNumber      = $pageNumber
                ExportTimestamp = (Get-Date).ToString("o")
                RecordTimeRange = $recordTimeRange
                RecordCount     = $recordCount
                WaterMark       = $result.WaterMark
                Records         = $pageRecords
            }

            try {
                $serializablePage = ConvertTo-SerializableObject -InputObject $pageData
                $pageJson = $serializablePage | ConvertTo-Json -Depth 20
                Set-Content -Path $pageFilePath -Value $pageJson -Encoding UTF8
            }
            catch {
                $msg = "  Page {0}: Failed to save page file: {1}" -f $pageNumber, $_.Exception.Message
                Write-ExportLog -Message $msg -Level Error
                Write-ProgressEntry -Path $ProgressLogPath -Message $msg
            }

            $totalRecords += $recordCount
        }

        # Update tracker
        $pageElapsed = (Get-Date) - $pageStartTime
        $Tracker['CompletedPages'] = $pageNumber
        $Tracker['TotalRecords'] = $totalRecords
        $Tracker['LastWaterMark'] = $result.WaterMark
        $Tracker['LastPageTime'] = $pageElapsed.TotalSeconds

        Save-ActivityExplorerRunTracker -Tracker $Tracker -TrackerPath $TrackerPath

        # Progress logging
        $totalElapsed = (Get-Date) - $exportStartTime
        $pagesPerMin = if ($totalElapsed.TotalMinutes -gt 0) {
            [math]::Round($pageNumber / $totalElapsed.TotalMinutes, 1)
        }
        else { 0 }

        $pctMsg = ""
        if ($totalAvailable -and $totalAvailable -gt 0) {
            $pct = [Math]::Round(($totalRecords / $totalAvailable) * 100, 1)
            $pctMsg = ", {0}% of {1:N0}" -f $pct, $totalAvailable
        }
        $progressMsg = "  Page {0}: {1} records ({2:N0} total{3}, {4} pages/min)" -f $pageNumber, $recordCount, $totalRecords, $pctMsg, $pagesPerMin
        Write-ExportLog -Message $progressMsg -Level Info
        Write-ProgressEntry -Path $ProgressLogPath -Message $progressMsg

        # Check if this is the last page
        if ($result.LastPage -eq $true) {
            $done = $true
            continue
        }

        # Get next page with retry logic
        $newWaterMark = $result.WaterMark
        $previousWaterMark = if ($exportParams.ContainsKey('PageCookie')) { $exportParams['PageCookie'] } else { $null }
        $exportParams['PageCookie'] = $newWaterMark

        $nextResult = $null
        $retryCount = 0
        $maxRetries = 3
        $sameCookieRetries = 0

        while ($null -eq $nextResult) {
            try {
                $nextResult = Export-ActivityExplorerData @exportParams
            }
            catch {
                $errorMsg = Get-PageErrorMessage -ErrorRecord $_
                $errorInfo = Get-HttpErrorExplanation -ErrorMessage $errorMsg -ErrorRecord $_

                # Auth error - throw immediately for caller to handle
                if ($errorInfo.Category -eq "AuthError") {
                    Write-ExportLog -Message "  AUTH ERROR on page retrieval - throwing for caller to handle" -Level Error
                    throw
                }

                # Future dates error on subsequent pages
                $isFutureDate = $errorMsg -match "future" -and $errorMsg -match "date"
                if ($isFutureDate) {
                    $futureDateRetries++
                    if ($futureDateRetries -le $maxFutureDateRetries) {
                        $currentEndTime = $currentEndTime.AddHours(-1)
                        $endStr = $currentEndTime.ToString("MM/dd/yyyy HH:mm:ss")
                        $exportParams['EndTime'] = $endStr

                        $msg = "  Future date error on page {0} - reducing end time: {1}" -f ($pageNumber + 1), $endStr
                        Write-ExportLog -Message $msg -Level Warning
                        Write-ProgressEntry -Path $ProgressLogPath -Message $msg
                        Add-PartialError -Tracker $Tracker -PageNumber ($pageNumber + 1) -ErrorMessage $errorMsg -ErrorType "FutureDateRetry"
                        continue
                    }
                }

                $retryCount++
                Add-PartialError -Tracker $Tracker -PageNumber ($pageNumber + 1) -ErrorMessage $errorMsg -ErrorType "PageRetrievalError"

                if ($retryCount -gt $maxRetries) {
                    # All retries exhausted - save what we have
                    $msg = "  Page {0}: All {1} retries exhausted. Saving progress." -f ($pageNumber + 1), $maxRetries
                    Write-ExportLog -Message $msg -Level Error
                    Write-ProgressEntry -Path $ProgressLogPath -Message $msg

                    $Tracker['PartialFailure'] = $true
                    $Tracker['Status'] = "PartialFailure"
                    Save-ActivityExplorerRunTracker -Tracker $Tracker -TrackerPath $TrackerPath

                    return @{
                        TotalRecords = $totalRecords
                        PageCount    = $pageNumber
                        ResumedFrom  = $resumedFrom
                    }
                }

                $delay = Get-RetryDelay -AttemptNumber $retryCount -MaxRetries $maxRetries `
                    -IsTransient $true -ScaleByAttempt
                $msg = "  Page {0}: Retry {1}/{2} - waiting {3}s" -f ($pageNumber + 1), $retryCount, $maxRetries, $delay
                Write-ExportLog -Message $msg -Level Warning
                Write-ProgressEntry -Path $ProgressLogPath -Message $msg
                Start-Sleep -Seconds $delay
            }
        }

        # Analyze the new result
        $nextHasContent = Test-PageHasContent -Result $nextResult
        $nextWaterMark = $nextResult.WaterMark
        $cookieChanged = ($nextWaterMark -ne $newWaterMark)

        if ($cookieChanged -and $nextHasContent) {
            # Normal success: new cookie with content - continue
            $result = $nextResult
        }
        elseif ($cookieChanged -and -not $nextHasContent) {
            # New cookie but no content: retry previous cookie once
            $msg = "  Page {0}: New PageCookie but no content - retrying previous cookie (30s wait)" -f ($pageNumber + 1)
            Write-ExportLog -Message $msg -Level Warning
            Write-ProgressEntry -Path $ProgressLogPath -Message $msg

            Start-Sleep -Seconds 30

            # Retry with the previous (working) cookie
            $retryParams = $exportParams.Clone()
            $retryParams['PageCookie'] = $newWaterMark

            try {
                $retryResult = Export-ActivityExplorerData @retryParams
                $retryHasContent = Test-PageHasContent -Result $retryResult

                if ($retryHasContent) {
                    # Previous cookie now has content
                    $result = $retryResult
                }
                else {
                    # Still no content - continue with the new cookie regardless
                    $msg = "  Page {0}: Retry also returned no content - continuing with new cookie" -f ($pageNumber + 1)
                    Write-ExportLog -Message $msg -Level Warning
                    Write-ProgressEntry -Path $ProgressLogPath -Message $msg
                    $result = $nextResult
                }
            }
            catch {
                # Retry failed - continue with new cookie
                $msg = "  Page {0}: Previous cookie retry failed - continuing with new cookie" -f ($pageNumber + 1)
                Write-ExportLog -Message $msg -Level Warning
                $result = $nextResult
            }
        }
        elseif (-not $cookieChanged) {
            # Same cookie returned - retry with delays
            $sameCookieRetries++

            if ($sameCookieRetries -gt $maxRetries) {
                # Same cookie exhausted all retries
                $msg = "  Page {0}: Same PageCookie returned {1} times - saving progress" -f ($pageNumber + 1), $sameCookieRetries
                Write-ExportLog -Message $msg -Level Error
                Write-ProgressEntry -Path $ProgressLogPath -Message $msg

                Add-PartialError -Tracker $Tracker -PageNumber ($pageNumber + 1) `
                    -ErrorMessage "Same PageCookie returned after $sameCookieRetries retries" `
                    -ErrorType "SameCookieRetryExhausted"

                if ($sameCookieRetries -eq $maxRetries + 1) {
                    # One final attempt after 120s
                    $msg = "  Page {0}: Final attempt after 120s wait" -f ($pageNumber + 1)
                    Write-ExportLog -Message $msg -Level Warning
                    Write-ProgressEntry -Path $ProgressLogPath -Message $msg
                    Start-Sleep -Seconds 120

                    try {
                        $finalResult = Export-ActivityExplorerData @exportParams
                        $finalCookie = $finalResult.WaterMark
                        if ($finalCookie -ne $newWaterMark -or (Test-PageHasContent -Result $finalResult)) {
                            $result = $finalResult
                            $sameCookieRetries = 0
                            continue
                        }
                    }
                    catch {
                        # Final attempt also failed - will fall through to PartialFailure handling below
                        Write-ExportLog -Message ("  Page {0}: Final attempt also failed: {1}" -f ($pageNumber + 1), $_.Exception.Message) -Level Warning
                    }
                }

                $Tracker['PartialFailure'] = $true
                $Tracker['Status'] = "PartialFailure"
                Save-ActivityExplorerRunTracker -Tracker $Tracker -TrackerPath $TrackerPath

                return @{
                    TotalRecords = $totalRecords
                    PageCount    = $pageNumber
                    ResumedFrom  = $resumedFrom
                }
            }

            # Retry with 60s delay
            $delay = 60
            $msg = "  Page {0}: Same PageCookie (retry {1}/{2}) - waiting {3}s" -f ($pageNumber + 1), $sameCookieRetries, $maxRetries, $delay
            Write-ExportLog -Message $msg -Level Warning
            Write-ProgressEntry -Path $ProgressLogPath -Message $msg
            Start-Sleep -Seconds $delay

            try {
                $retryResult = Export-ActivityExplorerData @exportParams
                $result = $retryResult
                # Reset same-cookie counter if cookie changed
                if ($retryResult.WaterMark -ne $newWaterMark) {
                    $sameCookieRetries = 0
                }
            }
            catch {
                $errorMsg = Get-PageErrorMessage -ErrorRecord $_
                Add-PartialError -Tracker $Tracker -PageNumber ($pageNumber + 1) -ErrorMessage $errorMsg -ErrorType "SameCookieRetryError"

                # Auth error - throw
                $sameCookieErrorInfo = Get-HttpErrorExplanation -ErrorMessage $errorMsg -ErrorRecord $_
                if ($sameCookieErrorInfo.Category -eq "AuthError") {
                    throw
                }

                # Continue retry loop
                $result = $nextResult
            }
        }
        else {
            # Fallback: use whatever we got
            $result = $nextResult
        }
    }

    # Export completed successfully
    $totalElapsed = (Get-Date) - $exportStartTime
    $Tracker['Status'] = "Completed"
    $Tracker['EndTime'] = (Get-Date).ToString("o")
    $Tracker['Duration'] = $totalElapsed.ToString()
    Save-ActivityExplorerRunTracker -Tracker $Tracker -TrackerPath $TrackerPath

    $completionMsg = "  Export completed: {0} records in {1} pages ({2:N1} minutes)" -f $totalRecords, $pageNumber, $totalElapsed.TotalMinutes
    Write-ExportLog -Message $completionMsg -Level Success
    Write-ProgressEntry -Path $ProgressLogPath -Message $completionMsg

    return @{
        TotalRecords = $totalRecords
        PageCount    = $pageNumber
        ResumedFrom  = $resumedFrom
    }
}

function Merge-ActivityExplorerPages {
    <#
    .SYNOPSIS
        Merges and deduplicates Activity Explorer page files from multi-terminal worker directories.
    .DESCRIPTION
        Scans Data/ActivityExplorer/YYYY-MM-DD/Page-*.json files across all day directories,
        deduplicates records by RecordIdentity, and writes a combined output file.
        Uses streaming JSON write for large datasets (>50k records) to minimize memory.
    .PARAMETER ExportDirectory
        Root export directory containing Data/ActivityExplorer/ subdirectories.
    .PARAMETER OutputPath
        Path for the combined output file. Defaults to ExportDirectory/Data/ActivityExplorer/ActivityExplorer-Combined.json.
    .PARAMETER StreamingOutput
        Force streaming output mode regardless of record count.
    .OUTPUTS
        Hashtable with TotalRecords, UniqueRecords, DuplicatesRemoved, PagesProcessed.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ExportDirectory,

        [string]$OutputPath,

        [switch]$StreamingOutput
    )

    if (-not $OutputPath) {
        $OutputPath = Join-Path (Get-AEDataDir $ExportDirectory) "ActivityExplorer-Combined.json"
    }

    # Find all page files in Data/ActivityExplorer/ day directories
    $aeDataDir = Get-AEDataDir $ExportDirectory
    $pageFiles = @(Get-ChildItem -Path $aeDataDir -Recurse -Filter "Page-*.json" -ErrorAction SilentlyContinue |
        Sort-Object FullName)

    if ($pageFiles.Count -eq 0) {
        Write-Verbose "No AE page files found in worker directories"
        return @{
            TotalRecords      = 0
            UniqueRecords     = 0
            DuplicatesRemoved = 0
            PagesProcessed    = 0
        }
    }

    Write-Verbose ("Found {0} page file(s) across worker directories" -f $pageFiles.Count)

    # First pass: collect all records and dedup by RecordIdentity
    $seenIdentities = @{}
    $allRecords = [System.Collections.ArrayList]::new()
    $totalRecords = 0
    $pagesProcessed = 0

    foreach ($pageFile in $pageFiles) {
        try {
            $pageData = Get-Content -Raw -Path $pageFile.FullName -ErrorAction Stop | ConvertFrom-Json
            if ($null -eq $pageData) {
                Write-Warning ("Page file {0} parsed as null, skipping" -f $pageFile.Name)
                continue
            }
            if ($pageData.Records) {
                foreach ($record in $pageData.Records) {
                    $totalRecords++
                    $identity = $record.RecordIdentity
                    if ($identity -and $seenIdentities.ContainsKey($identity)) {
                        continue  # Duplicate
                    }
                    if ($identity) {
                        $seenIdentities[$identity] = $true
                    }
                    [void]$allRecords.Add($record)
                }
            }
            $pagesProcessed++
        }
        catch {
            Write-Warning ("Failed to read page file {0}: {1}" -f $pageFile.Name, $_.Exception.Message)
        }
    }

    $uniqueRecords = $allRecords.Count
    $duplicatesRemoved = $totalRecords - $uniqueRecords

    Write-Verbose ("Total: {0}, Unique: {1}, Duplicates removed: {2}" -f $totalRecords, $uniqueRecords, $duplicatesRemoved)

    # Write output
    $useStreaming = $StreamingOutput -or ($uniqueRecords -gt 50000)

    if ($useStreaming) {
        Write-Verbose "Writing combined file (streaming mode)..."
        $stream = $null
        $writer = $null
        try {
            $stream = [System.IO.FileStream]::new($OutputPath, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None, 65536)
            $writer = [System.IO.StreamWriter]::new($stream, [System.Text.Encoding]::UTF8, 65536)

            $writer.WriteLine("[")
            $isFirst = $true
            $written = 0

            foreach ($record in $allRecords) {
                if (-not $isFirst) {
                    $writer.WriteLine(",")
                }
                $json = $record | ConvertTo-Json -Depth 20 -Compress
                $writer.Write("  ")
                $writer.Write($json)
                $isFirst = $false
                $written++

                if ($written % 10000 -eq 0) {
                    Write-Verbose "  Written $written records..."
                }
            }

            $writer.WriteLine("")
            $writer.WriteLine("]")
        }
        finally {
            if ($writer) { $writer.Dispose() }
            if ($stream) { $stream.Dispose() }
        }
    }
    else {
        Write-Verbose "Writing combined file (standard mode)..."
        if ($uniqueRecords -gt 0) {
            $json = $allRecords.ToArray() | ConvertTo-Json -Depth 20
            [System.IO.File]::WriteAllText($OutputPath, $json, [System.Text.Encoding]::UTF8)
        }
        else {
            [System.IO.File]::WriteAllText($OutputPath, "[]", [System.Text.Encoding]::UTF8)
        }
    }

    return @{
        TotalRecords      = $totalRecords
        UniqueRecords     = $uniqueRecords
        DuplicatesRemoved = $duplicatesRemoved
        PagesProcessed    = $pagesProcessed
    }
}

function Find-UnknownActivityTypes {
    <#
    .SYNOPSIS
        Detects activity types or workloads in exported data that aren't in the config.
    .DESCRIPTION
        Compares unique Activity and Workload values found in exported records against
        the known lists from the configuration file. Logs any unknown values as warnings
        so the config can be updated for future exports.
    .PARAMETER Records
        Array of Activity Explorer records to analyze.
    .PARAMETER KnownActivities
        Array of activity type names from the configuration file.
    .PARAMETER KnownWorkloads
        Array of workload names from the configuration file.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [array]$Records,

        [string[]]$KnownActivities,

        [string[]]$KnownWorkloads
    )

    if ($Records.Count -eq 0) {
        return
    }

    # Extract unique activity types from records
    $recordActivities = @($Records | Where-Object { $_.Activity } | ForEach-Object { $_.Activity } | Select-Object -Unique)
    $recordWorkloads = @($Records | Where-Object { $_.Workload } | ForEach-Object { $_.Workload } | Select-Object -Unique)

    $foundUnknown = $false

    # Check for unknown activities
    if ($KnownActivities -and $KnownActivities.Count -gt 0 -and $recordActivities.Count -gt 0) {
        $unknownActivities = @($recordActivities | Where-Object { $_ -notin $KnownActivities })

        if ($unknownActivities.Count -gt 0) {
            $foundUnknown = $true
            $activityList = $unknownActivities -join ", "
            $msg = "  UNKNOWN ACTIVITY TYPES found in export ({0}): {1}" -f $unknownActivities.Count, $activityList
            Write-ExportLog -Message $msg -Level Warning
            Write-ExportLog -Message "  Consider adding these to ConfigFiles\ActivityExplorerSelector.json" -Level Warning
        }
    }

    # Check for unknown workloads
    if ($KnownWorkloads -and $KnownWorkloads.Count -gt 0 -and $recordWorkloads.Count -gt 0) {
        $unknownWorkloads = @($recordWorkloads | Where-Object { $_ -notin $KnownWorkloads })

        if ($unknownWorkloads.Count -gt 0) {
            $foundUnknown = $true
            $workloadList = $unknownWorkloads -join ", "
            $msg = "  UNKNOWN WORKLOADS found in export ({0}): {1}" -f $unknownWorkloads.Count, $workloadList
            Write-ExportLog -Message $msg -Level Warning
            Write-ExportLog -Message "  Consider adding these to ConfigFiles\ActivityExplorerSelector.json" -Level Warning
        }
    }

    if (-not $foundUnknown) {
        Write-ExportLog -Message "  All activity types and workloads match configuration" -Level Info
    }
}

function Get-ActivityExplorerFilters {
    <#
    .SYNOPSIS
        Loads Activity Explorer filter configuration from ActivityExplorerSelector.json.
    .DESCRIPTION
        Reads the config file, extracts enabled activities and workloads, and returns
        a hashtable suitable for passing to Export-ActivityExplorerWithProgress.
        Only adds a filter category if some (but not all) items are disabled.
        Returns $null if no filters are needed (all items enabled or no config found).
    .PARAMETER ConfigPath
        Path to the ActivityExplorerSelector.json config file.
    .PARAMETER ConfigObject
        Pre-loaded config object (e.g., from ExportSettings.json manifest).
        When provided, skips file read and uses this object directly.
    .PARAMETER LogDetails
        When set, logs details about which filters are active using Write-ExportLog.
    #>
    [CmdletBinding(DefaultParameterSetName = 'ByPath')]
    param(
        [Parameter(ParameterSetName = 'ByPath')]
        [string]$ConfigPath,

        [Parameter(Mandatory, ParameterSetName = 'ByObject')]
        [PSCustomObject]$ConfigObject,

        [switch]$LogDetails
    )

    if ($PSCmdlet.ParameterSetName -eq 'ByObject') {
        $config = $ConfigObject
    }
    elseif ($ConfigPath) {
        $config = Read-JsonConfig -Path $ConfigPath
    }
    else {
        $config = $null
    }

    if (-not $config) {
        if ($LogDetails) {
            Write-ExportLog -Message "  No config file found, exporting all activities" -Level Info
        }
        return $null
    }

    $filters = @{}

    # Only add activity filter if some (but not all) activities are disabled
    if ($config.Activities) {
        $enabledActivities = Get-EnabledItems -Config $config.Activities
        $totalActivities = @($config.Activities.PSObject.Properties | Where-Object { $_.Name -notlike "_*" }).Count
        if ($enabledActivities.Count -gt 0 -and $enabledActivities.Count -lt $totalActivities) {
            $filters['Activity'] = $enabledActivities
            if ($LogDetails) {
                Write-ExportLog -Message "  Filtering to $($enabledActivities.Count) of $totalActivities activities" -Level Info
            }
        }
        elseif ($LogDetails) {
            Write-ExportLog -Message "  All $totalActivities activities enabled (no activity filter)" -Level Info
        }
    }

    # Only add workload filter if some (but not all) workloads are disabled
    if ($config.Workloads) {
        $enabledWorkloads = Get-EnabledItems -Config $config.Workloads
        $totalWorkloads = @($config.Workloads.PSObject.Properties | Where-Object { $_.Name -notlike "_*" }).Count
        if ($enabledWorkloads.Count -gt 0 -and $enabledWorkloads.Count -lt $totalWorkloads) {
            $filters['Workload'] = $enabledWorkloads
            if ($LogDetails) {
                Write-ExportLog -Message "  Filtering to $($enabledWorkloads.Count) of $totalWorkloads workloads" -Level Info
            }
        }
        elseif ($LogDetails) {
            Write-ExportLog -Message "  All $totalWorkloads workloads enabled (no workload filter)" -Level Info
        }
    }

    if ($filters.Count -eq 0) {
        if ($LogDetails) {
            Write-ExportLog -Message "  No filters applied (exporting all activities)" -Level Info
        }
        return $null
    }

    return $filters
}

#endregion

#region Private Helper

function Write-ProgressEntry {
    <#
    .SYNOPSIS
        Writes a timestamped entry to the progress log file.
    .DESCRIPTION
        Appends a line to the progress log with a timestamp prefix.
        Silently skips if no path is provided or the write fails.
    .PARAMETER Path
        Path to the progress log file. Also accepts -LogPath.
    .PARAMETER Message
        The message to write.
    #>
    [CmdletBinding()]
    param(
        [Alias('LogPath')]
        [string]$Path,
        [string]$Message
    )

    if ([string]::IsNullOrEmpty($Path)) {
        return
    }

    try {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
        $entry = "[{0}] {1}" -f $timestamp, $Message
        [System.IO.File]::AppendAllText($Path, $entry + [Environment]::NewLine, [System.Text.Encoding]::UTF8)
    }
    catch {
        # Progress log write failures are non-critical - export continues without progress logging
        Write-Verbose "Progress log write failed (non-critical): $($_.Exception.Message)"
    }
}

#endregion

#region Content Explorer - Retry Bucket

function Get-RetryBucketTasks {
    <#
    .SYNOPSIS
        Identifies completed detail tasks where actual exported count differs from expected by more than a threshold.
    .DESCRIPTION
        Iterates completed detail tasks from a DetailTasks.csv and returns those where the discrepancy
        between OriginalExpectedCount and the actual exported count (stored in ExpectedCount after overwrite)
        exceeds the specified threshold. Skips tasks with errors, missing data, or very small counts.
    .PARAMETER DetailTasks
        Array of task objects from Read-TaskCsv (DetailTasks.csv).
    .PARAMETER Threshold
        Fractional threshold for discrepancy detection. Default 0.02 (2%).
    .PARAMETER MinCount
        Minimum OriginalExpectedCount to consider. Tasks below this are excluded
        because percentage-based thresholds are meaningless for tiny counts. Default 10.
    .OUTPUTS
        Array of PSCustomObjects with: TagType, TagName, Workload, OriginalExpectedCount, ActualCount, DiscrepancyPct, PageSize
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [array]$DetailTasks,

        [double]$Threshold = 0.02,

        [int]$MinCount = 10
    )

    $retryTasks = @()

    foreach ($task in $DetailTasks) {
        # Only check completed tasks
        if ($task.Status -ne "Completed") { continue }

        # Skip tasks with error messages (aggregate errors, etc.)
        if ($task.ErrorMessage -and $task.ErrorMessage.Trim() -ne "") { continue }

        # Get original expected count - skip if column missing (backward compat) or zero
        $originalExpected = $task.OriginalExpectedCount -as [int]
        if (-not $originalExpected -or $originalExpected -eq 0) { continue }

        # Skip small tasks where percentage threshold is meaningless
        if ($originalExpected -lt $MinCount) { continue }

        # ActualCount is stored in ExpectedCount after the orchestrator overwrites it
        $actualCount = $task.ExpectedCount -as [int]
        if ($null -eq $actualCount) { $actualCount = 0 }

        # Compute discrepancy
        $discrepancy = [Math]::Abs($actualCount - $originalExpected) / $originalExpected

        if ($discrepancy -gt $Threshold) {
            $discrepancyPct = [Math]::Round(($actualCount - $originalExpected) / $originalExpected * 100, 1)
            $retryTasks += [PSCustomObject]@{
                TagType               = $task.TagType
                TagName               = $task.TagName
                Workload              = $task.Workload
                OriginalExpectedCount = $originalExpected
                ActualCount           = $actualCount
                DiscrepancyPct        = $discrepancyPct
                PageSize              = ($task.PageSize -as [int])
            }
        }
    }

    return $retryTasks
}

#endregion

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

#region Phase/Type I/O

function Write-ExportPhase {
    <#
    .SYNOPSIS
        Writes the current export phase to ExportPhase.txt (atomic write).
    #>
    param(
        [Parameter(Mandatory)][string]$ExportDir,
        [Parameter(Mandatory)][ValidateSet('Aggregate','Detail','Completed','AEExport','AECompleted')][string]$Phase
    )
    $coordDir = Get-CoordinationDir $ExportDir
    if (-not (Test-Path $coordDir)) { New-Item -ItemType Directory -Force -Path $coordDir | Out-Null }
    $phasePath = Join-Path $coordDir "ExportPhase.txt"
    $tmpPath = $phasePath + ".tmp.$PID"
    [System.IO.File]::WriteAllText($tmpPath, $Phase)
    [System.IO.File]::Move($tmpPath, $phasePath, $true)
}

function Read-ExportPhase {
    <#
    .SYNOPSIS
        Reads the current export phase from ExportPhase.txt.
    #>
    param([Parameter(Mandatory)][string]$ExportDir)
    $phasePath = Join-Path (Get-CoordinationDir $ExportDir) "ExportPhase.txt"
    if (Test-Path $phasePath) {
        return ([System.IO.File]::ReadAllText($phasePath)).Trim()
    }
    return $null
}

function Write-ExportType {
    <#
    .SYNOPSIS
        Writes the export type marker to ExportType.txt (atomic write).
    .DESCRIPTION
        Used to identify whether an export directory contains a Content Explorer
        or Activity Explorer multi-terminal export. Workers read this at startup
        to determine which worker function to invoke.
    #>
    param(
        [Parameter(Mandatory)][string]$ExportDir,
        [Parameter(Mandatory)][ValidateSet('ContentExplorer','ActivityExplorer')][string]$Type
    )
    $coordDir = Get-CoordinationDir $ExportDir
    if (-not (Test-Path $coordDir)) { New-Item -ItemType Directory -Force -Path $coordDir | Out-Null }
    $typePath = Join-Path $coordDir "ExportType.txt"
    $tmpPath = $typePath + ".tmp.$PID"
    [System.IO.File]::WriteAllText($tmpPath, $Type)
    [System.IO.File]::Move($tmpPath, $typePath, $true)
}

function Read-ExportType {
    <#
    .SYNOPSIS
        Reads the export type from ExportType.txt.
    .OUTPUTS
        'ContentExplorer', 'ActivityExplorer', or $null if file doesn't exist.
    #>
    param([Parameter(Mandatory)][string]$ExportDir)
    $typePath = Join-Path (Get-CoordinationDir $ExportDir) "ExportType.txt"
    if (Test-Path $typePath) {
        return ([System.IO.File]::ReadAllText($typePath)).Trim()
    }
    return $null
}

#endregion

#region Task CSV I/O

function Write-AETaskCsv {
    <#
    .SYNOPSIS
        Writes an Activity Explorer day task CSV file (AEDayTasks.csv) atomically.
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][array]$Tasks
    )
    $tmpPath = $Path + ".tmp.$PID"
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine("Day,StartTime,EndTime,AssignedPID,Status,PageCount,RecordCount,ErrorMessage")
    foreach ($task in $Tasks) {
        $escapedErr = if ($task.ErrorMessage) { ($task.ErrorMessage -replace '"','""') } else { "" }
        $line = '{0},{1},{2},{3},{4},{5},{6},"{7}"' -f $task.Day, $task.StartTime, $task.EndTime, ($task.AssignedPID -as [int]), $task.Status, ($task.PageCount -as [int]), ($task.RecordCount -as [int]), $escapedErr
        [void]$sb.AppendLine($line)
    }
    [System.IO.File]::WriteAllText($tmpPath, $sb.ToString(), [System.Text.Encoding]::UTF8)
    [System.IO.File]::Move($tmpPath, $Path, $true)
}

function Read-AETaskCsv {
    <#
    .SYNOPSIS
        Reads an Activity Explorer day task CSV file and returns task objects.
    #>
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path $Path)) { return @() }
    try {
        return @(Import-Csv -Path $Path -Encoding UTF8)
    }
    catch {
        return @()
    }
}

function Write-TaskCsv {
    <#
    .SYNOPSIS
        Writes a task CSV file (AggregateTasks.csv or DetailTasks.csv) atomically.
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][array]$Tasks
    )
    $tmpPath = $Path + ".tmp.$PID"
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine("TagType,TagName,Workload,Location,LocationType,ExpectedCount,PageSize,AssignedPID,Status,ErrorMessage,OriginalExpectedCount")
    foreach ($task in $Tasks) {
        $escapedTag = $task.TagName -replace '"','""'
        $escapedLoc = if ($task.Location) { ($task.Location -replace '"','""') } else { "" }
        $locType = if ($task.LocationType) { $task.LocationType } else { "" }
        $escapedErr = if ($task.ErrorMessage) { ($task.ErrorMessage -replace '"','""') } else { "" }
        $origExpected = if ($task.OriginalExpectedCount) { $task.OriginalExpectedCount -as [int] } else { $task.ExpectedCount -as [int] }
        $line = '{0},"{1}",{2},"{3}",{4},{5},{6},{7},{8},"{9}",{10}' -f $task.TagType, $escapedTag, $task.Workload, $escapedLoc, $locType, ($task.ExpectedCount -as [int]), ($task.PageSize -as [int]), ($task.AssignedPID -as [int]), $task.Status, $escapedErr, $origExpected
        [void]$sb.AppendLine($line)
    }
    [System.IO.File]::WriteAllText($tmpPath, $sb.ToString(), [System.Text.Encoding]::UTF8)
    [System.IO.File]::Move($tmpPath, $Path, $true)
}

function Read-TaskCsv {
    <#
    .SYNOPSIS
        Reads a task CSV file and returns task objects.
    #>
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path $Path)) { return @() }
    try {
        return @(Import-Csv -Path $Path -Encoding UTF8)
    }
    catch {
        return @()
    }
}

function Update-DetailTaskPageSizes {
    <#
    .SYNOPSIS
        Recalculates PageSize for pending location-level tasks in an existing DetailTasks.csv.
    .DESCRIPTION
        Updates PageSize for tasks that haven't started yet, using single-location sizing:
          <100 items -> 100, <1000 -> 500, <10000 -> 2000, else -> 5000.
        WorkloadFallback tasks are left unchanged. Already completed/in-progress tasks are untouched.
    .PARAMETER Path
        Path to the DetailTasks.csv file.
    .PARAMETER WhatIf
        If set, shows what would change without writing.
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [switch]$WhatIf
    )

    if (-not (Test-Path $Path)) {
        Write-Host "DetailTasks.csv not found: $Path" -ForegroundColor Red
        return
    }

    $tasks = @(Import-Csv -Path $Path -Encoding UTF8)
    $updated = 0
    $skipped = 0

    foreach ($task in $tasks) {
        # Only update pending tasks with a location (not WorkloadFallback, not started)
        if ($task.Status -ne "Pending") { $skipped++; continue }
        if (-not $task.Location -or $task.LocationType -eq "WorkloadFallback") { $skipped++; continue }

        $expected = [int]$task.ExpectedCount
        $oldPageSize = [int]$task.PageSize

        # Single-location page sizing: no multi-folder overhead
        $newPageSize = if ($expected -lt 100) { 100 }
                       elseif ($expected -lt 1000) { 500 }
                       elseif ($expected -lt 10000) { 2000 }
                       else { 5000 }

        if ($newPageSize -ne $oldPageSize) {
            if ($WhatIf) {
                Write-Host ("  {0}/{1} [{2}]: {3} -> {4} (expected: {5:N0})" -f $task.TagName, $task.Workload, $task.Location.Substring(0, [Math]::Min(40, $task.Location.Length)), $oldPageSize, $newPageSize, $expected)
            }
            $task.PageSize = $newPageSize
            $updated++
        }
    }

    if ($WhatIf) {
        Write-Host "`n$updated tasks would be updated, $skipped skipped (completed/in-progress/fallback)" -ForegroundColor Cyan
    } else {
        Write-TaskCsv -Path $Path -Tasks $tasks
        Write-Host "$updated tasks updated, $skipped skipped" -ForegroundColor Green
    }
}

function Write-RetryTasksCsv {
    <#
    .SYNOPSIS
        Writes a RetryTasks.csv file from retry bucket task objects.
    .PARAMETER Path
        Output file path.
    .PARAMETER RetryTasks
        Array of retry task objects from Get-RetryBucketTasks.
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][array]$RetryTasks
    )
    $tmpPath = $Path + ".tmp.$PID"
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine("TagType,TagName,Workload,OriginalExpectedCount,ActualCount,DiscrepancyPct,PageSize")
    foreach ($task in $RetryTasks) {
        $escapedTag = ($task.TagName -replace '"','""')
        $line = '{0},"{1}",{2},{3},{4},{5},{6}' -f $task.TagType, $escapedTag, $task.Workload, ($task.OriginalExpectedCount -as [int]), ($task.ActualCount -as [int]), $task.DiscrepancyPct, ($task.PageSize -as [int])
        [void]$sb.AppendLine($line)
    }
    [System.IO.File]::WriteAllText($tmpPath, $sb.ToString(), [System.Text.Encoding]::UTF8)
    [System.IO.File]::Move($tmpPath, $Path, $true)
}

function Show-RetryBucketSummary {
    <#
    .SYNOPSIS
        Displays the retry bucket summary after combination phase.
    .PARAMETER RetryTasks
        Array of retry task objects from Get-RetryBucketTasks.
    .PARAMETER ExportDir
        Path to the export directory (for displaying the retry command).
    #>
    param(
        [array]$RetryTasks,
        [string]$ExportDir
    )
    Write-ExportLog -Message "`n--- Retry Bucket ---" -Level Info
    if ($RetryTasks -and $RetryTasks.Count -gt 0) {
        Write-ExportLog -Message ("  Tasks with >2%% discrepancy: {0}" -f $RetryTasks.Count) -Level Warning
        foreach ($rt in $RetryTasks) {
            $sign = if ($rt.DiscrepancyPct -ge 0) { "+" } else { "" }
            Write-ExportLog -Message ("    {0} / {1}: expected {2}, got {3} ({4}{5}%%)" -f $rt.TagName, $rt.Workload, $rt.OriginalExpectedCount.ToString('N0'), $rt.ActualCount.ToString('N0'), $sign, $rt.DiscrepancyPct) -Level Warning
        }
        Write-ExportLog -Message "  Retry file: RetryTasks.csv" -Level Info
        Write-ExportLog -Message ("  To retry: .\Export-Compl8Configuration.ps1 -CERetryDir ""{0}""" -f $ExportDir) -Level Info
    }
    else {
        Write-ExportLog -Message "  All tasks within 2%% tolerance" -Level Success
    }
}

function Write-RemainingTasksCsv {
    <#
    .SYNOPSIS
        Writes non-completed detail tasks to RemainingTasks.csv for follow-on runs.
    .PARAMETER ExportDir
        Path to the export run directory.
    .OUTPUTS
        Int - count of remaining tasks written (0 if all completed).
    #>
    param(
        [Parameter(Mandatory)][string]$ExportDir
    )
    $detailCsvPath = Join-Path (Get-CoordinationDir $ExportDir) "DetailTasks.csv"
    if (-not (Test-Path $detailCsvPath)) { return 0 }

    $allTasks = Read-TaskCsv -Path $detailCsvPath
    $remaining = @($allTasks | Where-Object { $_.Status -ne "Completed" })
    if ($remaining.Count -eq 0) { return 0 }

    $remainingPath = Join-Path (Get-CoordinationDir $ExportDir) "RemainingTasks.csv"
    Write-TaskCsv -Path $remainingPath -Tasks $remaining
    return $remaining.Count
}

#endregion

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

#region Text UI Helpers

# ANSI color map for dashboard rendering (matches PowerShell ConsoleColor names)
$script:AnsiColorMap = @{
    'Black'       = "`e[30m"
    'DarkRed'     = "`e[31m"
    'DarkGreen'   = "`e[32m"
    'DarkYellow'  = "`e[33m"
    'DarkBlue'    = "`e[34m"
    'DarkMagenta' = "`e[35m"
    'DarkCyan'    = "`e[36m"
    'Gray'        = "`e[37m"
    'DarkGray'    = "`e[90m"
    'Red'         = "`e[91m"
    'Green'       = "`e[92m"
    'Yellow'      = "`e[93m"
    'Blue'        = "`e[94m"
    'Magenta'     = "`e[95m"
    'Cyan'        = "`e[96m"
    'White'       = "`e[97m"
}
$script:AnsiReset = "`e[0m"

function Get-TerminalSize {
    <#
    .SYNOPSIS
        Returns the current terminal viewport dimensions.
    .DESCRIPTION
        Single source of truth for terminal width/height. Reads on every call
        (no caching) so resizes are picked up. Clamps to sensible minimums and
        subtracts 1 from width to avoid last-column line-wrap artifacts.
    #>
    try {
        @{
            Width  = [Math]::Max(40, [Console]::WindowWidth - 1)
            Height = [Math]::Max(10, [Console]::WindowHeight)
        }
    }
    catch {
        @{ Width = 119; Height = 30 }
    }
}

function Format-ProgressBar {
    <#
    .SYNOPSIS
        Returns a width-proportional progress bar string like [####........].
    .PARAMETER Percent
        Completion percentage (0-100).
    .PARAMETER MaxWidth
        Maximum character width for the bar including brackets. Defaults to
        a fraction of terminal width (clamped between 20 and 50).
    #>
    [CmdletBinding()]
    param(
        [double]$Percent = 0,
        [int]$MaxWidth = 0
    )

    if ($MaxWidth -le 0) {
        $termWidth = (Get-TerminalSize).Width
        $MaxWidth = [Math]::Min(50, [Math]::Max(20, [int]($termWidth * 0.35)))
    }

    $inner = $MaxWidth - 2  # subtract brackets
    $filled = [Math]::Min($inner, [int]($Percent / 100 * $inner))
    $empty = $inner - $filled
    return "[" + ("#" * $filled) + ("." * $empty) + "]"
}

function Write-BoxTop {
    <#
    .SYNOPSIS
        Writes the top border of a box: ╔═══════╗ or ┌───────┐
    #>
    param([int]$InnerWidth, [string]$Color = 'Cyan', [switch]$Double)
    $h = if ($Double) { [char]0x2550 } else { [char]0x2500 }  # ═ or ─
    $tl = if ($Double) { [char]0x2554 } else { [char]0x250C }  # ╔ or ┌
    $tr = if ($Double) { [char]0x2557 } else { [char]0x2510 }  # ╗ or ┐
    Write-Host ("  {0}{1}{2}" -f $tl, ([string]$h * $InnerWidth), $tr) -ForegroundColor $Color
}

function Write-BoxBottom {
    <#
    .SYNOPSIS
        Writes the bottom border of a box: ╚═══════╝ or └───────┘
    #>
    param([int]$InnerWidth, [string]$Color = 'Cyan', [switch]$Double)
    $h = if ($Double) { [char]0x2550 } else { [char]0x2500 }
    $bl = if ($Double) { [char]0x255A } else { [char]0x2514 }  # ╚ or └
    $br = if ($Double) { [char]0x255D } else { [char]0x2518 }  # ╝ or ┘
    Write-Host ("  {0}{1}{2}" -f $bl, ([string]$h * $InnerWidth), $br) -ForegroundColor $Color
}

function Write-BoxLine {
    <#
    .SYNOPSIS
        Writes a content line inside a box: ║ text ║ or │ text │
    .DESCRIPTION
        Pads or truncates text to fit InnerWidth. Text longer than available
        space is truncated with an ellipsis character.
    #>
    param(
        [string]$Text = '',
        [int]$InnerWidth,
        [string]$Color = 'Cyan',
        [switch]$Double
    )
    $v = if ($Double) { [char]0x2551 } else { [char]0x2502 }  # ║ or │
    $contentWidth = $InnerWidth - 2  # 1 space padding each side
    if ($Text.Length -gt $contentWidth) {
        $Text = $Text.Substring(0, $contentWidth - 1) + [char]0x2026  # …
    }
    $padded = $Text.PadRight($contentWidth)
    Write-Host ("  {0} {1} {2}" -f $v, $padded, $v) -ForegroundColor $Color
}

function Write-BoxSeparator {
    <#
    .SYNOPSIS
        Writes a separator line inside a box: ╠═══════╣ or ├───────┤
    #>
    param([int]$InnerWidth, [string]$Color = 'Cyan', [switch]$Double)
    $h = if ($Double) { [char]0x2550 } else { [char]0x2500 }
    $ml = if ($Double) { [char]0x2560 } else { [char]0x251C }  # ╠ or ├
    $mr = if ($Double) { [char]0x2563 } else { [char]0x2524 }  # ╣ or ┤
    Write-Host ("  {0}{1}{2}" -f $ml, ([string]$h * $InnerWidth), $mr) -ForegroundColor $Color
}

function Get-BoxInnerWidth {
    <#
    .SYNOPSIS
        Computes the inner width for a box based on terminal width.
    .PARAMETER MaxWidth
        Maximum inner width. The result is clamped to (terminal width - 6)
        to leave room for the 2-char indent and the border characters.
    #>
    param([int]$MaxWidth = 62)
    $termWidth = (Get-TerminalSize).Width
    return [Math]::Min($MaxWidth, $termWidth - 6)
}

function Write-SectionHeader {
    <#
    .SYNOPSIS
        Writes an adaptive-width section header: ── Text ──────────
    .PARAMETER Text
        The header text.
    .PARAMETER Color
        Foreground color. Default: DarkCyan.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Text,
        [string]$Color = 'DarkCyan'
    )
    $termWidth = (Get-TerminalSize).Width
    $h = [char]0x2500  # ─
    $prefix = "  ${h}${h} "
    $suffix = " "
    $remaining = $termWidth - $prefix.Length - $Text.Length - $suffix.Length - 2
    if ($remaining -lt 2) { $remaining = 2 }
    $trail = [string]$h * $remaining
    Write-Host "${prefix}${Text}${suffix}${trail}" -ForegroundColor $Color
}

function Write-Banner {
    <#
    .SYNOPSIS
        Writes a bordered banner box with a title and optional content lines.
    .DESCRIPTION
        Adapts to terminal width. Uses double-line borders for emphasis.
        All content is padded/truncated to ensure the right border is always present.
    .PARAMETER Title
        The banner title (displayed on first content line).
    .PARAMETER Lines
        Optional additional content lines.
    .PARAMETER Color
        Border and text color. Default: Cyan.
    .PARAMETER Double
        Use double-line border characters. Default: True.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Title,
        [string[]]$Lines = @(),
        [string]$Color = 'Cyan',
        [switch]$Single
    )

    $double = -not $Single
    $innerWidth = Get-BoxInnerWidth -MaxWidth 62
    Write-BoxTop -InnerWidth $innerWidth -Color $Color -Double:$double
    Write-BoxLine -Text $Title -InnerWidth $innerWidth -Color $Color -Double:$double
    if ($Lines.Count -gt 0) {
        Write-BoxSeparator -InnerWidth $innerWidth -Color $Color -Double:$double
        foreach ($line in $Lines) {
            Write-BoxLine -Text $line -InnerWidth $innerWidth -Color $Color -Double:$double
        }
    }
    Write-BoxBottom -InnerWidth $innerWidth -Color $Color -Double:$double
}

function Write-DashboardFrame {
    <#
    .SYNOPSIS
        Renders a dashboard frame in-place using ANSI escape sequences.
    .DESCRIPTION
        Builds the entire frame in a StringBuilder and emits it with a single
        [Console]::Write() call for flicker-free rendering. Uses ANSI cursor-up
        (relative positioning) instead of absolute SetCursorPosition to avoid
        scroll drift. Hides the cursor during redraw.

        Each line in $Lines should be a hashtable with Text and Color keys.
        The Color value should be a PowerShell ConsoleColor name.
    .PARAMETER Lines
        Array of @{ Text = '...'; Color = 'Cyan' } hashtables.
    .PARAMETER PreviousLineCount
        The line count from the previous frame render. Used to calculate how
        many lines to cursor-up. Pass 0 on first render.
    .OUTPUTS
        [int] The number of lines rendered (pass back as PreviousLineCount next call).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Collections.ArrayList]$Lines,

        [int]$PreviousLineCount = 0
    )

    $term = Get-TerminalSize
    $width = $term.Width
    $height = $term.Height

    # Clamp to terminal height (leave 2 lines safety margin)
    $maxLines = $height - 2
    $renderCount = [Math]::Min($Lines.Count, $maxLines)

    # Build frame in StringBuilder for single-write output
    $sb = [System.Text.StringBuilder]::new(4096)

    # Hide cursor
    [void]$sb.Append("`e[?25l")

    # Move cursor up to overwrite previous frame (relative, scroll-safe)
    if ($PreviousLineCount -gt 0) {
        [void]$sb.Append("`e[${PreviousLineCount}A`e[0G")
    }

    $prevAnsi = ''
    $linesWritten = 0

    for ($i = 0; $i -lt $renderCount; $i++) {
        $line = $Lines[$i]
        $text = if ($line.Text) { $line.Text } else { '' }
        $colorName = if ($line.Color) { $line.Color } else { 'White' }

        # Truncate to terminal width
        if ($text.Length -gt $width) { $text = $text.Substring(0, $width) }

        # Apply ANSI color (only emit if changed from previous line)
        $ansi = if ($script:AnsiColorMap.ContainsKey($colorName)) { $script:AnsiColorMap[$colorName] } else { "`e[37m" }
        if ($ansi -ne $prevAnsi) {
            [void]$sb.Append($ansi)
            $prevAnsi = $ansi
        }

        [void]$sb.Append($text)
        [void]$sb.Append("`e[K")  # Erase to end of line (clears stale chars)
        [void]$sb.Append("`n")
        $linesWritten++
    }

    # Clear leftover lines from previous taller frame
    if ($PreviousLineCount -gt $linesWritten) {
        $extra = $PreviousLineCount - $linesWritten
        for ($j = 0; $j -lt $extra; $j++) {
            [void]$sb.Append("`e[K`n")
            $linesWritten++
        }
    }

    # Reset colors and show cursor
    [void]$sb.Append("`e[0m`e[?25h")

    [Console]::Write($sb.ToString())

    return $linesWritten
}

#endregion

#region Dashboard Functions

function Reset-OrchestratorDashboard {
    <#
    .SYNOPSIS
        Resets the Content Explorer orchestrator dashboard line counter.
        Call before entering a new dashboard loop to prevent cursor misalignment.
    #>
    $script:DashboardLineCount = 0
}

function Reset-AEDashboard {
    <#
    .SYNOPSIS
        Resets the Activity Explorer dashboard line counter.
        Call before entering a new dashboard loop to prevent cursor misalignment.
    #>
    $script:AEDashboardLineCount = 0
}

function Show-OrchestratorDashboard {
    <#
    .SYNOPSIS
        Displays a compact progress dashboard for the orchestrator.
        Redraws in place using cursor positioning for a static display.
    #>
    param(
        [string]$Phase,
        [int]$Completed,
        [int]$Total,
        [array]$Workers,
        [array]$RecentErrors,
        [array]$RecentActivity,
        [array]$DispatchLog,
        [datetime]$ExportStartTime,
        [datetime]$PhaseStartTime,
        [long]$CompletedItems = 0,
        [long]$TotalItems = 0,
        [int]$RemainingAggregates = 0,
        [array]$DetailTasks = @(),
        [hashtable]$ClassifierGroups = @{},
        [int]$TotalLocations = 0,
        [int]$TotalCompleted = 0,
        [int]$TotalErrors = 0,
        [int]$TotalActive = 0,
        [int]$CompletedBaseline = 0,
        [long]$CompletedItemsBaseline = 0
    )

    $pct = if ($Total -gt 0) { [Math]::Round(($Completed / $Total) * 100, 1) } else { 0 }
    $bar = Format-ProgressBar -Percent $pct
    $width = (Get-TerminalSize).Width

    # Build all output lines with their colors
    $lines = [System.Collections.ArrayList]::new()

    [void]$lines.Add(@{ Text = ""; Color = "White" })
    [void]$lines.Add(@{ Text = ("  Content Explorer - Phase: {0} [{1}/{2}] {3} {4}%" -f $Phase.ToUpper(), $Completed, $Total, $bar, $pct); Color = "Cyan" })

    if ($RemainingAggregates -gt 0) {
        [void]$lines.Add(@{ Text = ("  ** {0} aggregate task(s) completing in background **" -f $RemainingAggregates); Color = "Yellow" })
    }

    # Timing line: export start, phase elapsed, ETA
    $now = Get-Date
    $exportElapsed = if ($ExportStartTime -ne [datetime]::MinValue) { $now - $ExportStartTime } else { $null }
    $phaseElapsed = if ($PhaseStartTime -ne [datetime]::MinValue) { $now - $PhaseStartTime } else { $null }

    $timingParts = @()
    if ($ExportStartTime -ne [datetime]::MinValue) {
        $timingParts += "Started: {0}" -f $ExportStartTime.ToString("HH:mm:ss")
    }
    if ($phaseElapsed) {
        $timingParts += "Phase: {0}" -f (Format-TimeSpan -Seconds $phaseElapsed.TotalSeconds)
    }
    if ($exportElapsed) {
        $timingParts += "Total: {0}" -f (Format-TimeSpan -Seconds $exportElapsed.TotalSeconds)
    }

    if ($timingParts.Count -gt 0) {
        [void]$lines.Add(@{ Text = ("  {0}" -f ($timingParts -join "  |  ")); Color = "DarkGray" })
    }

    # ETA calculation — use session-only completions (subtract baseline from prior run) for rate math
    $sessionCompleted = $Completed - $CompletedBaseline
    $sessionCompletedItems = $CompletedItems - $CompletedItemsBaseline
    $etaText = $null
    if ($phaseElapsed -and $phaseElapsed.TotalSeconds -gt 10) {
        if ($Phase -eq "Aggregate") {
            # Aggregate ETA: based on tasks completed vs remaining
            if ($sessionCompleted -gt 0 -and $Completed -lt $Total) {
                $avgSecondsPerTask = $phaseElapsed.TotalSeconds / $sessionCompleted
                $remainingTasks = $Total - $Completed
                $etaSeconds = $avgSecondsPerTask * $remainingTasks
                $etaText = "ETA: {0}  ({1:N1}s/task avg)" -f (Format-TimeSpan -Seconds $etaSeconds), $avgSecondsPerTask
            }
        }
        elseif ($Phase -eq "Detail") {
            # Detail ETA: blended seed rate + actual throughput
            # Seed rate: 40s per 1000 files (0.04s/file). As actual data comes in,
            # linearly blend from seed toward measured rate up to a 5000-file threshold.
            $seedSecondsPerItem = 0.04  # 40s per 1000 files
            $blendThreshold = 5000

            if ($TotalItems -gt 0 -and $CompletedItems -lt $TotalItems) {
                $remainingItems = $TotalItems - $CompletedItems

                if ($sessionCompletedItems -le 0) {
                    # No items completed this session: use seed rate (or task-based fallback)
                    if ($sessionCompleted -gt 0) {
                        $avgSecondsPerTask = $phaseElapsed.TotalSeconds / $sessionCompleted
                        $remainingTasks = $Total - $Completed
                        $etaSeconds = $avgSecondsPerTask * $remainingTasks
                        $etaText = "ETA: ~{0}  (task-based, no items yet)" -f (Format-TimeSpan -Seconds $etaSeconds)
                    }
                    else {
                        $etaSeconds = $seedSecondsPerItem * $remainingItems
                        $etaText = "ETA: ~{0}  (initial estimate, {1:N0} items)" -f (Format-TimeSpan -Seconds $etaSeconds), $TotalItems
                    }
                }
                else {
                    # Blend seed rate with actual measured rate (session-only items)
                    $actualSecondsPerItem = $phaseElapsed.TotalSeconds / $sessionCompletedItems
                    if ($sessionCompletedItems -ge $blendThreshold) {
                        # Past threshold: use pure actual rate
                        $blendedRate = $actualSecondsPerItem
                    }
                    else {
                        # Sliding blend: weight shifts linearly from seed toward actual
                        $actualWeight = $sessionCompletedItems / $blendThreshold
                        $blendedRate = ($seedSecondsPerItem * (1 - $actualWeight)) + ($actualSecondsPerItem * $actualWeight)
                    }
                    $etaSeconds = $blendedRate * $remainingItems
                    $itemPct = [Math]::Round(($CompletedItems / $TotalItems) * 100, 1)
                    $etaText = "ETA: {0}{1}  ({2:N0}/{3:N0} items, {4}%)" -f $(if ($sessionCompletedItems -lt $blendThreshold) { "~" } else { "" }), (Format-TimeSpan -Seconds $etaSeconds), $CompletedItems, $TotalItems, $itemPct
                }
            }
        }
    }

    if ($etaText) {
        [void]$lines.Add(@{ Text = ("  {0}" -f $etaText); Color = "Yellow" })
    }

    [void]$lines.Add(@{ Text = ("  Updated: {0}  [W] Add worker" -f (Get-Date -Format "HH:mm:ss")); Color = "DarkGray" })
    [void]$lines.Add(@{ Text = ""; Color = "White" })

    # --- Aggregated Detail Progress by Classifier ---
    if ($Phase -eq "Detail" -and (($ClassifierGroups.Count -gt 0) -or ($DetailTasks -and $DetailTasks.Count -gt 0))) {
        # Use pre-computed classifier groups if provided, otherwise build from DetailTasks
        if ($ClassifierGroups.Count -gt 0) {
            $classifierGroups = $ClassifierGroups
        } else {
            # Group tasks by TagName + Workload
            $classifierGroups = @{}
            foreach ($dt in $DetailTasks) {
                $groupKey = "{0} / {1}" -f $dt.TagName, $dt.Workload
                if (-not $classifierGroups.ContainsKey($groupKey)) {
                    $classifierGroups[$groupKey] = @{
                        TagName        = $dt.TagName
                        Workload       = $dt.Workload
                        Completed      = 0
                        InProgress     = 0
                        Pending        = 0
                        Error          = 0
                        Total          = 0
                        TotalFiles     = [long]0
                        CompletedFiles = [long]0
                        IsFallback     = $false
                    }
                }
                $classifierGroups[$groupKey].Total++
                $taskFiles = ($dt.ExpectedCount -as [long])
                if ($taskFiles -gt 0) { $classifierGroups[$groupKey].TotalFiles += $taskFiles }
                switch ($dt.Status) {
                    "Completed"  { $classifierGroups[$groupKey].Completed++; if ($taskFiles -gt 0) { $classifierGroups[$groupKey].CompletedFiles += $taskFiles } }
                    "InProgress" { $classifierGroups[$groupKey].InProgress++ }
                    "Pending"    { $classifierGroups[$groupKey].Pending++ }
                    "Error"      { $classifierGroups[$groupKey].Error++ }
                }
                if ($dt.LocationType -eq "WorkloadFallback") {
                    $classifierGroups[$groupKey].IsFallback = $true
                }
            }
        }

        # Sort by total file count descending (largest classifiers first), top 10
        $sortedGroups = @($classifierGroups.GetEnumerator() | Sort-Object { $_.Value.TotalFiles } -Descending)
        $maxClassifierRows = 10

        # Compute column width for alignment
        $displayGroups = if ($sortedGroups.Count -le $maxClassifierRows) { $sortedGroups } else { @($sortedGroups | Select-Object -First ($maxClassifierRows - 1)) }
        $maxKeyLen = 0
        foreach ($g in $displayGroups) {
            $prefix = if ($g.Value.IsFallback) { "[Fallback] " } else { "" }
            $keyLen = $prefix.Length + $g.Key.Length
            if ($keyLen -gt $maxKeyLen) { $maxKeyLen = $keyLen }
        }
        $maxKeyLen = [Math]::Max($maxKeyLen, 20)
        $maxKeyLen = [Math]::Min($maxKeyLen, 50)

        $hBar = [string][char]0x2500
        [void]$lines.Add(@{ Text = "  $($hBar * 2) Detail Progress by Classifier (top 10 by file count) $($hBar * 10)"; Color = "DarkCyan" })

        # Use pre-computed totals if provided, otherwise count from DetailTasks
        if ($TotalLocations -gt 0) {
            $totalLocations = $TotalLocations
            $totalCompleted = $TotalCompleted
            $totalErrors = $TotalErrors
            $totalActive = $TotalActive
        } else {
            $totalLocations = $DetailTasks.Count
            $totalCompleted = @($DetailTasks | Where-Object { $_.Status -eq "Completed" }).Count
            $totalErrors = @($DetailTasks | Where-Object { $_.Status -eq "Error" }).Count
            $totalActive = @($DetailTasks | Where-Object { $_.Status -eq "InProgress" }).Count
        }

        $displayCount = 0
        foreach ($g in $displayGroups) {
            $grp = $g.Value
            $prefix = if ($grp.IsFallback) { "[Fallback] " } else { "" }
            $label = "{0}{1}" -f $prefix, $g.Key
            $padLabel = $label.PadRight($maxKeyLen)

            # File progress
            $filePct = if ($grp.TotalFiles -gt 0) { [Math]::Round(($grp.CompletedFiles / $grp.TotalFiles) * 100, 1) } else { 0 }
            # Location progress
            $locPct = if ($grp.Total -gt 0) { [Math]::Round(($grp.Completed / $grp.Total) * 100, 1) } else { 0 }

            if ($grp.Completed -eq $grp.Total -and $grp.Error -eq 0) {
                $detail = "done  {0:N0} files, {1:N0} locs" -f $grp.TotalFiles, $grp.Total
                $lineColor = "DarkGreen"
            }
            elseif ($grp.Total -gt 0) {
                $detail = "{0:N0}/{1:N0} files ({2}%)  {3:N0}/{4:N0} locs ({5}%)" -f $grp.CompletedFiles, $grp.TotalFiles, $filePct, $grp.Completed, $grp.Total, $locPct
                if ($grp.InProgress -gt 0) {
                    $detail += "  [{0} active]" -f $grp.InProgress
                }
                if ($grp.Error -gt 0) {
                    $detail += "  [{0} err]" -f $grp.Error
                }
                $lineColor = if ($grp.InProgress -gt 0) { "Green" } else { "Gray" }
            }
            else {
                $detail = "0 files, 0 locs"
                $lineColor = "Gray"
            }

            [void]$lines.Add(@{ Text = ("    {0}  {1}" -f $padLabel, $detail); Color = $lineColor })
            $displayCount++
        }

        # Summary line for remaining groups not shown
        if ($sortedGroups.Count -gt $maxClassifierRows) {
            $remainingCount = $sortedGroups.Count - ($maxClassifierRows - 1)
            $remainingFiles = [long]0
            $remainingLocs = 0
            foreach ($g in @($sortedGroups | Select-Object -Skip ($maxClassifierRows - 1))) {
                $remainingFiles += $g.Value.TotalFiles
                $remainingLocs += $g.Value.Total
            }
            [void]$lines.Add(@{ Text = ("    ... and {0} more classifiers ({1:N0} files, {2:N0} locations)" -f $remainingCount, $remainingFiles, $remainingLocs); Color = "DarkGray" })
        }

        # Total summary line
        $activeWorkerCount = if ($Workers) { @($Workers | Where-Object { $_.State -eq "Busy" }).Count } else { 0 }
        [void]$lines.Add(@{ Text = ("  Total: {0:N0}/{1:N0} locations completed | {2} errors | {3} active workers" -f $totalCompleted, $totalLocations, $totalErrors, $activeWorkerCount); Color = "Cyan" })
        [void]$lines.Add(@{ Text = ""; Color = "White" })
    }

    if ($Workers -and $Workers.Count -gt 0) {
        [void]$lines.Add(@{ Text = "  Workers:"; Color = "Gray" })
        # Determine if extended columns are available (detail phase provides them)
        $hasExtended = $Workers[0].ContainsKey('Expected')
        if ($hasExtended) {
            [void]$lines.Add(@{ Text = ("    {0,-8} {1,-10} {2,-10} {3,-10} {4,-16} {5,-8} {6,-8} {7}" -f "PID", "Status", "Time", "Expected", "Progress", "PgSize", "PgTime", "Current Task"); Color = "DarkGray" })
        }
        else {
            [void]$lines.Add(@{ Text = ("    {0,-8} {1,-15} {2}" -f "PID", "Status", "Current Task"); Color = "DarkGray" })
        }
        foreach ($w in $Workers) {
            $color = switch ($w.State) {
                "Busy"    { "Green" }
                "Idle"    { "Yellow" }
                "Dead"    { "Red" }
                default   { "Gray" }
            }
            if ($hasExtended) {
                $timeCol = if ($w.TaskTime) { $w.TaskTime } else { "-" }
                $expCol = if ($w.Expected) { $w.Expected } else { "-" }
                $progCol = if ($w.Progress) { $w.Progress } else { "-" }
                $pgCol = if ($w.PageSize) { $w.PageSize } else { "-" }
                $ptCol = if ($w.LastPage) { $w.LastPage } else { "-" }
                [void]$lines.Add(@{ Text = ("    {0,-8} {1,-10} {2,-10} {3,-10} {4,-16} {5,-8} {6,-8} {7}" -f $w.PID, $w.State, $timeCol, $expCol, $progCol, $pgCol, $ptCol, ($w.CurrentTask ?? "-")); Color = $color })
            }
            else {
                [void]$lines.Add(@{ Text = ("    {0,-8} {1,-15} {2}" -f $w.PID, $w.State, ($w.CurrentTask ?? "-")); Color = $color })
            }
        }
    }

    if ($DispatchLog -and $DispatchLog.Count -gt 0) {
        [void]$lines.Add(@{ Text = ""; Color = "White" })
        $dispatchSlice = @($DispatchLog | Select-Object -Last 4)
        [void]$lines.Add(@{ Text = ("  Dispatch Log ({0} total):" -f $DispatchLog.Count); Color = "DarkMagenta" })
        foreach ($d in $dispatchSlice) {
            [void]$lines.Add(@{ Text = ("    {0}  PID {1,-6} -> {2}" -f $d.Time, $d.PID, $d.Task); Color = "DarkGray" })
        }
    }

    if ($RecentActivity -and $RecentActivity.Count -gt 0) {
        [void]$lines.Add(@{ Text = ""; Color = "White" })
        $activitySlice = @($RecentActivity | Select-Object -Last 4)
        [void]$lines.Add(@{ Text = ("  Recent Activity ({0} total):" -f $RecentActivity.Count); Color = "DarkCyan" })
        foreach ($act in $activitySlice) {
            [void]$lines.Add(@{ Text = ("    {0}  {1}" -f $act.Time, $act.Message); Color = "DarkGray" })
        }
    }

    if ($RecentErrors -and $RecentErrors.Count -gt 0) {
        [void]$lines.Add(@{ Text = ""; Color = "White" })
        $errorSlice = @($RecentErrors | Select-Object -Last 4)
        [void]$lines.Add(@{ Text = ("  Recent Errors ({0} total):" -f $RecentErrors.Count); Color = "Red" })
        foreach ($err in $errorSlice) {
            [void]$lines.Add(@{ Text = ("    {0}  {1}" -f $err.Time, $err.Message); Color = "DarkRed" })
        }
    }
    [void]$lines.Add(@{ Text = ""; Color = "White" })

    $script:DashboardLineCount = Write-DashboardFrame -Lines $lines -PreviousLineCount $(if ($script:DashboardLineCount) { $script:DashboardLineCount } else { 0 })
}

function Show-AEDashboard {
    <#
    .SYNOPSIS
        Displays a compact progress dashboard for Activity Explorer multi-terminal export.
        Redraws in place using cursor positioning for a static display.
    #>
    param(
        [string]$Phase,
        [int]$Completed,
        [int]$Total,
        [array]$Workers,
        [array]$DayTasks,
        [array]$RecentActivity,
        [array]$RecentErrors,
        [datetime]$ExportStartTime,
        [long]$TotalRecords = 0,
        [double]$WeightedPct = -1
    )

    $width = (Get-TerminalSize).Width

    $lines = [System.Collections.ArrayList]::new()

    # Progress bar: use weighted percentage (per-day record progress) when available
    [void]$lines.Add(@{ Text = ""; Color = "White" })
    if ($WeightedPct -ge 0) {
        $pct = [Math]::Round($WeightedPct, 1)
        $bar = Format-ProgressBar -Percent $pct
        $recText = if ($TotalRecords -gt 0) { "  ({0:N0} records)" -f $TotalRecords } else { "" }
        [void]$lines.Add(@{ Text = ("  Activity Explorer - [{0}/{1} days] {2} {3}%{4}" -f $Completed, $Total, $bar, $pct, $recText); Color = "Cyan" })
    }
    else {
        $pct = if ($Total -gt 0) { [Math]::Round(($Completed / $Total) * 100, 1) } else { 0 }
        $bar = Format-ProgressBar -Percent $pct
        [void]$lines.Add(@{ Text = ("  Activity Explorer - [{0}/{1} days] {2} {3}%" -f $Completed, $Total, $bar, $pct); Color = "Cyan" })
    }

    # Timing + ETA
    $now = Get-Date
    $elapsed = if ($ExportStartTime -ne [datetime]::MinValue) { $now - $ExportStartTime } else { $null }
    $timingParts = @()
    if ($ExportStartTime -ne [datetime]::MinValue) {
        $timingParts += "Started: {0}" -f $ExportStartTime.ToString("HH:mm:ss")
    }
    if ($elapsed) {
        $timingParts += "Elapsed: {0}" -f (Format-TimeSpan -Seconds $elapsed.TotalSeconds)
    }

    # ETA based on weighted percentage progress rate
    $etaText = $null
    if ($elapsed -and $elapsed.TotalSeconds -gt 10 -and $pct -gt 0 -and $pct -lt 100) {
        $pctPerSecond = $pct / $elapsed.TotalSeconds
        $remainingPct = 100.0 - $pct
        if ($pctPerSecond -gt 0) {
            $etaSeconds = $remainingPct / $pctPerSecond
            $etaText = "ETA: {0}" -f (Format-TimeSpan -Seconds $etaSeconds)
        }
    }

    if ($timingParts.Count -gt 0) {
        [void]$lines.Add(@{ Text = ("  {0}" -f ($timingParts -join "  |  ")); Color = "DarkGray" })
    }
    if ($etaText) {
        [void]$lines.Add(@{ Text = ("  {0}" -f $etaText); Color = "Yellow" })
    }

    [void]$lines.Add(@{ Text = ("  Updated: {0}" -f (Get-Date -Format "HH:mm:ss")); Color = "DarkGray" })
    [void]$lines.Add(@{ Text = ""; Color = "White" })

    # Task status summary (compact: one line instead of full day table)
    if ($DayTasks -and $DayTasks.Count -gt 0) {
        $completedDays = @($DayTasks | Where-Object { $_.Status -eq "Completed" }).Count
        $activeDays = @($DayTasks | Where-Object { $_.Status -eq "InProgress" }).Count
        $pendingDays = @($DayTasks | Where-Object { $_.Status -eq "Pending" }).Count
        $errorDays = @($DayTasks | Where-Object { $_.Status -eq "Error" }).Count
        $summaryParts = @()
        if ($completedDays -gt 0) { $summaryParts += "Completed: $completedDays" }
        if ($activeDays -gt 0) { $summaryParts += "Active: $activeDays" }
        if ($pendingDays -gt 0) { $summaryParts += "Pending: $pendingDays" }
        if ($errorDays -gt 0) { $summaryParts += "Errors: $errorDays" }
        [void]$lines.Add(@{ Text = ("  {0}" -f ($summaryParts -join "  |  ")); Color = "White" })
        [void]$lines.Add(@{ Text = ""; Color = "White" })
    }

    # Worker status
    if ($Workers -and $Workers.Count -gt 0) {
        $hBar = [string][char]0x2500
        [void]$lines.Add(@{ Text = "  $($hBar * 2) Workers $($hBar * 10)"; Color = "DarkCyan" })
        [void]$lines.Add(@{ Text = ("    {0,-8} {1,-12} {2,-14} {3,8} {4,12} {5,8}" -f "PID", "Status", "Current Day", "Pages", "Records", "%"); Color = "DarkGray" })
        foreach ($w in $Workers) {
            $color = switch ($w.State) {
                "Busy"    { "Green" }
                "Idle"    { "Yellow" }
                "Dead"    { "Red" }
                default   { "Gray" }
            }
            $recText = if ($w.Records) { "{0:N0}" -f [long]$w.Records } else { "-" }
            $pctText = if ($w.RecordPct) { "{0}%" -f $w.RecordPct } else { "-" }
            [void]$lines.Add(@{ Text = ("    {0,-8} {1,-12} {2,-14} {3,8} {4,12} {5,8}" -f $w.PID, $w.State, ($w.CurrentDay ?? "-"), ($w.Pages ?? "-"), $recText, $pctText); Color = $color })
        }
        [void]$lines.Add(@{ Text = ""; Color = "White" })
    }

    # Recent activity
    if ($RecentActivity -and $RecentActivity.Count -gt 0) {
        $activitySlice = @($RecentActivity | Select-Object -Last 4)
        [void]$lines.Add(@{ Text = ("  Recent Activity ({0} total):" -f $RecentActivity.Count); Color = "DarkCyan" })
        foreach ($act in $activitySlice) {
            [void]$lines.Add(@{ Text = ("    {0}  {1}" -f $act.Time, $act.Message); Color = "DarkGray" })
        }
    }

    # Recent errors
    if ($RecentErrors -and $RecentErrors.Count -gt 0) {
        [void]$lines.Add(@{ Text = ""; Color = "White" })
        $errorSlice = @($RecentErrors | Select-Object -Last 3)
        [void]$lines.Add(@{ Text = ("  Recent Errors ({0} total):" -f $RecentErrors.Count); Color = "Red" })
        foreach ($err in $errorSlice) {
            [void]$lines.Add(@{ Text = ("    {0}  {1}" -f $err.Time, $err.Message); Color = "DarkRed" })
        }
    }
    [void]$lines.Add(@{ Text = ""; Color = "White" })

    $script:AEDashboardLineCount = Write-DashboardFrame -Lines $lines -PreviousLineCount $(if ($script:AEDashboardLineCount) { $script:AEDashboardLineCount } else { 0 })
}

#endregion

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

            # Step 5: Dispatch pending tasks to idle workers
            foreach ($w in $WorkerProcesses) {
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

#region Module Exports

Export-ModuleMember -Function @(
    # Export Directory Path Helpers
    'ConvertTo-SafeDirectoryName',
    'Get-CoordinationDir',
    'Get-CompletionsDir',
    'Get-WorkerCoordDir',
    'Get-LogsDir',
    'Get-CEDataDir',
    'Get-AEDataDir',
    'Get-CEClassifierDir',
    'Get-AEDayDir',
    'Initialize-ExportDirectories',
    'Write-CEManifest',
    'Write-AEManifest',

    # Logging
    'Initialize-ExportLog',
    'Write-ExportLog',
    'Get-ExportStatistics',
    'Add-ExportCount',

    # Connection
    'Test-ExportPrerequisites',
    'Connect-Compl8Compliance',
    'Disconnect-Compl8Compliance',

    # Configuration
    'Read-JsonConfig',
    'Get-EnabledItems',
    'Test-ExportConfiguration',

    # File Export
    'ConvertTo-SerializableObject',
    'Export-ToJsonFile',
    'Export-ToCsvFile',

    # Error Handling
    'Write-ExportErrorLog',
    'Format-ErrorDetail',
    'Get-HttpErrorExplanation',

    # Auth Recovery
    'Invoke-WithAuthRecovery',

    # Utility
    'Format-TimeSpan',

    # SIT and Tenant
    'Get-SITsToSkip',
    'Get-Compl8TenantInfo',
    'Get-SitGuidMapping',

    # Content Explorer - Basic
    'Export-ContentExplorer',

    # Content Explorer - Aggregate Discovery
    'Find-RecentAggregateCsv',
    'Save-AggregateMetadata',
    'Save-ExportSettings',
    'Get-ExportSettings',
    'Resolve-CEPageSize',
    'Resolve-AEFilters',
    'Get-TagNamesFromAggregateCsv',
    'Import-AggregateDataFromCsv',

    # Content Explorer - Work Plan & Export
    'New-ContentExplorerWorkPlan',
    'Export-ContentExplorerWithProgress',

    # Content Explorer - Run Tracker
    'Get-ContentExplorerRunTracker',
    'Save-ContentExplorerRunTracker',

    # Content Explorer - Deduplication
    'Remove-DuplicateContentRecordsV2',

    # Content Explorer - Telemetry & Adaptive Paging
    'New-ContentExplorerTelemetry',
    'Get-AdaptivePageSize',
    'Save-ContentExplorerTelemetry',
    'Get-ContentExplorerTelemetryStats',

    # Content Explorer - Progress
    'Get-ContentExplorerAggregateProgress',
    'Write-ContentExplorerProgress',

    # Content Explorer - Retry Bucket
    'Get-RetryBucketTasks',

    # Activity Explorer - Basic
    'Export-ActivityExplorer',

    # Activity Explorer - Resilient Export
    'Export-ActivityExplorerWithProgress',

    # Activity Explorer - Run Tracker
    'Get-ActivityExplorerRunTracker',
    'Save-ActivityExplorerRunTracker',

    # Activity Explorer - Merge
    'Merge-ActivityExplorerPages',

    # Activity Explorer - Helpers
    'Get-PageErrorMessage',
    'Test-PageHasContent',
    'Add-PartialError',
    'Get-RetryDelay',
    'Invoke-RetryWithBackoff',
    'Find-UnknownActivityTypes',
    'Get-ActivityExplorerFilters',

    # Shared Utility
    'Write-ProgressEntry',

    # Worker Health Monitoring
    'Test-WorkerAlive',
    'Get-WorkerState',

    # Phase/Type I/O
    'Write-ExportPhase',
    'Read-ExportPhase',
    'Write-ExportType',
    'Read-ExportType',

    # Task CSV I/O
    'Write-AETaskCsv',
    'Read-AETaskCsv',
    'Write-TaskCsv',
    'Read-TaskCsv',
    'Update-DetailTaskPageSizes',
    'Write-RetryTasksCsv',
    'Show-RetryBucketSummary',
    'Write-RemainingTasksCsv',

    # File-Drop Coordination
    'Get-ExportRunSigningKey',
    'ConvertTo-SignedEnvelopeJson',
    'ConvertFrom-SignedEnvelopeJson',
    'Send-WorkerTask',
    'Receive-WorkerTask',
    'Complete-WorkerTask',

    # Text UI Helpers
    'Get-TerminalSize',
    'Format-ProgressBar',
    'Write-BoxTop',
    'Write-BoxBottom',
    'Write-BoxLine',
    'Write-BoxSeparator',
    'Get-BoxInnerWidth',
    'Write-SectionHeader',
    'Write-Banner',
    'Write-DashboardFrame',

    # Dashboard Functions
    'Reset-OrchestratorDashboard',
    'Reset-AEDashboard',
    'Show-OrchestratorDashboard',
    'Show-AEDashboard',

    # Dispatch Loop Engine
    'Invoke-DispatchLoop',

    # DLP
    'Export-DlpPolicies',
    'Export-SensitiveInfoTypes',

    # Labels
    'Export-SensitivityLabels',
    'Export-RetentionLabels',

    # eDiscovery
    'Export-eDiscoveryCases',

    # RBAC
    'Export-RbacConfiguration'
)

#endregion
