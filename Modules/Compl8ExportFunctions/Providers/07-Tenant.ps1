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
        $ConfigPath = Join-Path (Join-Path $projectRoot "ConfigFiles") "SITstoSkip.json"
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

function Get-TrainableClassifiersFromCache {
    <#
    .SYNOPSIS
        Reads the trainable-classifier cache produced by Helpers/Get-TrainableClassifiers.py.

    .DESCRIPTION
        Microsoft has not shipped a public cmdlet/API for listing trainable
        classifiers, so the Compl8 pipeline gets them via a Playwright-based
        scraper that talks to the Purview portal's internal IPML endpoints.
        That helper writes ConfigFiles/CurrentTenantTCs.local.json. This
        function reads it and returns objects shaped like the other CE
        discovery cmdlets (each item has a .Name property).

    .PARAMETER ConfigPath
        Path to the cache file. Defaults to ConfigFiles/CurrentTenantTCs.local.json
        relative to the project root.

    .PARAMETER StaleAfterDays
        Emit a warning if the cache is older than this many days. Default 30.

    .OUTPUTS
        Array of PSCustomObjects with Id, Name, DisplayName, Type, ModelStatus,
        IsDeprecated. Returns empty array if the cache file is missing or
        unreadable (the caller will surface the warning).

    .EXAMPLE
        $classifiers = Get-TrainableClassifiersFromCache
        $names = @($classifiers.Name)
    #>
    [CmdletBinding()]
    param(
        [string]$ConfigPath,
        [int]$StaleAfterDays = 30
    )

    if (-not $ConfigPath) {
        $projectRoot = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $PSCommandPath))
        $ConfigPath = Join-Path $projectRoot "ConfigFiles" "CurrentTenantTCs.local.json"
    }

    if (-not (Test-Path $ConfigPath)) {
        Write-ExportLog -Message "  Trainable classifier cache not found: $ConfigPath" -Level Warning
        Write-ExportLog -Message "  Run: python Helpers/Get-TrainableClassifiers.py" -Level Info
        return @()
    }

    try {
        $cache = Get-Content -Path $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
    }
    catch {
        Write-ExportLog -Message "  Trainable classifier cache is malformed ($ConfigPath): $($_.Exception.Message)" -Level Error
        return @()
    }

    if (-not $cache.Classifiers) {
        Write-ExportLog -Message "  Trainable classifier cache has no classifiers" -Level Warning
        return @()
    }

    # Staleness check (best-effort; cache may not carry DiscoveredAt on old versions)
    try {
        if ($cache.DiscoveredAt) {
            $discoveredAt = [datetime]::Parse($cache.DiscoveredAt)
            $age = (Get-Date) - $discoveredAt
            if ($age.TotalDays -gt $StaleAfterDays) {
                Write-ExportLog -Message ("  Trainable classifier cache is {0:N0} days old (>{1}). Consider rerunning Helpers/Get-TrainableClassifiers.py" -f $age.TotalDays, $StaleAfterDays) -Level Warning
            }
        }
    }
    catch {
        # Bad timestamp - ignore, just don't warn
    }

    $tcCount = @($cache.Classifiers).Count
    Write-ExportLog -Message ("  Loaded {0} trainable classifiers from cache (discovered {1})" -f $tcCount, $cache.DiscoveredAt) -Level Info

    # Project to the shape the orchestrator's discovery loop expects ($_.Name)
    return @($cache.Classifiers | ForEach-Object {
        [PSCustomObject]@{
            Id           = $_.Id
            Name         = $_.Name
            DisplayName  = $_.DisplayName
            Type         = $_.Type
            ModelStatus  = $_.ModelStatus
            IsDeprecated = [bool]$_.IsDeprecated
        }
    })
}

#endregion

