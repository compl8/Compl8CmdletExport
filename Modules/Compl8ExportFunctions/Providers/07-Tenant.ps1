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

