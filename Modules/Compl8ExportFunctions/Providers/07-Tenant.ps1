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
        Reads the externally-provided trainable-classifier cache file.

    .DESCRIPTION
        Microsoft has not shipped a public cmdlet/API for listing trainable
        classifiers, so the names are provided externally: the tool owner's
        separate GetTCs utility (distributed independently of this repo)
        produces a cache file that you drop at
        ConfigFiles/CurrentTenantTCs.local.json (gitignored).

        Input contract — the cache is a JSON object:
          {
            "SchemaVersion": 1,
            "DiscoveredAt": "<ISO-8601 timestamp>",   # drives the staleness warning
            "ClassifierCount": <int>,                  # informational
            "Classifiers": [                           # required
              { "Id": "<guid>", "Name": "...", "DisplayName": "...",
                "Type": "...", "ModelStatus": "...", "IsDeprecated": false },
              ...
            ]
          }
        Only the Classifiers array is required; each entry needs at least a
        Name (the CE discovery loop keys on .Name). Extra properties are
        ignored. tests/test_powershell_smoke.py exercises this shape.

        This function reads the cache and returns objects shaped like the
        other CE discovery cmdlets (each item has a .Name property). If the
        file is missing the CE export proceeds without trainable classifiers.

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
        Write-ExportLog -Message "  The cache is produced by the external GetTCs tool (distributed separately); place its output at ConfigFiles/CurrentTenantTCs.local.json" -Level Info
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
                Write-ExportLog -Message ("  Trainable classifier cache is {0:N0} days old (>{1}). Consider refreshing it with the external GetTCs tool" -f $age.TotalDays, $StaleAfterDays) -Level Warning
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

function Get-SitNamesFromRulePackXml {
    <#
    .SYNOPSIS
        Parses a sensitive information type rule package XML into a GUID-to-name hashtable.

    .DESCRIPTION
        Reads a ClassificationRuleCollection / RulePackage XML document (the format
        Get-DlpSensitiveInformationTypeRulePackage returns in its
        SerializedClassificationRuleCollection property) and extracts the localized
        display name for every rule GUID.

        The rule package schema requires a LocalizedStrings/Resource element (keyed by
        idRef GUID) for every Entity and Affinity rule, so parsing the Resource elements
        yields names for ALL classification rule GUIDs in the pack - including
        Microsoft-internal sub-entity GUIDs that the flat Get-DlpSensitiveInformationType
        list does not surface. For each Resource the default-language Name is preferred
        (Name default="true"), falling back to the first non-empty Name.

        Parsing is namespace-agnostic and resilient to the byte-order-mark / UTF-16
        encoding the cmdlet emits. If the raw bytes are not directly loadable, the text
        is decoded and trimmed to the first '<' before a second parse attempt.

    .PARAMETER XmlBytes
        Raw rule package bytes (e.g. $pack.SerializedClassificationRuleCollection).

    .PARAMETER Path
        Path to a rule package XML file saved on disk.

    .OUTPUTS
        Hashtable of lowercase GUID -> display name. Empty hashtable when nothing parses.

    .EXAMPLE
        $names = Get-SitNamesFromRulePackXml -XmlBytes $pack.SerializedClassificationRuleCollection

    .EXAMPLE
        $names = Get-SitNamesFromRulePackXml -Path "C:\Export\Data\Reference\RulePackages\Microsoft Rule Package.xml"
    #>
    [CmdletBinding(DefaultParameterSetName = 'Bytes')]
    param(
        [Parameter(Mandatory, ParameterSetName = 'Bytes')]
        [byte[]]$XmlBytes,

        [Parameter(Mandatory, ParameterSetName = 'Path')]
        [string]$Path
    )

    if ($PSCmdlet.ParameterSetName -eq 'Path') {
        $XmlBytes = [System.IO.File]::ReadAllBytes($Path)
    }

    $names = @{}
    if (-not $XmlBytes -or $XmlBytes.Count -eq 0) {
        return $names
    }

    # Load the XML. First attempt: stream load (auto-detects BOM/encoding declaration).
    # Fallback: decode as text, trim any leading junk before the first '<', re-parse.
    $doc = [System.Xml.XmlDocument]::new()
    $loaded = $false
    try {
        $stream = [System.IO.MemoryStream]::new($XmlBytes)
        try {
            $doc.Load($stream)
            $loaded = $true
        }
        finally {
            $stream.Dispose()
        }
    }
    catch {
        Write-Verbose "Stream XML load failed, retrying with text decode: $($_.Exception.Message)"
    }

    if (-not $loaded) {
        foreach ($encoding in @([System.Text.Encoding]::Unicode, [System.Text.Encoding]::UTF8)) {
            try {
                $text = $encoding.GetString($XmlBytes)
                $start = $text.IndexOf('<')
                if ($start -lt 0) { continue }
                $doc = [System.Xml.XmlDocument]::new()
                $doc.LoadXml($text.Substring($start))
                $loaded = $true
                break
            }
            catch {
                continue
            }
        }
    }

    if (-not $loaded) {
        Write-ExportLog -Message "  Rule package XML could not be parsed (skipping pack)" -Level Warning
        return $names
    }

    # LocalizedStrings/Resource elements carry the GUID -> localized-name mapping
    # for every Entity and Affinity in the pack (required by the schema).
    foreach ($resource in $doc.GetElementsByTagName('Resource')) {
        $idRef = [string]$resource.GetAttribute('idRef')
        if ($idRef -notmatch '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$') { continue }

        $defaultName = $null
        $firstName = $null
        foreach ($child in $resource.ChildNodes) {
            if ($child.LocalName -ne 'Name') { continue }
            $value = ([string]$child.InnerText).Trim()
            if (-not $value) { continue }
            if (-not $firstName) { $firstName = $value }
            if ([string]$child.GetAttribute('default') -eq 'true') {
                $defaultName = $value
                break
            }
        }

        $name = if ($defaultName) { $defaultName } else { $firstName }
        if ($name) {
            $names[$idRef.ToLowerInvariant()] = $name
        }
    }

    return $names
}

function Export-SitReferenceSnapshot {
    <#
    .SYNOPSIS
        Exports a tenant SIT GUID-to-name reference snapshot into an export directory.

    .DESCRIPTION
        Builds the GUID->name reference data the analytics pipeline uses to resolve
        sensitive information type GUIDs (including Microsoft-internal classification
        sub-entity GUIDs seen in Activity Explorer detections) to display names,
        entirely from supported Security & Compliance PowerShell cmdlets:

        1. Get-DlpSensitiveInformationType - the flat tenant SIT list (Id, Name).
        2. Get-DlpSensitiveInformationTypeRulePackage - every rule package's
           ClassificationRuleCollection XML. The raw XML is saved per pack under
           <ExportDir>\Data\Reference\RulePackages\ and parsed for rule GUID->name
           pairs (Entities AND Affinities, via Get-SitNamesFromRulePackXml).

        The merged map (flat-list names win; rule-pack entries fill the gaps) is
        written to <ExportDir>\CurrentTenantSITs.json - the location the star-schema
        converter (py -m parquet_builder.star.convert) auto-detects, so names ship
        with each tenant export. Properties starting with '_' are metadata.

        Robust by design: if the cmdlets are unavailable (no S&C session) the function
        warns and returns $null; per-pack save/parse failures are warned and skipped;
        whatever succeeded is still written. Never throws.

    .PARAMETER ExportRunDirectory
        Root export directory (the snapshot lands at its top level).

    .PARAMETER Force
        Rebuild the snapshot even if CurrentTenantSITs.json already exists in the
        export directory (default: skip - keeps full exports from running it twice).

    .OUTPUTS
        Hashtable summary: SnapshotPath, SitCount, RulePackCount, RulePackNameCount,
        TotalNames. $null when the cmdlets are unavailable or the snapshot was skipped.

    .EXAMPLE
        Export-SitReferenceSnapshot -ExportRunDirectory $script:ExportRunDirectory

    .EXAMPLE
        # Standalone, against an existing export (connected S&C session required):
        Export-SitReferenceSnapshot -ExportRunDirectory "C:\Exports\Export-20260609-162814" -Force
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ExportRunDirectory,

        [switch]$Force
    )

    try {
        $snapshotPath = Join-Path $ExportRunDirectory "CurrentTenantSITs.json"
        if ((Test-Path $snapshotPath) -and -not $Force) {
            Write-ExportLog -Message "SIT reference snapshot already exists (skipping): $snapshotPath" -Level Info
            return $null
        }

        if (-not (Get-Command Get-DlpSensitiveInformationType -ErrorAction SilentlyContinue)) {
            Write-ExportLog -Message "SIT reference snapshot skipped: Get-DlpSensitiveInformationType not available (no S&C session?)" -Level Warning
            return $null
        }

        Write-ExportLog -Message "Building SIT reference snapshot (flat list + rule packages)..." -Level Info

        # --- 1. Flat tenant SIT list (authoritative names) ---
        $flatMap = @{}
        try {
            foreach ($sit in @(Get-DlpSensitiveInformationType -ErrorAction Stop)) {
                if ($sit.Id -and $sit.Name) {
                    $flatMap[$sit.Id.ToString().ToLowerInvariant()] = [string]$sit.Name
                }
            }
            Write-ExportLog -Message ("  Flat SIT list: {0} entries" -f $flatMap.Count) -Level Info
        }
        catch {
            Write-ExportLog -Message ("  Flat SIT list failed: {0}" -f $_.Exception.Message) -Level Warning
        }

        # --- 2. Rule packages: save raw XML + parse rule GUID->name pairs ---
        $rulePackMap = @{}
        $rulePackCount = 0
        $rulePackDir = Join-Path $ExportRunDirectory "Data" "Reference" "RulePackages"
        if (Get-Command Get-DlpSensitiveInformationTypeRulePackage -ErrorAction SilentlyContinue) {
            try {
                $packs = @(Get-DlpSensitiveInformationTypeRulePackage -ErrorAction Stop)
            }
            catch {
                Write-ExportLog -Message ("  Rule package retrieval failed: {0}" -f $_.Exception.Message) -Level Warning
                $packs = @()
            }

            $packIndex = 0
            foreach ($pack in $packs) {
                $packIndex++
                $packName = $null
                foreach ($candidate in @($pack.LocalizedName, $pack.RuleCollectionName, $pack.Name, $pack.Identity)) {
                    if ($candidate -and -not [string]::IsNullOrWhiteSpace([string]$candidate)) {
                        $packName = [string]$candidate
                        break
                    }
                }
                if (-not $packName) { $packName = "RulePack-$packIndex" }

                try {
                    $packBytes = $pack.SerializedClassificationRuleCollection
                    if ($packBytes -is [string]) {
                        # Defensive: some transports hand the blob back base64-encoded.
                        try { $packBytes = [Convert]::FromBase64String($packBytes) }
                        catch { $packBytes = [System.Text.Encoding]::Unicode.GetBytes($packBytes) }
                    }
                    if (-not $packBytes -or @($packBytes).Count -eq 0) {
                        Write-ExportLog -Message ("  Rule package '{0}' has no serialized rule collection (skipping)" -f $packName) -Level Warning
                        continue
                    }
                    $packBytes = [byte[]]$packBytes

                    # Save the raw XML artifact (faithful copy, as the docs export it)
                    if (-not (Test-Path $rulePackDir)) {
                        New-Item -ItemType Directory -Force -Path $rulePackDir | Out-Null
                    }
                    $packFile = Join-Path $rulePackDir ((ConvertTo-SafeDirectoryName -Name $packName) + ".xml")
                    [System.IO.File]::WriteAllBytes($packFile, $packBytes)
                    $rulePackCount++

                    # Parse rule GUID -> localized name pairs
                    $packNames = Get-SitNamesFromRulePackXml -XmlBytes $packBytes
                    foreach ($key in $packNames.Keys) {
                        if (-not $rulePackMap.ContainsKey($key)) {
                            $rulePackMap[$key] = $packNames[$key]
                        }
                    }
                    Write-ExportLog -Message ("  Rule package '{0}': {1} named rules (saved: {2})" -f $packName, $packNames.Count, $packFile) -Level Info
                }
                catch {
                    Write-ExportLog -Message ("  Rule package '{0}' failed: {1}" -f $packName, $_.Exception.Message) -Level Warning
                }
            }
        }
        else {
            Write-ExportLog -Message "  Get-DlpSensitiveInformationTypeRulePackage not available - snapshot will contain flat-list names only" -Level Warning
        }

        # --- 3. Merge: flat-list names win, rule-pack entries fill gaps ---
        $merged = @{}
        foreach ($key in $rulePackMap.Keys) { $merged[$key] = $rulePackMap[$key] }
        foreach ($key in $flatMap.Keys) { $merged[$key] = $flatMap[$key] }

        if ($merged.Count -eq 0) {
            Write-ExportLog -Message "  No SIT names retrieved - snapshot not written" -Level Warning
            return $null
        }

        # --- 4. Write the snapshot (atomic temp+move; ordered keys) ---
        $payload = [ordered]@{
            "_Description" = "Tenant SIT/classifier GUID->name map. Generated by Export-SitReferenceSnapshot from Get-DlpSensitiveInformationType + rule package XML. Consumed by the star-schema converter (--sit-names auto-detection)."
            "_GeneratedAt" = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
            "_FlatListCount" = $flatMap.Count
            "_RulePackCount" = $rulePackCount
            "_RulePackNameCount" = $rulePackMap.Count
            "_Count" = $merged.Count
        }
        foreach ($key in ($merged.Keys | Sort-Object)) {
            $payload[$key] = $merged[$key]
        }

        if (-not (Test-Path $ExportRunDirectory)) {
            New-Item -ItemType Directory -Force -Path $ExportRunDirectory | Out-Null
        }
        $tmpPath = "$snapshotPath.tmp.$PID"
        $payload | ConvertTo-Json -Depth 5 | Set-Content -Path $tmpPath -Encoding UTF8
        [System.IO.File]::Move($tmpPath, $snapshotPath, $true)

        Write-ExportLog -Message ("SIT reference snapshot written: {0} names ({1} flat-list, {2} from {3} rule packs) -> {4}" -f $merged.Count, $flatMap.Count, $rulePackMap.Count, $rulePackCount, $snapshotPath) -Level Success

        return @{
            SnapshotPath      = $snapshotPath
            SitCount          = $flatMap.Count
            RulePackCount     = $rulePackCount
            RulePackNameCount = $rulePackMap.Count
            TotalNames        = $merged.Count
        }
    }
    catch {
        Write-ExportLog -Message ("SIT reference snapshot failed: {0}" -f $_.Exception.Message) -Level Warning
        return $null
    }
}

#endregion

