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
                Write-ExportLog -Message "OUTDATED (v$version, need 3.2.0+)" -Level Error
                Write-ExportLog -Message "  Update with: Install-Module ExchangeOnlineManagement -Force" -Level Info
                $passed = $false
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


function Test-AuthConfig {
    <#
    .SYNOPSIS
        Validates AuthConfig.json shape and, when a thumbprint is given, confirms the cert
        exists in the certificate store and has not expired (within a configurable buffer).

    .DESCRIPTION
        Designed to be called before attempting a connection in unattended mode so that a
        misconfigured cert fails fast with a structured result rather than an opaque
        Connect-IPPSSession error.

        Validation order:
          1. Config file exists and parses → else ConfigError
          2. UseCertificateAuth = "True" with AppId, Organization, and CertificateThumbprint
             all non-empty → else ConfigError (UPN / interactive config is not valid for
             unattended use; caller must configure cert auth)
          3. Cert found via -GetCertificate scriptblock → else AuthFailed
          4. Cert NotAfter > (now + Buffer) → else AuthFailed ("expired or expiring")

        Returns a [pscustomobject] with:
          IsValid    [bool]    — $true only when all checks pass
          Status     [string]  — '' | 'ConfigError' | 'AuthFailed'
          AuthType   [string]  — 'Certificate' | 'Interactive' | 'UPN'
          Thumbprint [string]  — thumbprint from config (or '' if absent)
          NotAfter   [datetime] or $null — cert expiry from the store
          Errors     [string[]] — human-readable failure reasons (empty on success)

    .PARAMETER ConfigPath
        Absolute path to AuthConfig.json.

    .PARAMETER Buffer
        How long before actual expiry a cert is considered "expiring" (default 1 day).

    .PARAMETER GetCertificate
        Scriptblock that accepts a thumbprint string and returns an object with a NotAfter
        property, or $null when not found.  Default: searches Cert:\CurrentUser\My then
        Cert:\LocalMachine\My.  Override in tests to avoid needing a real certificate store.

    .OUTPUTS
        [pscustomobject] — see Description for field list.

    .EXAMPLE
        # Real call in production:
        $check = Test-AuthConfig -ConfigPath (Join-Path $scriptRoot 'ConfigFiles\AuthConfig.json')
        if (-not $check.IsValid) { Write-Error "Auth pre-flight failed: $($check.Errors -join '; ')" }

    .EXAMPLE
        # In-process test with an injected fake cert:
        $fakeCert = [pscustomobject]@{ NotAfter = (Get-Date).AddDays(30) }
        $check = Test-AuthConfig -ConfigPath $tmpJson -GetCertificate { param($tp) $fakeCert }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ConfigPath,

        [timespan]$Buffer = [timespan]::FromDays(1),

        [scriptblock]$GetCertificate = {
            param([string]$Thumbprint)
            $cert = Get-ChildItem -Path 'Cert:\CurrentUser\My' -ErrorAction SilentlyContinue |
                    Where-Object { $_.Thumbprint -eq $Thumbprint } |
                    Select-Object -First 1
            if ($cert) { return $cert }
            Get-ChildItem -Path 'Cert:\LocalMachine\My' -ErrorAction SilentlyContinue |
                    Where-Object { $_.Thumbprint -eq $Thumbprint } |
                    Select-Object -First 1
        }
    )

    $errors    = [System.Collections.Generic.List[string]]::new()
    $thumbprint = ''
    $notAfter  = $null
    $authType  = 'Interactive'

    # ── 1. Config file present and parseable ──────────────────────────────────
    if (-not (Test-Path $ConfigPath)) {
        $errors.Add("AuthConfig.json not found at: $ConfigPath")
        return [pscustomobject]@{
            IsValid    = $false
            Status     = 'ConfigError'
            AuthType   = $authType
            Thumbprint = $thumbprint
            NotAfter   = $notAfter
            Errors     = $errors.ToArray()
        }
    }

    $authConfig = $null
    try {
        $authConfig = Get-Content -Raw -Path $ConfigPath -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        $errors.Add("AuthConfig.json could not be parsed: $($_.Exception.Message)")
        return [pscustomobject]@{
            IsValid    = $false
            Status     = 'ConfigError'
            AuthType   = $authType
            Thumbprint = $thumbprint
            NotAfter   = $notAfter
            Errors     = $errors.ToArray()
        }
    }

    # ── 2. Certificate auth fields present ───────────────────────────────────
    $useCert = $authConfig.UseCertificateAuth -eq 'True'
    if ($authConfig.PSObject.Properties['UserPrincipalName'] -and -not [string]::IsNullOrWhiteSpace($authConfig.UserPrincipalName)) {
        $authType = 'UPN'
    }
    if ($useCert) {
        $authType = 'Certificate'
    }

    if (-not $useCert) {
        $errors.Add("AuthConfig.json does not enable certificate auth (UseCertificateAuth != 'True'). Unattended mode requires cert auth.")
        return [pscustomobject]@{
            IsValid    = $false
            Status     = 'ConfigError'
            AuthType   = $authType
            Thumbprint = $thumbprint
            NotAfter   = $notAfter
            Errors     = $errors.ToArray()
        }
    }

    $missingFields = @()
    if ([string]::IsNullOrWhiteSpace($authConfig.AppId))                   { $missingFields += 'AppId' }
    if ([string]::IsNullOrWhiteSpace($authConfig.Organization))            { $missingFields += 'Organization' }
    if ([string]::IsNullOrWhiteSpace($authConfig.CertificateThumbprint))   { $missingFields += 'CertificateThumbprint' }

    if ($missingFields.Count -gt 0) {
        $errors.Add("AuthConfig.json is missing required cert-auth field(s): $($missingFields -join ', ')")
        return [pscustomobject]@{
            IsValid    = $false
            Status     = 'ConfigError'
            AuthType   = $authType
            Thumbprint = $thumbprint
            NotAfter   = $notAfter
            Errors     = $errors.ToArray()
        }
    }

    # Trim so a copy-paste leading/trailing space in AuthConfig.json doesn't turn a
    # present cert into a confusing AuthFailed (store lookup is exact on thumbprint).
    $thumbprint = $authConfig.CertificateThumbprint.Trim()

    # ── 3. Cert exists in store ───────────────────────────────────────────────
    $cert = $null
    $certLookupFailed = $false
    try {
        $cert = & $GetCertificate $thumbprint
    }
    catch {
        $errors.Add("GetCertificate scriptblock threw: $($_.Exception.Message)")
        $certLookupFailed = $true
    }

    if (-not $cert) {
        # Only add the "not found" message when the lookup actually returned null;
        # if it threw, the catch already recorded the real cause (avoid double error).
        if (-not $certLookupFailed) {
            $errors.Add("Certificate with thumbprint '$thumbprint' not found in Cert:\CurrentUser\My or Cert:\LocalMachine\My")
        }
        return [pscustomobject]@{
            IsValid    = $false
            Status     = 'AuthFailed'
            AuthType   = $authType
            Thumbprint = $thumbprint
            NotAfter   = $notAfter
            Errors     = $errors.ToArray()
        }
    }

    $notAfter = $cert.NotAfter

    # ── 4. Cert not expired (with buffer) ────────────────────────────────────
    $deadline = (Get-Date).Add($Buffer)
    if ($notAfter -le $deadline) {
        $errors.Add("Certificate '$thumbprint' has expired or will expire within $($Buffer.TotalHours)h (NotAfter: $($notAfter.ToString('u'))). Renew before running unattended.")
        return [pscustomobject]@{
            IsValid    = $false
            Status     = 'AuthFailed'
            AuthType   = $authType
            Thumbprint = $thumbprint
            NotAfter   = $notAfter
            Errors     = $errors.ToArray()
        }
    }

    return [pscustomobject]@{
        IsValid    = $true
        Status     = ''
        AuthType   = $authType
        Thumbprint = $thumbprint
        NotAfter   = $notAfter
        Errors     = @()
    }
}

#endregion

