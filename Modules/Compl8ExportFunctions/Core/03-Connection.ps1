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

#endregion

