<#
.SYNOPSIS
    Signs all PowerShell scripts in this tool with a code signing certificate.

.DESCRIPTION
    Creates a self-signed code signing certificate (if needed) or uses an existing one,
    then applies Authenticode signatures to all .ps1 and .psm1 files.

    After signing, you can run the scripts on machines with ExecutionPolicy set to
    AllSigned or RemoteSigned (if the cert is trusted on that machine).

.PARAMETER CertificateThumbprint
    Thumbprint of an existing code signing certificate to use. If omitted, the script
    will search for existing self-signed certs or offer to create one.

.PARAMETER CertificateName
    Subject name for a new self-signed certificate. Default: "Compl8CmdletExport Code Signing"

.PARAMETER ExportCert
    Export the certificate to a .cer file for importing on other machines.

.PARAMETER TimestampServer
    URL of a timestamp server. Timestamping ensures signatures remain valid after cert expiry.
    Default: http://timestamp.digicert.com

.EXAMPLE
    # Auto-create or reuse a self-signed cert, sign everything
    .\Sign-Scripts.ps1

.EXAMPLE
    # Use a specific certificate by thumbprint
    .\Sign-Scripts.ps1 -CertificateThumbprint "A1B2C3D4..."

.EXAMPLE
    # Create cert and export .cer for other machines
    .\Sign-Scripts.ps1 -ExportCert
#>
[CmdletBinding()]
param(
    [string]$CertificateThumbprint,
    [string]$CertificateName = "Compl8CmdletExport Code Signing",
    [switch]$ExportCert,
    [string]$TimestampServer = "http://timestamp.digicert.com"
)

$ErrorActionPreference = 'Stop'
$scriptRoot = $PSScriptRoot

# --- Gather files to sign ---
$filesToSign = @(
    Get-ChildItem -Path $scriptRoot -Filter '*.ps1' -File | Where-Object { $_.Name -ne 'Sign-Scripts.ps1' }
    Get-ChildItem -Path $scriptRoot -Recurse -Filter '*.psm1' -File
)

if ($filesToSign.Count -eq 0) {
    Write-Host "No .ps1 or .psm1 files found to sign." -ForegroundColor Yellow
    return
}

Write-Host "Files to sign:" -ForegroundColor Cyan
foreach ($f in $filesToSign) {
    $rel = $f.FullName.Substring($scriptRoot.Length + 1)
    Write-Host "  $rel"
}
Write-Host ""

# --- Locate or create certificate ---
$cert = $null

if ($CertificateThumbprint) {
    # Use the specified thumbprint
    $cert = Get-ChildItem Cert:\CurrentUser\My | Where-Object {
        $_.Thumbprint -eq $CertificateThumbprint -and
        ($_.Extensions | Where-Object { $_.Oid.FriendlyName -eq 'Enhanced Key Usage' }).EnhancedKeyUsages.FriendlyName -contains 'Code Signing'
    }
    if (-not $cert) {
        Write-Error "No code signing certificate found with thumbprint: $CertificateThumbprint"
        return
    }
    Write-Host "Using certificate: $($cert.Subject) [$($cert.Thumbprint)]" -ForegroundColor Green
}
else {
    # Search for existing self-signed code signing certs with our name
    $existing = Get-ChildItem Cert:\CurrentUser\My | Where-Object {
        $_.Subject -like "*$CertificateName*" -and
        $_.NotAfter -gt (Get-Date) -and
        ($_.Extensions | Where-Object { $_.Oid.FriendlyName -eq 'Enhanced Key Usage' }).EnhancedKeyUsages.FriendlyName -contains 'Code Signing'
    } | Sort-Object NotAfter -Descending

    if ($existing) {
        $cert = $existing[0]
        Write-Host "Found existing certificate:" -ForegroundColor Green
        Write-Host "  Subject:    $($cert.Subject)"
        Write-Host "  Thumbprint: $($cert.Thumbprint)"
        Write-Host "  Expires:    $($cert.NotAfter.ToString('yyyy-MM-dd'))"
        Write-Host ""

        $use = Read-Host "Use this certificate? [Y/n]"
        if ($use -eq 'n' -or $use -eq 'N') {
            $cert = $null
        }
    }

    if (-not $cert) {
        Write-Host "Creating new self-signed code signing certificate..." -ForegroundColor Cyan
        Write-Host "  Name:    CN=$CertificateName"
        Write-Host "  Store:   Cert:\CurrentUser\My"
        Write-Host "  Expires: 3 years"
        Write-Host ""

        $cert = New-SelfSignedCertificate `
            -Type CodeSigningCert `
            -Subject "CN=$CertificateName" `
            -CertStoreLocation Cert:\CurrentUser\My `
            -NotAfter (Get-Date).AddYears(3) `
            -FriendlyName $CertificateName

        Write-Host "Certificate created:" -ForegroundColor Green
        Write-Host "  Thumbprint: $($cert.Thumbprint)"
        Write-Host ""

        # Trust the cert on this machine
        Write-Host "Adding certificate to Trusted Root store (may require elevation)..." -ForegroundColor Cyan
        try {
            $rootStore = New-Object System.Security.Cryptography.X509Certificates.X509Store(
                [System.Security.Cryptography.X509Certificates.StoreName]::Root,
                [System.Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser
            )
            $rootStore.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
            $rootStore.Add($cert)
            $rootStore.Close()
            Write-Host "  Added to CurrentUser\Root (trusted on this machine)" -ForegroundColor Green
        }
        catch {
            Write-Warning "Could not add to trusted root: $($_.Exception.Message)"
            Write-Warning "Scripts will show as signed but 'untrusted' until you trust the cert."
        }
        Write-Host ""
    }
}

# --- Export certificate if requested ---
if ($ExportCert) {
    $certFile = Join-Path $scriptRoot "$($CertificateName -replace '[^a-zA-Z0-9]', '-').cer"
    $certBytes = $cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert)
    [System.IO.File]::WriteAllBytes($certFile, $certBytes)
    Write-Host "Certificate exported to: $certFile" -ForegroundColor Green
    Write-Host "  Import this on other machines to trust the signature:" -ForegroundColor Cyan
    Write-Host "    Import-Certificate -FilePath `"$($certFile | Split-Path -Leaf)`" -CertStoreLocation Cert:\CurrentUser\Root" -ForegroundColor DarkGray
    Write-Host ""
}

# --- Sign all files ---
Write-Host "Signing files..." -ForegroundColor Cyan
$succeeded = 0
$failed = 0

foreach ($file in $filesToSign) {
    $rel = $file.FullName.Substring($scriptRoot.Length + 1)
    try {
        $params = @{
            FilePath    = $file.FullName
            Certificate = $cert
            HashAlgorithm = 'SHA256'
        }
        if ($TimestampServer) {
            $params['TimestampServer'] = $TimestampServer
        }
        $sig = Set-AuthenticodeSignature @params

        if ($sig.Status -eq 'Valid') {
            Write-Host "  SIGNED  $rel" -ForegroundColor Green
            $succeeded++
        }
        else {
            Write-Host "  WARN   $rel - $($sig.StatusMessage)" -ForegroundColor Yellow
            $succeeded++
        }
    }
    catch {
        Write-Host "  FAIL   $rel - $($_.Exception.Message)" -ForegroundColor Red
        $failed++
    }
}

# --- Summary ---
Write-Host ""
Write-Host "--- Signing Complete ---" -ForegroundColor Cyan
Write-Host "  Signed: $succeeded"
if ($failed -gt 0) {
    Write-Host "  Failed: $failed" -ForegroundColor Red
}
Write-Host "  Cert:   $($cert.Thumbprint)"
Write-Host ""
Write-Host "To verify a signature:" -ForegroundColor DarkGray
Write-Host "  Get-AuthenticodeSignature .\Export-Compl8Configuration.ps1" -ForegroundColor DarkGray
Write-Host ""
if (-not $ExportCert) {
    Write-Host "To export the cert for other machines, re-run with -ExportCert" -ForegroundColor DarkGray
}
