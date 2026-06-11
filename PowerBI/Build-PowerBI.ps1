<#
.SYNOPSIS
    Generate a Power BI project from the star-schema SSOT and compile it to PBIT.

.DESCRIPTION
    Wraps the python builders in PowerBI\builders and pbi-tools.core:
      1. Generates the PbixProj folder (model TMDL from parquet_builder.star.schema).
      2. Compiles it to a .pbit with pbi-tools.core (requires DOTNET_ROLL_FORWARD=Major).
      3. Reports the output paths.

.EXAMPLE
    .\Build-PowerBI.ps1 -Project _smoke -ParquetRoot "D:\Exports\PowerBI-AE-Parquet-v6"
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('ActivityExplorer', 'ContentExplorerSITRisk', '_smoke')]
    [string] $Project,

    # When omitted, each builder's own CHANGEME placeholder default applies
    # (AE: PowerBI-AE-Parquet-v6, CE: PowerBI-CE-Parquet).
    [string] $ParquetRoot = '',

    [string] $PbiToolsCorePath = 'C:\Tools\pbi-tools-net9\pbi-tools.core.exe'
)

$ErrorActionPreference = 'Stop'

$powerBIRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = Split-Path -Parent $powerBIRoot

$builderModules = @{
    '_smoke'                 = 'PowerBI.builders.build_smoke'
    'ActivityExplorer'       = 'PowerBI.builders.build_activity_explorer'
    'ContentExplorerSITRisk' = 'PowerBI.builders.build_content_explorer'
}
$pbitNames = @{
    '_smoke'                 = 'Compl8Smoke.pbit'
    'ActivityExplorer'       = 'ActivityExplorerRisk.pbit'
    'ContentExplorerSITRisk' = 'ContentExplorerSITRisk.pbit'
}

$builderModule = $builderModules[$Project]
$moduleRelPath = Join-Path $repoRoot (($builderModule -replace '\.', '\') + '.py')
if (-not (Test-Path -LiteralPath $moduleRelPath)) {
    throw "Builder for project '$Project' is not implemented yet (missing $moduleRelPath)."
}

if (-not (Test-Path -LiteralPath $PbiToolsCorePath)) {
    $cmd = Get-Command 'pbi-tools.core' -ErrorAction SilentlyContinue
    if ($cmd) {
        $PbiToolsCorePath = $cmd.Source
    }
    else {
        throw "pbi-tools.core not found. Install it or pass -PbiToolsCorePath."
    }
}

$projectDir = Join-Path $powerBIRoot "projects\$Project\pbix"
$outputPbit = Join-Path $powerBIRoot "projects\$Project\$($pbitNames[$Project])"

Push-Location $repoRoot
try {
    Write-Host "Generating project '$Project' -> $projectDir"
    $generateArgs = @('-m', $builderModule, '--output-dir', $projectDir, '--overwrite')
    if ($ParquetRoot) {
        $generateArgs += @('--parquet-root', $ParquetRoot)
    }
    py @generateArgs
    if ($LASTEXITCODE -ne 0) {
        throw "Project generation failed with exit code $LASTEXITCODE"
    }

    Write-Host "Compiling -> $outputPbit"
    $env:DOTNET_ROLL_FORWARD = 'Major'
    & $PbiToolsCorePath compile $projectDir $outputPbit PBIT true
    if ($LASTEXITCODE -ne 0) {
        throw "pbi-tools compile failed with exit code $LASTEXITCODE"
    }

    Write-Host ''
    Write-Host "Project folder : $projectDir"
    Write-Host "Compiled PBIT  : $outputPbit"
    $effectiveRoot = if ($ParquetRoot) { $ParquetRoot } else { '(builder CHANGEME default)' }
    Write-Host "ParquetRoot    : $effectiveRoot (change in Power BI Desktop via Transform data > Edit parameters)"
}
finally {
    Pop-Location
}
