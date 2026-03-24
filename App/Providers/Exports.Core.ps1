#region Main Export Functions

function Save-ExportData {
    param(
        [Parameter(Mandatory)]
        $Data,

        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter(Mandatory)]
        [string]$Format,

        [Parameter(Mandatory)]
        [string]$Directory
    )

    # No timestamp in filename since directory already has timestamp
    if ($Format -eq "JSON") {
        $path = Join-Path $Directory "$Name.json"
        Export-ToJsonFile -Data $Data -Path $path
    }
    else {
        $path = Join-Path $Directory "$Name.csv"
        Export-ToCsvFile -Data $Data -Path $path
    }
}

function Invoke-FullExport {
    Write-ExportLog -Message "`n========== Full Configuration Export ==========" -Level Info
    Write-ExportLog -Message "Each module will be saved to a separate file" -Level Info

    $sectionsCompleted = 0
    $sectionsFailed = 0

    # DLP
    Write-ExportLog -Message "`n--- DLP Configuration ---" -Level Info
    try {
        $dlp = Export-DlpPolicies
        $sits = Export-SensitiveInfoTypes
        $dlpExport = @{
            ExportTimestamp = Get-Date -Format "yyyy-MM-ddTHH:mm:ssZ"
            Policies = $dlp.Policies
            Rules = $dlp.Rules
            SensitiveInfoTypes = $sits
        }
        Save-ExportData -Data $dlpExport -Name "DLP-Config" -Format $OutputFormat -Directory $script:ExportRunDirectory
        $sectionsCompleted++
    }
    catch {
        Write-ExportLog -Message "DLP export failed: $($_.Exception.Message)" -Level Error
        $sectionsFailed++
    }

    # Sensitivity Labels
    Write-ExportLog -Message "`n--- Sensitivity Labels ---" -Level Info
    try {
        $sensLabels = Export-SensitivityLabels
        $sensExport = @{
            ExportTimestamp = Get-Date -Format "yyyy-MM-ddTHH:mm:ssZ"
            Labels = $sensLabels.Labels
            Policies = $sensLabels.Policies
        }
        Save-ExportData -Data $sensExport -Name "SensitivityLabels-Config" -Format $OutputFormat -Directory $script:ExportRunDirectory
        $sectionsCompleted++
    }
    catch {
        Write-ExportLog -Message "Sensitivity Labels export failed: $($_.Exception.Message)" -Level Error
        $sectionsFailed++
    }

    # Retention Labels
    Write-ExportLog -Message "`n--- Retention Labels ---" -Level Info
    try {
        $retLabels = Export-RetentionLabels
        $retExport = @{
            ExportTimestamp = Get-Date -Format "yyyy-MM-ddTHH:mm:ssZ"
            Labels = $retLabels.Labels
            Policies = $retLabels.Policies
            Rules = $retLabels.Rules
        }
        Save-ExportData -Data $retExport -Name "RetentionLabels-Config" -Format $OutputFormat -Directory $script:ExportRunDirectory
        $sectionsCompleted++
    }
    catch {
        Write-ExportLog -Message "Retention Labels export failed: $($_.Exception.Message)" -Level Error
        $sectionsFailed++
    }

    # eDiscovery
    Write-ExportLog -Message "`n--- eDiscovery ---" -Level Info
    try {
        $ediscovery = Export-eDiscoveryCases
        $ediscExport = @{
            ExportTimestamp = Get-Date -Format "yyyy-MM-ddTHH:mm:ssZ"
            Cases = $ediscovery.Cases
            Searches = $ediscovery.Searches
        }
        Save-ExportData -Data $ediscExport -Name "eDiscovery-Config" -Format $OutputFormat -Directory $script:ExportRunDirectory
        $sectionsCompleted++
    }
    catch {
        Write-ExportLog -Message "eDiscovery export failed: $($_.Exception.Message)" -Level Error
        $sectionsFailed++
    }

    # RBAC
    Write-ExportLog -Message "`n--- RBAC ---" -Level Info
    try {
        $rbac = Export-RbacConfiguration
        $rbacExport = @{
            ExportTimestamp = Get-Date -Format "yyyy-MM-ddTHH:mm:ssZ"
            RoleGroups = $rbac.RoleGroups
            Members = $rbac.Members
        }
        Save-ExportData -Data $rbacExport -Name "RBAC-Config" -Format $OutputFormat -Directory $script:ExportRunDirectory
        $sectionsCompleted++
    }
    catch {
        Write-ExportLog -Message "RBAC export failed: $($_.Exception.Message)" -Level Error
        $sectionsFailed++
    }

    # Content Explorer
    if ($NoContent) {
        Write-ExportLog -Message "`n--- Content Explorer ---" -Level Info
        Write-ExportLog -Message "  Skipped (-NoContent specified)" -Level Warning
    }
    else {
        Write-ExportLog -Message "`n--- Content Explorer ---" -Level Info
        try {
            Invoke-ContentExplorerExport
            $sectionsCompleted++
        }
        catch {
            Write-ExportLog -Message "Content Explorer export failed: $($_.Exception.Message)" -Level Error
            $sectionsFailed++
        }
    }

    # Activity Explorer
    if ($NoActivity) {
        Write-ExportLog -Message "`n--- Activity Explorer ---" -Level Info
        Write-ExportLog -Message "  Skipped (-NoActivity specified)" -Level Warning
    }
    else {
        Write-ExportLog -Message "`n--- Activity Explorer ---" -Level Info
        try {
            Invoke-ActivityExplorerExport
            $sectionsCompleted++
        }
        catch {
            Write-ExportLog -Message "Activity Explorer export failed: $($_.Exception.Message)" -Level Error
            $sectionsFailed++
        }
    }

    Write-ExportLog -Message "`nSections completed: $sectionsCompleted, Failed: $sectionsFailed" -Level Info
}

function Invoke-DlpExport {
    Write-ExportLog -Message "`n========== DLP Export ==========" -Level Info

    $exportResult = @{
        ExportTimestamp = Get-Date -Format "yyyy-MM-ddTHH:mm:ssZ"
        Policies = @()
        Rules = @()
        SensitiveInfoTypes = @()
    }

    # DLP Policies
    try {
        $dlp = Export-DlpPolicies
        $exportResult.Policies = $dlp.Policies
        $exportResult.Rules = $dlp.Rules
    }
    catch {
        Write-ExportLog -Message "DLP Policies export failed: $($_.Exception.Message)" -Level Error
    }

    # Sensitive Info Types
    try {
        $sits = Export-SensitiveInfoTypes
        $exportResult.SensitiveInfoTypes = $sits
    }
    catch {
        Write-ExportLog -Message "SITs export failed: $($_.Exception.Message)" -Level Error
    }

    if ($OutputFormat -eq "JSON") {
        Save-ExportData -Data $exportResult -Name "DLP-Config" -Format $OutputFormat -Directory $script:ExportRunDirectory
    }
    else {
        # CSV: separate files for each type
        if (@($exportResult.Policies).Count -gt 0) {
            Save-ExportData -Data $exportResult.Policies -Name "DLP-Policies" -Format $OutputFormat -Directory $script:ExportRunDirectory
        }
        if (@($exportResult.Rules).Count -gt 0) {
            Save-ExportData -Data $exportResult.Rules -Name "DLP-Rules" -Format $OutputFormat -Directory $script:ExportRunDirectory
        }
        if (@($exportResult.SensitiveInfoTypes).Count -gt 0) {
            Save-ExportData -Data $exportResult.SensitiveInfoTypes -Name "SensitiveInfoTypes" -Format $OutputFormat -Directory $script:ExportRunDirectory
        }
    }
}

function Invoke-LabelsExport {
    Write-ExportLog -Message "`n========== Labels Export ==========" -Level Info

    $exportResult = @{
        ExportTimestamp = Get-Date -Format "yyyy-MM-ddTHH:mm:ssZ"
        SensitivityLabels = @{ Labels = @(); Policies = @() }
        RetentionLabels = @{ Labels = @(); Policies = @() }
    }

    # Sensitivity Labels
    try {
        $sensLabels = Export-SensitivityLabels
        $exportResult.SensitivityLabels = $sensLabels
    }
    catch {
        Write-ExportLog -Message "Sensitivity Labels export failed: $($_.Exception.Message)" -Level Error
    }

    # Retention Labels
    try {
        $retLabels = Export-RetentionLabels
        $exportResult.RetentionLabels = $retLabels
    }
    catch {
        Write-ExportLog -Message "Retention Labels export failed: $($_.Exception.Message)" -Level Error
    }

    if ($OutputFormat -eq "JSON") {
        Save-ExportData -Data $exportResult -Name "Labels-Config" -Format $OutputFormat -Directory $script:ExportRunDirectory
    }
    else {
        # CSV: separate files
        if (@($exportResult.SensitivityLabels.Labels).Count -gt 0) {
            Save-ExportData -Data $exportResult.SensitivityLabels.Labels -Name "SensitivityLabels" -Format $OutputFormat -Directory $script:ExportRunDirectory
        }
        if (@($exportResult.SensitivityLabels.Policies).Count -gt 0) {
            Save-ExportData -Data $exportResult.SensitivityLabels.Policies -Name "LabelPolicies" -Format $OutputFormat -Directory $script:ExportRunDirectory
        }
        if (@($exportResult.RetentionLabels.Labels).Count -gt 0) {
            Save-ExportData -Data $exportResult.RetentionLabels.Labels -Name "RetentionLabels" -Format $OutputFormat -Directory $script:ExportRunDirectory
        }
        if (@($exportResult.RetentionLabels.Policies).Count -gt 0) {
            Save-ExportData -Data $exportResult.RetentionLabels.Policies -Name "RetentionPolicies" -Format $OutputFormat -Directory $script:ExportRunDirectory
        }
    }
}

