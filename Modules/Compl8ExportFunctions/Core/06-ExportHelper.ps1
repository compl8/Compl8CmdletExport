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

