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

