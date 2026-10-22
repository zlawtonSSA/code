# CSV Diff Script - Compare two CSV files and extract changes
# Usage: .\Compare-CSVFiles.ps1 -File1 "old.csv" -File2 "new.csv" -KeyColumn "ID"

param(
    [Parameter(Mandatory=$true)]
    [string]$File1,
    
    [Parameter(Mandatory=$true)]
    [string]$File2,
    
    [Parameter(Mandatory=$true)]
    [string]$KeyColumn,
    
    [string]$OutputFile = "differences.csv"
)

# Import both CSV files
Write-Host "Loading CSV files..." -ForegroundColor Cyan
$csv1 = Import-Csv -Path $File1
$csv2 = Import-Csv -Path $File2

#sort both CSV files by the key column to ensure they are in the same order for comparison
$csv1 = $csv1 | Sort-Object -Property $KeyColumn
$csv2 = $csv2 | Sort-Object -Property $KeyColumn

# Initialize results array
$differences = @()

# Create hashtables for faster lookups
$hash1 = @{}
$hash2 = @{}

foreach ($row in $csv1) {
    #only include the rows where the Active column is equal to 'Y'
    if ($row.Active -eq 'Y') {
        $hash1[$row.$KeyColumn] = $row
    }
}

foreach ($row in $csv2) {
    $hash2[$row.$KeyColumn] = $row
}

Write-Host "Comparing files..." -ForegroundColor Cyan

# Find added rows
foreach ($key in $hash2.Keys) {
    # If the key from the second file doesn't exist in the first file, it's an added row
    if (-not $hash1.ContainsKey($key)) {
        # Create a new object for the added row, excluding the ChangeType and ChangedFields properties if they exist, 
        # and add a ChangeType property with value 'Added'
        $row = $hash2[$key] | Select-Object * -ExcludeProperty ChangeType, ChangedFields | 
               Select-Object *, @{Name='ChangeType';Expression={'Added'}}
        $differences += $row
    }
}

# Find removed rows
foreach ($key in $hash1.Keys) {
    if (-not $hash2.ContainsKey($key)) {
        $row = $hash1[$key] | Select-Object * -ExcludeProperty ChangeType, ChangedFields | 
               Select-Object *, @{Name='ChangeType';Expression={'Removed'}}
        $differences += $row
    }
}

foreach ($key in $hash1.Keys) {
    if ($hash2.ContainsKey($key)) {
        $row1 = $hash1[$key]
        $row2 = $hash2[$key]
        
        $isModified = $false
        $changedFields = @()

        $excludeFromComparison = @('Remove Skill 1', 'Remove Skill 2', 'Remove Skill 3', 'Remove Skill 4', 'ChangeType', 'ChangedFields')

        foreach ($prop in $row1.PSObject.Properties.Name) {
            if ($prop -in $excludeFromComparison) { continue }
            if ($row1.$prop -ne $row2.$prop) {
                $isModified = $true
                $changedFields += "$prop (Old: '$($row1.$prop)' -> New: '$($row2.$prop)')"

                if ($prop -like "*Skill Code 1*" -and $row1.$prop -and -not $row2.$prop) {
                    $row2 | Add-Member -MemberType NoteProperty -Name "Remove Skill 1" -Value "Y" -Force
                }
                if ($prop -like "*Skill Code 2*" -and $row1.$prop -and -not $row2.$prop) {
                    $row2 | Add-Member -MemberType NoteProperty -Name "Remove Skill 2" -Value "Y" -Force
                }
                if ($prop -like "*Skill Code 3*" -and $row1.$prop -and -not $row2.$prop) {
                    $row2 | Add-Member -MemberType NoteProperty -Name "Remove Skill 3" -Value "Y" -Force
                }
                if ($prop -like "*Skill Code 4*" -and $row1.$prop -and -not $row2.$prop) {
                    $row2 | Add-Member -MemberType NoteProperty -Name "Remove Skill 4" -Value "Y" -Force
                }
            }
        }
        
        if ($isModified) {
            if ($row2.PSObject.Properties.Name -contains 'Remove Skill 4') {
                Write-Host "DEBUG: Remove Skill 4 = Y on ID $($row2.$KeyColumn)" -ForegroundColor Magenta
            }
            else {
                Write-Host "DEBUG: Remove Skill 4 NOT found on ID $($row2.$KeyColumn) after prop loop" -ForegroundColor Red
            }

            $row = $row2 | Select-Object * -ExcludeProperty ChangeType, ChangedFields | 
                   Select-Object *, 
                       @{Name='ChangeType';Expression={'Modified'}},
                       @{Name='ChangedFields';Expression={$changedFields -join '; '}}

            if ($isModified) {
    $row = $row2 | Select-Object * -ExcludeProperty ChangeType, ChangedFields | 
           Select-Object *, 
               @{Name='ChangeType';Expression={'Modified'}},
               @{Name='ChangedFields';Expression={$changedFields -join '; '}}

    # DEBUG
    Write-Host "DEBUG row properties: $($row.PSObject.Properties.Name -join ', ')" -ForegroundColor Yellow

    $differences += $row
}

            $differences += $row
        }
    }
}

# Display summary
Write-Host "`nComparison Summary:" -ForegroundColor Green
Write-Host "==================" -ForegroundColor Green
Write-Host "Total rows in last import: $($csv1.Count)" -ForegroundColor Yellow
Write-Host "Total rows in current import: $($csv2.Count)" -ForegroundColor Yellow
Write-Host "Added: $(($differences | Where-Object {$_.ChangeType -eq 'Added'}).Count)" -ForegroundColor Green
Write-Host "Removed: $(($differences | Where-Object {$_.ChangeType -eq 'Removed'}).Count)" -ForegroundColor Red
Write-Host "Modified: $(($differences | Where-Object {$_.ChangeType -eq 'Modified'}).Count)" -ForegroundColor Cyan

# Return removed items
#$removedItems = $differences | Where-Object {$_.ChangeType -eq 'Removed'} | Select-Object * -ExcludeProperty ChangeType, ChangedFields
#return $removedItems

return $differences