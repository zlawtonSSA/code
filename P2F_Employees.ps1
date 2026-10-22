# Paycom To Frontline EMPLOYEES Export Data Transformer
# This script transforms exported data from Paycom into a format that will import to Frontline

# Initialize error tracking
$ErrorActionPreference = 'Stop'
$scriptStartTime = Get-Date
$script:logMessages = [System.Collections.ArrayList]@()


function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Warning', 'Error', 'Success')]
        [string]$Level = 'Info'
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    [void]$script:logMessages.Add($logEntry)
    
    switch ($Level) {
        'Info'    { Write-Host $logEntry -ForegroundColor Cyan }
        'Warning' { Write-Host $logEntry -ForegroundColor Yellow }
        'Error'   { Write-Host $logEntry -ForegroundColor Red }
        'Success' { Write-Host $logEntry -ForegroundColor Green }
    }
}

function Save-LogFile {
    # FIX 1: Renamed parameter from $sourceFolder to $DestinationFolder to match all call sites.
    param([string]$DestinationFolder)
    
    try {
        $logFileName = "P2F_Employees_TransformLog_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
        $logFilePath = Join-Path -Path "C:\_Paycom Frontline Data Transfer\Export_from_Paycom\Employees\_logs" -ChildPath $logFileName
        $script:logMessages | Out-File -FilePath $logFilePath -Encoding UTF8
        Write-Log "Log file saved: $logFilePath" -Level Info
        return $logFilePath 
    }
    catch {
        Write-Log "Failed to save log file: $($_.Exception.Message)" -Level Warning
        return $null        
    }
}

function Get-DepartmentName {
    param (
        [string]$DepartmentCode
    )
    $departmentMapping = @{
        '111' = 'ENGLISH'
        '112' = 'HISTORY / SOCIAL'
        '113' = 'MATHEMATICS'
        '114' = 'SCIENCE'
        '121' = 'ARTS'
        '122' = 'COMPUTER SCIENCE'
        '123' = 'HEALTH PE'
        '124' = 'WORLD LANGUAGE'
        '131' = 'LIBRARY'
        '141' = 'COLLEGE COUNSELING'
        '151' = 'STUDENT SERVICES'
        '191' = 'SUBSTITUTE'
        '211' = 'ENGLISH'
        '212' = 'HISTORY / SOCIAL'
        '213' = 'MATHEMATICS'
        '214' = 'SCIENCE'
        '221' = 'ARTS'
        '222' = 'COMPUTER SCIENCE'
        '223' = 'HEALTH PE'
        '224' = 'WORLD LANGUAGE'
        '231' = 'LIBRARY'
        '251' = 'STUDENT SERVICES'
        '291' = 'SUBSTITUTE'
        '309' = 'PRE KINDERGARTEN'
        '310' = 'KINDERGARTEN'
        '311' = 'FIRST GRADE'
        '312' = 'SECOND GRADE'
        '313' = 'THIRD GRADE'
        '314' = 'FOURTH GRADE'
        '315' = 'FIFTH GRADE'
        '320' = 'SCIENCE'
        '321' = 'ARTS'
        '322' = 'COMPUTER SCIENCE'
        '323' = 'HEALTH PE'
        '324' = 'WORLD LANGUAGE'
        '331' = 'LIBRARY'
        '608' = 'JPK'
        '609' = 'SPK'
        '610' = 'KINDERGARTEN'
        '611' = 'FIRST GRADE'
        '612' = 'SECOND GRADE'
        '613' = 'THIRD GRADE'
        '614' = 'FOURTH GRADE'
        '615' = 'FIFTH GRADE'
        '620' = 'SCIENCE'
        '621' = 'ARTS'
        '623' = 'HEALTH PE'
        '624' = 'WORLD LANGUAGE'
        '631' = 'LIBRARY'
        '651' = 'STUDENT SERVICES'
    }

    return $departmentMapping[$DepartmentCode]
}

try {
    Write-Log "=== Paycom to Frontline EMPLOYEES Data Transformation Started ===" -Level Info
    
    # Get the script's directory
    $scriptPath = $PSScriptRoot
    if ([string]::IsNullOrEmpty($scriptPath)) {
        $scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
    }
    if ([string]::IsNullOrEmpty($scriptPath)) {
        $scriptPath = Get-Location
    }
    
    Write-Log "Script running from: $scriptPath" -Level Info
    
    $userProfile = $env:USERPROFILE
    $basePath = "C:\_Paycom Frontline Data Transfer"
    $sourceFolder = Join-Path -Path $basePath -ChildPath "Export_from_Paycom\Employees"
    $destinationFolder = Join-Path -Path $basePath -ChildPath "Import_to_Frontline\Employees"
    
    # Verify source folder exists
    Write-Log "Checking source folder: $sourceFolder" -Level Info
    if (-not (Test-Path -Path $sourceFolder)) {
        throw "Source folder not found: $sourceFolder`nPlease verify the path exists and you have access permissions."
    }
    Write-Log "Source folder verified" -Level Success
    
    # Verify destination folder exists, create if needed
    Write-Log "Checking destination folder: $destinationFolder" -Level Info
    if (-not (Test-Path -Path $destinationFolder)) {
        try {
            Write-Log "Creating destination folder..." -Level Info
            New-Item -ItemType Directory -Path $destinationFolder -Force -ErrorAction Stop | Out-Null
            Write-Log "Destination folder created successfully" -Level Success
        }
        catch {
            throw "Failed to create destination folder: $($_.Exception.Message)"
        }
    }
    else {
        Write-Log "Destination folder verified" -Level Success
    }
    
    # Get only the most recent CSV file in the folder
    Write-Log "Searching for CSV files in source folder..." -Level Info
    try {
        $paycomExport = Get-ChildItem -Path $sourceFolder -Filter *.csv -ErrorAction Stop | 
                        Sort-Object LastWriteTime -Descending | 
                        Select-Object -First 1
    }
    catch {
        throw "Error accessing source folder: $($_.Exception.Message)"
    }
    
    if ($null -eq $paycomExport) {
        throw "No CSV files found in: $sourceFolder`nPlease ensure CSV files are present in the source folder."
    }
    
    Write-Log "Found file: $($paycomExport.Name) (Last modified: $($paycomExport.LastWriteTime))" -Level Success
    Write-Log "File size: $([math]::Round($paycomExport.Length / 1KB, 2)) KB" -Level Info
    
    # Retrieve the last uploaded Paycom export file for comparison
    Write-Log "Searching for last uploaded Paycom export file..." -Level Info
    try {
        $lastPaycomExport = Get-ChildItem -Path "$($sourceFolder)\Completed_imports" -Filter *.csv -ErrorAction Stop | 
                        Sort-Object LastWriteTime -Descending | 
                        Select-Object -First 1
    }
    catch {
        throw "Error accessing last uploaded Paycom export file: $($_.Exception.Message)"
    }
    
    if ($null -eq $lastPaycomExport) {
        throw "No CSV files found in: $($sourceFolder)\Completed_imports`nPlease ensure a previously completed import file is present."
    }
    
    Write-Log "Found file: $($lastPaycomExport.Name) (Last modified: $($lastPaycomExport.LastWriteTime))" -Level Success
    Write-Log "File size: $([math]::Round($lastPaycomExport.Length / 1KB, 2)) KB" -Level Info
    
    # Compare the two CSV files to identify inactive employees
    Write-Log "Comparing current and last Paycom exports to identify inactive employees..." -Level Info
    $comparisonResults = C:\code\powershell_scripts\Paycom_Frontline\Compare-CSVFiles.ps1 -File1 $lastPaycomExport.FullName -File2 $paycomExport.FullName -KeyColumn "Identifier"

    $removedRows = $comparisonResults | Where-Object {$_.ChangeType -eq 'Removed'}

    Write-Log "CSV comparison completed" -Level Success
    Write-Log "Inactive employees identified: $($removedRows.Count)" -Level Info
    foreach ($emp in $removedRows) {
        Write-Log "Inactive Employee ID: $($emp.Identifier), Name: $($emp.'First Name') $($emp.'Last Name')" -Level Info
    }

    # Import the CSV data
    Write-Log "Importing CSV file..." -Level Info
    try {
        $data = Import-Csv -Path $paycomExport.FullName -ErrorAction Stop
    }
    catch {
        throw "Failed to import CSV file: $($_.Exception.Message)`nThe file may be corrupted, locked by another process, or not a valid CSV format."
    }
    
    if ($null -eq $data) {
        throw "Import returned null. The CSV file may be empty or improperly formatted."
    }
    
    if ($data.Count -eq 0) {
        throw "No data rows found in CSV file. The file appears to be empty or contains only headers."
    }
    
    Write-Log "Successfully imported $($data.Count) records" -Level Success

    # FIX 2: Build inactive rows as proper PSCustomObjects with all expected properties,
    # including 'Skill Code 1', before appending to $data. Previously, rows were added
    # via PSObject.Copy() without 'Skill Code 1', causing the transform block to fail or
    # silently produce blank values for those rows.
    foreach ($inactive in $removedRows) {
        $newRow = [PSCustomObject]@{
            'Identifier'          = $inactive.Identifier
            'First Name'          = $inactive.'First Name'
            'Last Name'           = $inactive.'Last Name'
            'Preferred_Firstname' = $inactive.'Preferred_Firstname'
            'Phone'               = $inactive.'Phone'
            'Work_Email'          = $inactive.'Work_Email'
            'Department'          = $inactive.'Department'
            'Department_Desc'     = $inactive.'Department_Desc'
            'Work_Location'       = $inactive.'Work_Location'
            'Position'            = $inactive.'Position'
            'Street 1'            = $inactive.'Street 1'
            'Street 2'            = $inactive.'Street 2'
            'City'                = $inactive.'City'
            'State'               = $inactive.'State'
            'Zip'                 = $inactive.'Zip'
            'Active'              = 'N'
            'Skill Code 1'        = ''   # Will remain blank; inactive rows skip campus mapping
        }
        $data += $newRow
        Write-Log "Marked Employee ID: $($inactive.Identifier) as Inactive" -Level Info
    }

    # Verify required columns exist
    Write-Log "Validating CSV structure..." -Level Info
    $requiredColumns = @('Identifier', 'First Name', 'Last Name', 'Phone','Work_Email','Department_Desc','Work_Location','Position')
    $missingColumns = @()
    
    $firstRow = $data | Select-Object -First 1
    foreach ($column in $requiredColumns) {
        if (-not ($firstRow.PSObject.Properties.Name -contains $column)) {
            $missingColumns += $column
        }
    }
    
    if ($missingColumns.Count -gt 0) {
        throw "Missing required columns: $($missingColumns -join ', ')`nPlease verify the CSV file format matches the expected Paycom export structure."
    }
    Write-Log "CSV structure validated successfully" -Level Success
    
    # Data cleaning and transformation
    Write-Log "Beginning data cleaning and transformation..." -Level Info
    $cleaningIssues = 0
    
    try {
        foreach ($row in $data) {
            $rowIdentifier = if ($row.Identifier) { $row.Identifier } else { "Unknown" }
            
            # If row does not have Active column, assume active
            if ($row.PSObject.Properties.Name -notcontains 'Active') {
                $row | Add-Member -MemberType NoteProperty -Name 'Active' -Value 'Y'
            }

            try {
                # If preferred name is not empty, use it as first name
                if (-not [string]::IsNullOrWhiteSpace($row.'Preferred_Firstname')) {
                    $originalFirstName = $row.'First Name'
                    $row.'First Name' = $row.'Preferred_Firstname'
                    Write-Log "Replaced First Name for ID $rowIdentifier : $originalFirstName -> $($row.'First Name')" -Level Info
                }
                
                # If zip code is longer than 5 characters, truncate it
                if ($row.'Zip'.Length -gt 5) {
                    $originalZip = $row.'Zip'
                    $row.'Zip' = $row.'Zip'.Substring(0,5)
                    Write-Log "Truncated Zip for ID $rowIdentifier : $originalZip -> $($row.'Zip')" -Level Info
                }

                # Standardize phone number format (remove non-numeric characters)
                $originalPhone = $row.'Phone'
                $row.'Phone' = ($row.'Phone' -replace '[^\d]', '')
                Write-Log "Standardized Phone for ID $rowIdentifier : $originalPhone -> $($row.'Phone')" -Level Info

                # FIX 3: Capture $originalCampus BEFORE the campus/skill mapping blocks,
                # so the Skill Code log message can reference it correctly.
                # Previously, $originalCampus was assigned after the skill code block,
                # meaning the log message always showed a null/empty value.
                $originalCampus = $row.'Work_Location'

                # Map Skill Code based on Work_Location keywords
                $skillCodeMapping = @{
                    'Senior'  = 'SS'
                    'Middle'  = 'MS'
                    'CDS'     = 'CDS'
                    'Junior'  = 'JS'
                }
                foreach ($key in $skillCodeMapping.Keys) {
                    if ($row.'Work_Location' -like "*$key*") {
                        # FIX 4: Guard Add-Member with a property existence check to prevent
                        # "duplicate member" errors if this property was already added in a
                        # prior iteration (e.g., on rows sourced from $removedRows).
                        if ($row.PSObject.Properties.Name -notcontains 'Skill Code 1') {
                            $row | Add-Member -MemberType NoteProperty -Name 'Skill Code 1' -Value ''
                        }
                        $row.'Skill Code 1' = $skillCodeMapping[$key]
                        Write-Log "Mapped Skill Code for ID $rowIdentifier : $originalCampus -> $($row.'Skill Code 1')" -Level Info
                        break
                    }
                }

                # Transform Campus/Work_Location to numeric campus code
                $campusMapping = @{
                    'Senior'  = '842'
                    'Middle'  = '843'
                    'CDS'     = '845'
                    'Junior'  = '844'
                }
                foreach ($key in $campusMapping.Keys) {
                    if ($row.'Work_Location' -like "*$key*") {
                        $row.'Work_Location' = $campusMapping[$key]
                        Write-Log "Mapped Campus for ID $rowIdentifier : $originalCampus -> $($row.'Work_Location')" -Level Info
                        break
                    }
                }

                # Map Department code to name using predefined mapping
                $originalDepartment = $row.'Department'
                $mappedDepartmentName = Get-DepartmentName -DepartmentCode $row.'Department'
                if ($mappedDepartmentName) {
                    $row.'Department' = $mappedDepartmentName
                    Write-Log "Mapped Department for ID $rowIdentifier : $originalDepartment -> $($row.'Department')" -Level Info
                }
                else {
                    Write-Log "Warning: Unrecognized Department code for ID $rowIdentifier : $originalDepartment" -Level Warning
                    $cleaningIssues++
                }

                # Validate critical fields
                if ([string]::IsNullOrWhiteSpace($row.'First Name')) {
                    Write-Log "Warning: Empty First Name for ID $rowIdentifier" -Level Warning
                    $cleaningIssues++
                }
                
                if ([string]::IsNullOrWhiteSpace($row.'Last Name')) {
                    Write-Log "Warning: Empty Last Name for ID $rowIdentifier" -Level Warning
                    $cleaningIssues++
                }

                if ([string]::IsNullOrWhiteSpace($row.'Work_Email')) {
                    Write-Log "Warning: Empty Email for ID $rowIdentifier" -Level Warning
                    $cleaningIssues++
                }
            }
            catch {
                Write-Log "Error processing row with ID $rowIdentifier : $($_.Exception.Message)" -Level Warning
                $cleaningIssues++
            }
        }
        
        Write-Log "Data cleaning completed. Issues found: $cleaningIssues" -Level $(if ($cleaningIssues -eq 0) { 'Success' } else { 'Warning' })
    }
    catch {
        throw "Critical error during data cleaning: $($_.Exception.Message)"
    }
    
    # Transform the data into Frontline format
    Write-Log "Transforming data to Frontline format..." -Level Info
    try {
        $transformedData = $data | ForEach-Object {
            [PSCustomObject]@{
                Identifier                  = $_.Identifier
                'First Name'                = $_.'First Name'
                'Last Name'                 = $_.'Last Name'
                'Phone'                     = $_.'Phone'
                'Pin'                       = ""
                'Email'                     = $_.'Work_Email'
                'Street 1'                  = $_.'Street 1'
                'Street 2'                  = $_.'Street 2'
                'City'                      = $_.'City'
                'State'                     = $_.'State'
                'Zip'                       = $_.'Zip'
                'Title'                     = $_.'Department'
                'Employee Type Description' = 'Teacher'
                'Substitute Requirements'   = 'O'
                'School External ID 1'      = $_.'Work_Location'
                'Can Cancel Absence'        = '2'
                'Skill Code 1'              = $_.'Skill Code 1'
                'Skill Type 1'              = 'MH'
                'Active'                    = $_.'Active'
            }
        }
        Write-Log "Successfully transformed $($transformedData.Count) records" -Level Success
    }
    catch {
        throw "Error during data transformation: $($_.Exception.Message)"
    }
    
    if ($null -eq $transformedData -or $transformedData.Count -eq 0) {
        throw "Transformation failed to produce any output data."
    }
    
    # Export the transformed data to a new CSV file
    Write-Log "Exporting transformed data..." -Level Info
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $destinationFile = Join-Path -Path $destinationFolder -ChildPath ("PaycomEmployees_P2F" + "_" + $timestamp + ".csv")
    
    if (Test-Path -Path $destinationFile) {
        Write-Log "Warning: Destination file already exists and will be overwritten: $destinationFile" -Level Warning
    }
    
    try {
        $transformedData | Export-Csv -Path $destinationFile -NoTypeInformation -ErrorAction Stop
        Write-Log "Transformed data exported to: $destinationFile" -Level Success
        Write-Log "Exported $($transformedData.Count) records" -Level Success
    }
    catch {
        throw "Failed to export transformed data: $($_.Exception.Message)`nThe destination folder may be read-only or you may lack write permissions."
    }
    
    if (-not (Test-Path -Path $destinationFile)) {
        throw "Export file was not created at expected location: $destinationFile"
    }
    
    $exportedFile = Get-Item -Path $destinationFile
    if ($exportedFile.Length -eq 0) {
        throw "Exported file is empty. Export may have failed."
    }
    Write-Log "Export file verified: $([math]::Round($exportedFile.Length / 1KB, 2)) KB" -Level Success
    
    # Move the original file to the Completed_imports folder
    Write-Log "Moving original file to completed folder..." -Level Info
    $completedFolder = Join-Path -Path $sourceFolder -ChildPath "Completed_imports"
    
    if (-not (Test-Path -Path $completedFolder)) {
        try {
            Write-Log "Creating completed imports folder: $completedFolder" -Level Info
            New-Item -ItemType Directory -Path $completedFolder -Force -ErrorAction Stop | Out-Null
            Write-Log "Completed imports folder created" -Level Success
        }
        catch {
            throw "Failed to create completed imports folder: $($_.Exception.Message)"
        }
    }
    
    try {
        $completedFilePath = Join-Path -Path $completedFolder -ChildPath $paycomExport.Name
        
        if (Test-Path -Path $completedFilePath) {
            $archiveTimestamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $newName = [System.IO.Path]::GetFileNameWithoutExtension($paycomExport.Name) + "_$archiveTimestamp" + [System.IO.Path]::GetExtension($paycomExport.Name)
            $completedFilePath = Join-Path -Path $completedFolder -ChildPath $newName
            Write-Log "File with same name exists. Renaming to: $newName" -Level Warning
        }
        
        Move-Item -Path $paycomExport.FullName -Destination $completedFilePath -ErrorAction Stop
        Write-Log "Original file moved to: $completedFilePath" -Level Success
    }
    catch {
        Write-Log "Warning: Failed to move original file to completed folder: $($_.Exception.Message)" -Level Warning
        Write-Log "The transformation was successful, but the source file remains in the original location." -Level Warning
    }
    
    # Calculate execution time
    $executionTime = (Get-Date) - $scriptStartTime
    Write-Log "Execution time: $($executionTime.TotalSeconds) seconds" -Level Info
    
    Write-Log "=== DATA TRANSFORMATION COMPLETED SUCCESSFULLY ===" -Level Success
    Write-Log "Summary: Processed $($data.Count) records, exported $($transformedData.Count) records, $cleaningIssues cleaning issues" -Level Success
    
    # Save log file
    Save-LogFile -DestinationFolder $destinationFolder
}
catch {
    Write-Log "=== CRITICAL ERROR ===" -Level Error
    Write-Log "Error Message: $($_.Exception.Message)" -Level Error
    Write-Log "Error Location: $($_.InvocationInfo.ScriptName):$($_.InvocationInfo.ScriptLineNumber)" -Level Error
    Write-Log "Stack Trace: $($_.ScriptStackTrace)" -Level Error
    
    $savedLogPath = $null
    if (Test-Path -Path $destinationFolder) {
        $savedLogPath = Save-LogFile -DestinationFolder $destinationFolder
    }
    
    Write-Log "Script failed. Please review the error details above." -Level Error

    if ($savedLogPath -and (Test-Path -Path $savedLogPath)) {
        Start-Process -FilePath $savedLogPath
    }

    exit 1
}

# End of script
