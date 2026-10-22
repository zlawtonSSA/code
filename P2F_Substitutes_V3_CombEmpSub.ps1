# Paycom To Frontline Substitutes Export Data Transformer
# This script transforms exported Substitute and Employee Data from Paycom into a format that will import succesfully via the Substitute template.

# Initialize error tracking
$ErrorActionPreference = 'Stop'
$scriptStartTime = Get-Date
$script:logMessages = [System.Collections.ArrayList]@()

# Function to write log messages with timestamp and severity level
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
        'Info' { Write-Host $logEntry -ForegroundColor Cyan }
        'Warning' { Write-Host $logEntry -ForegroundColor Yellow }
        'Error' { Write-Host $logEntry -ForegroundColor Red }
        'Success' { Write-Host $logEntry -ForegroundColor Green }
    }
}

# Function to save log messages to a file
function Save-LogFile {
    param([string]$destinationFolder)
    
    try {
        $logFileName = "P2F_Substitute_TransformLog_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
        $logFilePath = Join-Path -Path "C:\_Paycom Frontline Data Transfer\Export_from_Paycom\Substitutes\_logs" -ChildPath $logFileName
        $logMessages | Out-File -FilePath $logFilePath -Encoding UTF8
        Write-Log "Log file saved: $logFilePath" -Level Info
        return $logFilePath 
    }
    catch {
        Write-Log "Failed to save log file: $($_.Exception.Message)" -Level Warning
        return $null        
    }
}

try {
    Write-Log "=== Paycom to Frontline SUBSTITUTES Data Transformation Started ===" -Level Info
    
    # Get the script's directory
    $scriptPath = $PSScriptRoot
    if ([string]::IsNullOrEmpty($scriptPath)) {
        $scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
    }
    if ([string]::IsNullOrEmpty($scriptPath)) {
        $scriptPath = Get-Location
    }
    
    Write-Log "Script running from: $scriptPath" -Level Info
    
    # Define folder paths using environment variables and relative paths
    # Try to use OneDrive path if available, otherwise use local Desktop
    $userProfile = $env:USERPROFILE
    $basePath = "C:\_Paycom Frontline Data Transfer"
    $sourceFolderSubs = Join-Path -Path $basePath -ChildPath "Export_from_Paycom\Substitutes"
    $destinationFolder = Join-Path -Path $basePath -ChildPath "Import_to_Frontline\Substitutes"
    
    # Verify source folder exists
    Write-Log "Checking source folder: $sourceFolderSubs" -Level Info
    if (-not (Test-Path -Path $sourceFolderSubs)) {
        throw "Source folder not found: $sourceFolderSubs`nPlease verify the path exists and you have access permissions."
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

    ####################################################################################
    # --- Get the most recent SUBSTITUTES CSV file --- #

    Write-Log "Searching for SUBSTITUTES CSV files in source folder..." -Level Info
    try {
        $paycomSubsExport = Get-ChildItem -Path $sourceFolderSubs -Filter 'P2F-SUBSTITUTES_COMBINED*.csv' -ErrorAction Stop | 
        Sort-Object LastWriteTime -Descending | 
        Select-Object -First 1
    }
    catch {
        throw "Error accessing source folder: $($_.Exception.Message)"
    }
    
    # Check if a file was found
    if ($null -eq $paycomSubsExport) {
        throw "No CSV files found in: $sourceFolderSubs`nPlease ensure CSV files are present in the source folder."
    }
    
    Write-Log "Found file: $($paycomSubsExport.Name) (Last modified: $($paycomSubsExport.LastWriteTime))" -Level Success
    Write-Log "File size: $([math]::Round($paycomSubsExport.Length / 1KB, 2)) KB" -Level Info
    
    #########################################################################################
    # --- Import the SUBSTITUTE CSV file --- #

    Write-Log "Importing SUBSTITUTE CSV file..." -Level Info
    try {
        $subData = Import-Csv -Path $paycomSubsExport.FullName -ErrorAction Stop
    }
    catch {
        throw "Failed to import CSV file: $($_.Exception.Message)`nThe file may be corrupted, locked by another process, or not a valid CSV format."
    }
    
    # Check if subData was successfully imported
    if ($null -eq $subData) {
        throw "Import returned null. The CSV file may be empty or improperly formatted."
    }
    if ($subData.Count -eq 0) {
        throw "No subData rows found in CSV file. The file appears to be empty or contains only headers."
    }
    
    Write-Log "Successfully imported $($subData.Count) substitute records" -Level Success

    # Verify required columns exist in the subData
    Write-Log "Validating CSV structure..." -Level Info
    $requiredColumns = @('Identifier', 'First Name', 'Preferred_Firstname', 'Last Name', 'Phone', 'Email', 'Work_Email', 'Street 1', 'City', 'State', 'Zip', 'Skill Code 1', 'Skill Code 2', 'Skill Code 3', 'Skill Code 4', 'Pay_Type')
    $missingColumns = @()

    # Check the first row for column names to avoid iterating through all rows
    $firstRow = $subData | Select-Object -First 1
    foreach ($column in $requiredColumns) {
        if (-not ($firstRow.PSObject.Properties.Name -contains $column)) {
            $missingColumns += $column
        }
    }
    
    if ($missingColumns.Count -gt 0) {
        throw "Missing required columns: $($missingColumns -join ', ')`nPlease verify the CSV file format matches the expected Paycom export structure."
    }
    Write-Log "CSV structure validated successfully" -Level Success
    
    #################################################################################
    # --- Data cleaning and transformation --- #
    
    Write-Log "Beginning subData cleaning and transformation..." -Level Info
    $cleaningIssues = 0
    $transformedCount = 0
    
    try {
        foreach ($row in $subData) {
            # Track row for error reporting
            $rowIdentifier = if ($row.Identifier) { $row.Identifier } else { "Unknown" }
            
            try {
                #if preferred name is not empty, use it as first name
                if (-not [string]::IsNullOrWhiteSpace($row.'Preferred_Firstname')) {
                    $originalFirstName = $row.'First Name'
                    $row.'First Name' = $row.'Preferred_Firstname'
                    Write-Log "Replaced First Name for ID $rowIdentifier : $originalFirstName -> $($row.'First Name')" -Level Info
                }
                
                # If zip code is longer than 5 characters, truncate it
                if ($row.'Zip'.Length -gt 5) {
                    $originalZip = $row.'Zip'
                    $row.'Zip' = $row.'Zip'.Substring(0, 5)
                    Write-Log "Truncated Zip for ID $rowIdentifier : $originalZip -> $($row.'Zip')" -Level Info
                }

                # Standardize phone number format (remove non-numeric characters)
                $originalPhone = $row.'Phone'
                $row.'Phone' = ($row.'Phone' -replace '[^\d]', '')
                Write-Log "Standardized Phone for ID $rowIdentifier : $originalPhone -> $($row.'Phone')" -Level Info
                
                # Validate critical fields
                if ([string]::IsNullOrWhiteSpace($row.'First Name')) {
                    Write-Log "Warning: Empty First Name for ID $rowIdentifier" -Level Warning
                    $cleaningIssues++
                }
                if ([string]::IsNullOrWhiteSpace($row.'Last Name')) {
                    Write-Log "Warning: Empty Last Name for ID $rowIdentifier" -Level Warning
                    $cleaningIssues++
                }

                # Validate email - if both Email and Work_Email are empty, log a warning since at least one email should be present for contact purposes
                if ([string]::IsNullOrWhiteSpace($row.'Email') -and [string]::IsNullOrWhiteSpace($row.'Work_Email')) {
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
        Write-Log "Critical error during subData cleaning: $($_.Exception.Message)" -Level Error
        throw "Critical error during subData cleaning: $($_.Exception.Message)"
    }
    
    ###########################################################################################
    # --- Transform the subData --- #

    Write-Log "Transforming subData to Frontline format..." -Level Info
    
    # The transformation maps fields from the Paycom export to the required Frontline format, and applies logic for certain fields based on Pay Type
    try {
        $transformedData = $subData | ForEach-Object {
            $transformedCount++
            #$email = if ([string]::IsNullOrWhiteSpace($_.'Email')) { $_.'Work_Email' } else { $_.'Email' }
            
            [PSCustomObject]@{
                Identifier     = $_.Identifier
                'First Name'   = $_.'First Name'
                'Last Name'    = $_.'Last Name'
                'Phone'        = $_.'Phone'
                'External ID'  = if ($_.'Pay_Type' -eq 'Hourly') { 'External Substitute' } else { 'Internal Substitute' }
                'Email'        = if ($_.'Pay_Type' -eq 'Hourly') { $_.Email } else { $_.Work_Email }
                'Street 1'     = $_.'Street 1'
                'Street 2'     = $_.'Street 2'
                'City'         = $_.'City'
                'State'        = $_.'State'
                'Zip'          = $_.'Zip'
                'CanCancel'    = if ($_.'Pay_Type' -eq 'Hourly') { 'Y' } else { 'N' }
                'Skill Code 1' = $_.'Skill Code 1'
                'Skill Code 2' = $_.'Skill Code 2'
                'Skill Code 3' = $_.'Skill Code 3'
                'Skill Code 4' = $_.'Skill Code 4'
                'Active'       = 'Y'
            }
        }
        Write-Log "Successfully transformed $transformedCount records" -Level Success
    }
    catch {
        throw "Error during subData transformation: $($_.Exception.Message)"
    }
    
    # Verify transformation produced subData
    if ($null -eq $transformedData -or $transformedData.Count -eq 0) {
        throw "Transformation failed to produce any output subData."
    }

    # Export the transformed subData to a temporary CSV file for review
    Write-Log "Exporting transformed subData..." -Level Info
    
    # Check if destination folder is writable before attempting export
    try {
        $tempFolder = Join-Path -Path $destinationFolder -ChildPath "temp"
        $tempFile = Join-Path -Path $tempFolder -ChildPath ("temp_substitutes_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".csv")
       
        #export to temp folder first for review before exporting to final destination
        if (-not (Test-Path -Path $tempFolder)) {
            Write-Log "Creating temp folder for intermediate export: $tempFolder" -Level Info
            New-Item -ItemType Directory -Path $tempFolder -Force -ErrorAction Stop | Out-Null
            Write-Log "Temp folder created successfully" -Level Success
        }
        else {
            Write-Log "Temp folder already exists: $tempFolder" -Level Info
        }
        
        $transformedData | Export-Csv -Path $tempFile -NoTypeInformation -ErrorAction Stop

        Write-Log "Transformed subData exported to: $tempFile" -Level Success
        Write-Log "Exported $($transformedData.Count) records to temp file" -Level Success
    }
    catch {
        throw "Failed to export transformed subData: $($_.Exception.Message)`nThe destination folder may be read-only or you may lack write permissions."
    }

    ################################################################################################
    ### --- Check for active and inactive substitutes --- ###
    
    # Retrieve the last uploaded Paycom export file so it can be compared against the new file
    Write-Log "Searching for last uploaded Substitutes export file..." -Level Info
    try {
        $lastSubsExport = Get-ChildItem -Path $destinationFolder -Filter *.csv -File -ErrorAction Stop |
        Where-Object { $_.DirectoryName -notlike "*\temp*" } |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
    }
    catch {
        throw "Error accessing last uploaded Substitutes file: $($_.Exception.Message)"
    }
    
    # Check if a previously uploaded file was found
    if ($null -eq $lastSubsExport) {
        throw "No CSV files found in: $destinationFolder`nPlease ensure CSV files are present in the source folder."
    }
    
    Write-Log "Found file: $($lastSubsExport.Name) (Last modified: $($lastSubsExport.LastWriteTime))" -Level Success
    Write-Log "File size: $([math]::Round($lastSubsExport.Length / 1KB, 2)) KB" -Level Info
    
    # Compare the two CSV files to identify inactive employees
    Write-Log "Comparing current and last Substitutes exports to identify inactive employees..." -Level Info
    $comparisonResults = C:\code\powershell_scripts\Paycom_Frontline\Compare-CSVFiles.ps1 -File1 $lastSubsExport.FullName -File2 $tempFile -KeyColumn "Identifier"
    
    # The comparison results will include all added, removed, and modified records with a ChangeType property indicating the type of change. 
    # We will use this to identify which employees need to be marked as inactive in Frontline since Frontline does not allow for deletion of records and requires an active/inactive status instead.
    $removedRows = $comparisonResults | Where-Object {$_.ChangeType -eq 'Removed'}
    $modifiedRows = $comparisonResults | Where-Object {$_.ChangeType -eq 'Modified'}

    # Check if comparison results are null or empty, which could indicate an issue with the comparison script or that there are no changes between the files. Log a warning if this is the case.
    if ($null -eq $comparisonResults) {
        Write-Log "Warning: CSV comparison returned no results or failed silently" -Level Warning
        $comparisonResults = @()
    }

    Write-Log "CSV comparison completed" -Level Success
    Write-Log "Inactive employees identified: $($removedRows.Count)" -Level Info
    Write-Log "Modified employees identified: $($modifiedRows.Count)" -Level Info
    # Log details of removed employees for review
    foreach ($emp in $removedRows) {
        Write-Log "Inactive Employee ID: $($emp.Identifier), Name: $($emp.'First Name') $($emp.'Last Name')" -Level Info
    }

    # Add the inactive employees to the transformed data with status 'Inactive'
    $transformedData = [System.Collections.Generic.List[PSCustomObject]]$transformedData
    foreach ($inactive in $removedRows) {
        $newRow = $inactive.PSObject.Copy()
        $newRow.'Active' = 'N'
        $transformedData.Add($newRow)
        Write-Log "Marked Employee ID: $($inactive.Identifier) as Inactive" -Level Info
    }

    # update modified employees in the transformed data with any changes from the comparison results
    foreach ($modified in $modifiedRows) {
        $existingRow = $transformedData | Where-Object { $_.Identifier -eq $modified.Identifier }
        if ($existingRow) {
            # Update the existing row with the modified values, excluding the ChangeType and ChangedFields properties.
            # if the header column does not exist, add it to the existing row and set the value to the modified value
            foreach ($property in $modified.PSObject.Properties) {
                if ($property.Name -notin 'ChangeType', 'ChangedFields') {
                    if ($existingRow.PSObject.Properties.Name -contains $property.Name) {
                        $existingRow.$($property.Name) = $property.Value
                    }
                    else {
                        $existingRow | Add-Member -MemberType NoteProperty -Name $property.Name -Value $property.Value -Force
                    }
                }
            }
            Write-Log "Updated Employee ID: $($modified.Identifier) with changes from comparison" -Level Info
        }
    }

    ################################################################################################
    # --- Export the transformed subData to a new CSV file in the destination folder --- #

    # Update the exported file name to include a timestamp for versioning
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $exportedFileName = "PaycomSubstitutes" + "_$timestamp" + ".csv"

    # Export the transformed subData to a new CSV file
    Write-Log "Exporting transformed subData..." -Level Info
    $destinationFile = Join-Path -Path $destinationFolder -ChildPath ($exportedFileName)
    
    # Check if destination file already exists
    if (Test-Path -Path $destinationFile) {
        Write-Log "Warning: Destination file already exists and will be overwritten: $destinationFile" -Level Warning
    }
    # Export to destination folder
    try {
        $transformedData | Select-Object Identifier, 
                                 'First Name', 
                                 'Last Name', 
                                 'Phone', 
                                 'External ID', 
                                 'Email', 
                                 'Street 1', 
                                 'Street 2', 
                                 'City', 
                                 'State', 
                                 'Zip', 
                                 'CanCancel', 
                                 'Skill Code 1', 
                                 'Skill Code 2', 
                                 'Skill Code 3', 
                                 'Skill Code 4', 
                                 'Remove Skill 1',
                                 'Remove Skill 2',
                                 'Remove Skill 3',
                                 'Remove Skill 4',
                                 'Active' |
                   Sort-Object -Property 'Active','External ID','Identifier' | 
                   Export-Csv -Path $destinationFile -NoTypeInformation
        Write-Log "Transformed subData exported to: $destinationFile" -Level Success
        Write-Log "Exported $($transformedData.Count) records" -Level Success
    }
    catch {
        throw "Failed to export transformed subData: $($_.Exception.Message)`nThe destination folder may be read-only or you may lack write permissions."
    }
    
    # Verify the export file was created and has content
    if (-not (Test-Path -Path $destinationFile)) {
        throw "Export file was not created at expected location: $destinationFile"
    }
    
    # Verify the exported file is not empty
    $exportedFile = Get-Item -Path $destinationFile

    # Check if file size is greater than 0 bytes
    if ($exportedFile.Length -eq 0) {
        throw "Exported file is empty. Export may have failed."
    }
    Write-Log "Export file verified: $([math]::Round($exportedFile.Length / 1KB, 2)) KB" -Level Success
    
    # Move the original file to the Completed_imports folder
    Write-Log "Moving original file to completed folder..." -Level Info
    $archiveFolder = Join-Path "C:\_Paycom Frontline Data Transfer\Export_from_Paycom\Substitutes\Completed_imports" (Get-Date -Format "yyyy-MM-dd")

    # test if archive folder exists, throw error if it does not exist and leave file in original location so it can be manually moved to avoid data loss, if it does exist move the file to the archive folder
    if (-not (Test-Path -Path $archiveFolder)) {
        Write-Log "Archive folder does not exist: $archiveFolder" -Level Warning
        Write-Log "The source file will remain in the original location: $($paycomSubsExport.FullName)" -Level Warning
        Write-Log "Please create the archive folder and move the source file to avoid clutter and for better organization." -Level Warning
    }
    else {
        try {
            Move-Item -Path $paycomSubsExport.FullName -Destination $archiveFolder -ErrorAction Stop
            Write-Log "Moved source file: '$paycomSubsExport.FullName' -> '$archiveFolder'" -Level Info
        } 
        catch {
            Write-Log "Failed to move source file '$paycomSubsExport.FullName': $($_.Exception.Message)" -Level Warning
        }
    }
    
    # Calculate execution time
    $executionTime = (Get-Date) - $scriptStartTime
    Write-Log "Execution time: $($executionTime.TotalSeconds) seconds" -Level Info
    
    Write-Log "=== DATA TRANSFORMATION COMPLETED SUCCESSFULLY ===" -Level Success
    Write-Log "Summary: Processed $($subData.Count) records, exported $($transformedData.Count) records, $cleaningIssues cleaning issues" -Level Success
    
    # Save log file
    Save-LogFile -DestinationFolder $destinationFolder
}
catch {
    # Log the error details with severity level and save the log file before exiting
    Write-Log "=== CRITICAL ERROR ===" -Level Error
    Write-Log "Error Message: $($_.Exception.Message)" -Level Error
    Write-Log "Error Location: $($_.InvocationInfo.ScriptName):$($_.InvocationInfo.ScriptLineNumber)" -Level Error
    
    # Save log file even on error
    $savedLogPath = $null
    if (Test-Path -Path $destinationFolder) {
        $savedLogPath = Save-LogFile -DestinationFolder $destinationFolder
    }
    
    Write-Log "Script failed. Please review the error details above." -Level Error

    # Open the log file automatically if it was saved
    if ($savedLogPath -and (Test-Path -Path $savedLogPath)) {
        Start-Process -FilePath $savedLogPath
    }

    exit 1
}

# End of script