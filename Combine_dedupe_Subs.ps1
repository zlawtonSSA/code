# Initialize error tracking
$ErrorActionPreference = 'Stop'
$scriptStartTime = Get-Date
$script:logMessages = [System.Collections.ArrayList]@()

# Define logging function to capture messages with timestamps and severity levels
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

# Function to save log messages to a file with error handling
function Save-LogFile {
    try {
        $logDir = "C:\_Paycom Frontline Data Transfer\Export_from_Paycom\_logs"

        # Create log directory if it doesn't exist
        if (-not (Test-Path -Path $logDir)) {
            New-Item -ItemType Directory -Path $logDir -Force -ErrorAction Stop | Out-Null
        }

        $logFileName = "combine_dedupe_subs_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
        $logFilePath = Join-Path -Path $logDir -ChildPath $logFileName
        $script:logMessages | Out-File -FilePath $logFilePath -Encoding UTF8
        Write-Log "Log file saved: $logFilePath" -Level Info
        return $logFilePath
    }
    catch {
        Write-Log "Failed to save log file: $($_.Exception.Message)" -Level Warning
        return $null
    }
} 
 
 # --- Find Latest Substitute Files ---
        Write-Log "Searching for latest Substitute export files..." -Level Info

        $substituteFilePatterns = @(
            "SUBSTITUTES SS",
            "SUBSTITUTES MS",
            "SUBSTITUTES CDS",
            "SUBSTITUTES JS"
        )
        $downloadedFiles = @()

        $paycomExportFolder = "C:\_Paycom Frontline Data Transfer\Export_from_Paycom\Substitutes"
        # We use Get-ChildItem to retrieve the list of files in the Paycom export folder, which allows us to filter and sort them based on their names and modification dates
        $directoryInfo = Get-ChildItem -Path $paycomExportFolder -File

        # Loop through each pattern to find and download the latest matching Substitute file
        foreach ($pattern in $substituteFilePatterns) {
            Write-Log "Searching for '$pattern' export file..." -Level Info
            # Adjust the file pattern as needed to match the actual naming convention of the Substitute export files
            $latestSubstituteFile = $directoryInfo |
                Where-Object { $_.Name -like "*$pattern*.csv" } |
                Sort-Object LastWriteTime -Descending |
                Select-Object -First 1

            # Note: The pattern is included in the file name search to ensure we are looking for the correct type of Substitute file
            if ($latestSubstituteFile) {
                Write-Log "Found Substitute file: $($latestSubstituteFile.Name) (Last modified: $($latestSubstituteFile.LastWriteTime))" -Level Success
    
                # We add the full path of the found file to the list of downloaded files for later processing
                $downloadedFiles += $latestSubstituteFile.FullName
            }
            else {
                Write-Log "No matching Substitute files found for pattern '$pattern'" -Level Warning
            }
        }

        # --- Combine and Deduplicate Substitute Files ---
        if ($downloadedFiles.Count -gt 0) {
            Write-Log "Combining $($downloadedFiles.Count) Substitute file(s) and removing duplicates by Identifier..." -Level Info
            
            $combinedRows = @()

            # Loop through each downloaded Substitute file, read its contents, and combine the rows into a single collection
            foreach ($file in $downloadedFiles) {
                try {
                    $rows = Import-Csv -Path $file -Delimiter "," -ErrorAction Stop
                    Write-Log "Read $($rows.Count) rows from '$file'" -Level Info
                    $combinedRows += $rows
                }
                catch {
                    throw "Failed to read Substitute file '$file': $($_.Exception.Message)"
                }
            }

            # After combining all rows from the downloaded Substitute files, we perform deduplication based on the 'Identifier' column
            $deduplicatedRows = $combinedRows | Group-Object -Property Identifier | ForEach-Object { $_.Group | Select-Object -First 1 }

            Write-Log "Combined row count: $($combinedRows.Count) | After deduplication: $($deduplicatedRows.Count)" -Level Info
            
            $substituteDestFolder = "C:\_Paycom Frontline Data Transfer\Export_from_Paycom\Substitutes\"

            # The output file name includes a timestamp to ensure uniqueness and prevent overwriting previous combined files
            $combinedOutputPath = "$substituteDestFolder\P2F-SUBSTITUTES_COMBINED_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

            # We export the deduplicated collection of rows to a new CSV file in the Substitute destination folder, using a tab delimiter to match the original file format
            try {
                # sort the deduplicated rows by Identifier before exporting to ensure consistent ordering in the output file
                $deduplicatedRows = $deduplicatedRows | Sort-Object -Property 'Pay_Type', 'Identifier'
                $deduplicatedRows | Export-Csv -Path $combinedOutputPath -NoTypeInformation -Delimiter "," -Force -ErrorAction Stop
                Write-Log "Combined Substitute file saved to '$combinedOutputPath'" -Level Success
            }
            catch {
                throw "Failed to save combined Substitute file: $($_.Exception.Message)"
            }

            # --- Move Downloaded Source Files to Dated Archive Folder ---
            Write-Log "Archiving $($downloadedFiles.Count) downloaded source file(s)..." -Level Info

            $archiveFolder = Join-Path "C:\_Paycom Frontline Data Transfer\Export_from_Paycom\Substitutes\Completed_imports" (Get-Date -Format "yyyy-MM-dd")

            if (-not (Test-Path $archiveFolder)) {
                New-Item -ItemType Directory -Path $archiveFolder | Out-Null
                Write-Log "Created archive folder: '$archiveFolder'" -Level Info
            }

            foreach ($file in $downloadedFiles) {
                try {
                    Move-Item -Path $file -Destination $archiveFolder -Force -ErrorAction Stop
                    Write-Log "Moved source file: '$file' -> '$archiveFolder'" -Level Info
                } catch {
                    Write-Log "Failed to move source file '$file': $($_.Exception.Message)" -Level Warning
                }
            }
            Write-Log "Archiving complete." -Level Info


        }
        else {
            Write-Log "No Substitute files were downloaded; skipping combine step." -Level Warning
        }