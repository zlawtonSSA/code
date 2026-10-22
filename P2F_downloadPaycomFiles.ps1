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

        $logFileName = "P2F_downloadPaycomFiles_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
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

# Function to save log file and open it if an error occurs
function Save-AndOpenLogOnError {
    $savedLogPath = Save-LogFile
    if ($savedLogPath -and (Test-Path -Path $savedLogPath)) {
        Write-Log "Opening log file..." -Level Info
        Start-Process -FilePath $savedLogPath
    }
    else {
        Write-Log "Could not open log file - path unavailable." -Level Warning
    }
}

try {
    Write-Log "=== Paycom File Download Started ===" -Level Info

    # --- Load WinSCP Assembly ---
    Write-Log "Loading WinSCP .NET assembly..." -Level Info
    $dllPath = "C:\Program Files (x86)\WinSCP\netstandard2.0\WinSCPnet.dll"

    if (-not (Test-Path -Path $dllPath)) {
        throw "WinSCP DLL not found at: $dllPath`nPlease verify WinSCP is installed."
    }

    try {
        [System.Reflection.Assembly]::LoadFrom($dllPath) | Out-Null
        Write-Log "WinSCP assembly loaded successfully" -Level Success
    }
    catch {
        throw "Failed to load WinSCP assembly: $($_.Exception.Message)"
    }

    # --- Configure Session ---
    Write-Log "Configuring SFTP session..." -Level Info
    try {
        $sessionOptions = New-Object WinSCP.SessionOptions -Property @{
            Protocol              = [WinSCP.Protocol]::Sftp
            HostName              = "push.paycomonline.net"
            UserName              = "000_0vh13_frontline"
            Password              = "u`$9Tn=0Yx-7Yb)2K"
            SshHostKeyFingerprint = "ssh-rsa 2048 DOd8qz6x5/fu7FAdsuWXhsRRMseCMt4GowBatRxdAdQ"
        }
        $sessionOptions.AddRawSettings("ProxyPort", "0")
        Write-Log "Session options configured" -Level Success
    }
    catch {
        throw "Failed to configure session options: $($_.Exception.Message)"
    }

    # --- Verify Destination Folders Exist ---
    $employeeDestFolder   = "C:\_Paycom Frontline Data Transfer\Export_from_Paycom\Employees"
    $substituteDestFolder = "C:\_Paycom Frontline Data Transfer\Export_from_Paycom\Substitutes"

    foreach ($folder in @($employeeDestFolder, $substituteDestFolder)) {
        if (-not (Test-Path -Path $folder)) {
            try {
                Write-Log "Creating destination folder: $folder" -Level Info
                New-Item -ItemType Directory -Path $folder -Force -ErrorAction Stop | Out-Null
                Write-Log "Folder created: $folder" -Level Success
            }
            catch {
                throw "Failed to create destination folder '$folder': $($_.Exception.Message)"
            }
        }
        else {
            Write-Log "Destination folder verified: $folder" -Level Info
        }
    }

    # --- Open Session and Download Files ---
    Write-Log "Connecting to SFTP server..." -Level Info
    $session = New-Object WinSCP.Session
    $session.ExecutablePath = "C:\Program Files (x86)\WinSCP\WinSCP.exe"

    try {
        try {
            $session.Open($sessionOptions)
            Write-Log "Connected to SFTP server successfully" -Level Success
        }
        catch {
            throw "Failed to connect to SFTP server: $($_.Exception.Message)"
        }

        # --- List Remote Directory ---
        Write-Log "Listing remote directory /Outbound..." -Level Info
        try {
            $directoryInfo = $session.ListDirectory("/Outbound")
            Write-Log "Remote directory listed. Files found: $($directoryInfo.Files.Count)" -Level Success
        }
        catch {
            throw "Failed to list remote directory '/Outbound': $($_.Exception.Message)"
        }

        # --- Find Latest Employee File ---
        Write-Log "Searching for latest Employee export file..." -Level Info
        $latestEmployeeFile = $directoryInfo.Files |
            Where-Object { $_.Name -like "*Paycom Export for Frontline - EMPLOYEES*.csv" } |
            Sort-Object LastWriteTime -Descending |
            Select-Object -First 1

        # Note: Adjust the file pattern as needed to match the actual naming convention of the Employee export files
        if ($latestEmployeeFile) {
            Write-Log "Found Employee file: $($latestEmployeeFile.Name) (Last modified: $($latestEmployeeFile.LastWriteTime))" -Level Success
            $remotePath = "/Outbound/$($latestEmployeeFile.Name)"
            $localDest  = "$employeeDestFolder\*"
            
            # The local destination is set to the folder with a wildcard to allow WinSCP to determine the filename from the remote file
            try {
                Write-Log "Downloading Employee file from $remotePath to $employeeDestFolder..." -Level Info
                $session.GetFiles($remotePath, $localDest).Check()
                Write-Log "Employee file downloaded successfully" -Level Success
            }
            catch {
                throw "Failed to download Employee file '$($latestEmployeeFile.Name)': $($_.Exception.Message)"
            }
        }
        else {
            Write-Log "No matching Employee files found in /Outbound" -Level Warning
        }

        #############################################################################################
        # # --- Find Latest Substitute File ---
        # Write-Log "Searching for latest Substitute export file..." -Level Info
        # $latestSubstituteFile = $directoryInfo.Files |
        #     Where-Object { $_.Name -like "*Paycom Export for Frontline - SUBSTITUTES*.csv" } |
        #     Sort-Object LastWriteTime -Descending |
        #     Select-Object -First 1

        # if ($latestSubstituteFile) {
        #     Write-Log "Found Substitute file: $($latestSubstituteFile.Name) (Last modified: $($latestSubstituteFile.LastWriteTime))" -Level Success
        #     $remotePath = "/Outbound/$($latestSubstituteFile.Name)"
        #     $localDest  = "$substituteDestFolder\*"

        #     try {
        #         Write-Log "Downloading Substitute file from $remotePath to $substituteDestFolder..." -Level Info
        #         $session.GetFiles($remotePath, $localDest).Check()
        #         Write-Log "Substitute file downloaded successfully" -Level Success
        #     }
        #     catch {
        #         throw "Failed to download Substitute file '$($latestSubstituteFile.Name)': $($_.Exception.Message)"
        #     }
        # }
        # else {
        #     Write-Log "No matching Substitute files found in /Outbound" -Level Warning
        # }
        ############################################################################

        # --- Find Latest Substitute Files ---
        Write-Log "Searching for latest Substitute export files..." -Level Info

        $substituteFilePatterns = @(
            "SUBSTITUTES SS",
            "SUBSTITUTES MS",
            "SUBSTITUTES CDS",
            "SUBSTITUTES JS"
        )
        $downloadedFiles = @()

        # Loop through each pattern to find and download the latest matching Substitute file
        foreach ($pattern in $substituteFilePatterns) {
            Write-Log "Searching for '$pattern' export file..." -Level Info
            # Adjust the file pattern as needed to match the actual naming convention of the Substitute export files
            $latestSubstituteFile = $directoryInfo.Files |
                Where-Object { $_.Name -like "*Paycom Export for Frontline - $pattern*.csv" } |
                Sort-Object LastWriteTime -Descending |
                Select-Object -First 1

            # Note: The pattern is included in the file name search to ensure we are looking for the correct type of Substitute file
            if ($latestSubstituteFile) {
                Write-Log "Found Substitute file: $($latestSubstituteFile.Name) (Last modified: $($latestSubstituteFile.LastWriteTime))" -Level Success
                $remotePath = "/Outbound/$($latestSubstituteFile.Name)"
                $localDest = "$substituteDestFolder\*"
                
                # The local destination is set to the folder with a wildcard to allow WinSCP to determine the filename from the remote file
                try {
                    Write-Log "Downloading Substitute file from $remotePath to $substituteDestFolder..." -Level Info
                    $session.GetFiles($remotePath, $localDest).Check()
                    Write-Log "Substitute file '$($latestSubstituteFile.Name)' downloaded successfully" -Level Success
                    $downloadedFiles += "$substituteDestFolder\$($latestSubstituteFile.Name)"
                }
                catch {
                    throw "Failed to download Substitute file '$($latestSubstituteFile.Name)': $($_.Exception.Message)"
                }
            }
            else {
                Write-Log "No matching Substitute files found for pattern '$pattern' in /Outbound" -Level Warning
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

            # The output file name includes a timestamp to ensure uniqueness and prevent overwriting previous combined files
            $combinedOutputPath = "$substituteDestFolder\P2F-SUBSTITUTES_COMBINED_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

            # We export the deduplicated collection of rows to a new CSV file in the Substitute destination folder, using a tab delimiter to match the original file format
            try {
                $deduplicatedRows | Export-Csv -Path $combinedOutputPath -NoTypeInformation -Delimiter ","
                Write-Log "Combined Substitute file saved to '$combinedOutputPath'" -Level Success
            }
            catch {
                throw "Failed to save combined Substitute file: $($_.Exception.Message)"
            }

            # # --- Delete Downloaded Source Files After Combining ---
            # Write-Log "Cleaning up $($downloadedFiles.Count) downloaded source file(s)..." -Level Info
            # foreach ($file in $downloadedFiles) {
            #     try {
            #         Remove-Item -Path $file -Force -ErrorAction Stop
            #         Write-Log "Deleted source file: '$file'" -Level Info
            #     } catch {
            #         Write-Log "Failed to delete source file '$file': $($_.Exception.Message)" -Level Warning
            #     }
            # }
            # Write-Log "Cleanup complete." -Level Info

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
    }
    finally {
        # Always disconnect cleanly, even if an error occurred
        if ($session -and $session.Opened) {
            Write-Log "Disconnecting from SFTP server..." -Level Info
            $session.Dispose()
            Write-Log "Disconnected from SFTP server" -Level Info
        }
    }

    # --- Summary ---
    $executionTime = (Get-Date) - $scriptStartTime
    Write-Log "Execution time: $([math]::Round($executionTime.TotalSeconds, 2)) seconds" -Level Info
    Write-Log "=== DOWNLOAD COMPLETED SUCCESSFULLY ===" -Level Success
    
    # Save the log file at the end of the script execution to capture all messages, including the summary and success status
    Save-LogFile | Out-Null
}
catch {
    # Log the error details with severity level and save the log file before exiting
    Write-Log "=== CRITICAL ERROR ===" -Level Error
    Write-Log "Error Message: $($_.Exception.Message)" -Level Error
    Write-Log "Error Location: $($_.InvocationInfo.ScriptName):$($_.InvocationInfo.ScriptLineNumber)" -Level Error
    Write-Log "Script failed. Please review the error details above." -Level Error
    
    # Attempt to save the log file and open it for review when an error occurs, providing immediate access to the error details for troubleshooting
    Save-AndOpenLogOnError

    exit 1
}