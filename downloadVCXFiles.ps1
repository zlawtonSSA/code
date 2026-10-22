
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
        $logDir = "C:\_ParentSquare\_logs"

        # Create log directory if it doesn't exist
        if (-not (Test-Path -Path $logDir)) {
            New-Item -ItemType Directory -Path $logDir -Force -ErrorAction Stop | Out-Null
        }

        $logFileName = "ParentSquareLog_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
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

# Main script execution with error handling
try {
    Write-Log "=== STFPGO File Download Started ===" -Level Info

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
            HostName              = "sha-mjbbhxbwrcpb.softwareng.cloud"
            UserName              = "zlawtonSFTPGo"
            Password              = "ftb_rfb0wfc7gyr0FBQ"
            SshHostKeyFingerprint = "ssh-ed25519 255 rtRFg0EEJjPQP3alJwCOYa2r+cFcWRJ1WpihL0MCKv4"
        }
        $sessionOptions.AddRawSettings("ProxyPort", "0")
        Write-Log "Session options configured" -Level Success
    }
    catch {
        throw "Failed to configure session options: $($_.Exception.Message)"
    }

    # --- Verify Destination Folders Exist ---
    $parentSquareBaseFolder = "C:\_ParentSquare\Unprocessed"

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
        Write-Log "Listing remote directory /ParentSquare/ExportsFromVeracross..." -Level Info
        try {
            $directoryInfo = $session.ListDirectory("/ParentSquare/ExportsFromVeracross")
            Write-Log "Remote directory listed. Files found: $($directoryInfo.Files.Count)" -Level Success
        }
        catch {
            throw "Failed to list remote directory '/ParentSquare/ExportsFromVeracross': $($_.Exception.Message)"
        }

        # --- Find Latest Files ---
        Write-Log "Searching for latest ParentSquare export files..." -Level Info

        $FilePatterns = @(
            "rosters",
            "schools",
            "sections",
            "sections",
            "staff",
            "students",
            "terms",
            "parents"
        )
        $downloadedFiles = @()

        # Loop through each pattern to find and download the latest matching ParentSquare file
        foreach ($pattern in $FilePatterns) {
            Write-Log "Searching for '$pattern' export file..." -Level Info
            # Adjust the file pattern as needed to match the actual naming convention of the ParentSquare export files
            $latestParentSquareFile = $directoryInfo.Files |
                Where-Object { $_.Name -like "*parentsquare_$pattern*.csv" } |
                Sort-Object LastWriteTime -Descending |
                Select-Object -First 1

            # Note: The pattern is included in the file name search to ensure we are looking for the correct type of ParentSquare file
            if ($latestParentSquareFile) {
                Write-Log "Found file: $($latestParentSquareFile.Name) (Last modified: $($latestParentSquareFile.LastWriteTime))" -Level Success
                $remotePath = "/ParentSquare/ExportsFromVeracross/$($latestParentSquareFile.Name)"
                $localDest = "$parentSquareBaseFolder\*"
                # The local destination is set to the folder with a wildcard to allow WinSCP to determine the filename from the remote file
                try {
                    Write-Log "Downloading ParentSquare file from $remotePath to $parentSquareBaseFolder..." -Level Info
                    $session.GetFiles($remotePath, $localDest).Check()
                    Write-Log "ParentSquare file '$($latestParentSquareFile.Name)' downloaded successfully" -Level Success
                    $downloadedFiles += "$parentSquareBaseFolder\$($latestParentSquareFile.Name)"
                }
                catch {
                    throw "Failed to download ParentSquare file '$($latestParentSquareFile.Name)': $($_.Exception.Message)"
                }
            }
            else {
                Write-Log "No matching ParentSquare files found for pattern '$pattern' in /ParentSquare/ExportsFromVeracross" -Level Warning
            }
        }

        # # --- Move Downloaded Source Files to Dated Archive Folder ---
        # if ($downloadedFiles.Count -gt 0) {
        #     Write-Log "Archiving $($downloadedFiles.Count) downloaded source file(s)..." -Level Info

        #     $archiveFolder = Join-Path "C:\_ParentSquare\Processed" (Get-Date -Format "yyyy-MM-dd")

        #     if (-not (Test-Path $archiveFolder)) {
        #         New-Item -ItemType Directory -Path $archiveFolder | Out-Null
        #         Write-Log "Created archive folder: '$archiveFolder'" -Level Info
        #     }

        #     foreach ($file in $downloadedFiles) {
        #         try {
        #             Move-Item -Path $file -Destination $archiveFolder -Force -ErrorAction Stop
        #             Write-Log "Moved source file: '$file' -> '$archiveFolder'" -Level Info
        #         } catch {
        #             Write-Log "Failed to move source file '$file': $($_.Exception.Message)" -Level Warning
        #         }
        #     }
        #     Write-Log "Archiving complete." -Level Info
        # }
        # else {
        #     Write-Log "No ParentSquare files were downloaded; skipping combine step." -Level Warning
        # }
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
