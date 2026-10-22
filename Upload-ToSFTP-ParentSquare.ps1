# Parentsquare - Upload Files to Parentsquare SFTP
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
    try {
        $logDir = "C:\_Parentsquare\_logs"

        if (-not (Test-Path -Path $logDir)) {
            New-Item -ItemType Directory -Path $logDir -Force -ErrorAction Stop | Out-Null
        }

        $logFileName = "uploadParentSquareFiles_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
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
    Write-Log "=== Parentsquare File Upload Started ===" -Level Info

    # --- Configuration ---
    $winscpDllPath        = "C:\Program Files (x86)\WinSCP\netstandard2.0\WinSCPnet.dll"
    $winscpExePath        = "C:\Program Files (x86)\WinSCP\WinSCP.exe"
    $hostName = "sftp3.parentsquare.com"
    $userName = "shadyc401c4a1"
    $password = "YBAcJRUy2NgpyCkKp3wP"
    $remotePath = "/sftp-inbox-production-parentsquare/shadyc401c4a1/"
    $sshHostKeyFingerprint = "ssh-rsa 2048 5FRJ9tPAZ+NlprGaMZ5c4rZ1mT0BiDw+fkdXGmrVyRc"
    $uploadSourceFolder  = "C:\_ParentSquare\Processed"

    # --- Get files to upload ---
    Write-Log "Scanning for files to upload in: $uploadSourceFolder" -Level Info
    #get all files in the upload source folder
    $files = Get-ChildItem -Path $uploadSourceFolder -File -ErrorAction Stop
    
    if ($files.Count -eq 0) {
        Write-Log "No files found to upload in: $uploadSourceFolder" -Level Warning
        Save-LogFile | Out-Null
        Read-Host "Press Enter to exit"
        exit 0
    }

    Write-Log "$($files.Count) file(s) queued for upload" -Level Info

    # --- Load WinSCP Assembly ---
    Write-Log "Loading WinSCP .NET assembly..." -Level Info

    if (-not (Test-Path -Path $winscpDllPath)) {
        throw "WinSCP DLL not found at: $winscpDllPath`nPlease verify WinSCP is installed."
    }

    try {
        Add-Type -Path $winscpDllPath -ErrorAction Stop
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
            HostName              = $hostName
            UserName              = $userName
            Password              = $password
            SshHostKeyFingerprint = $sshHostKeyFingerprint
        }
        Write-Log "Session options configured for host: $hostName" -Level Success
    }
    catch {
        throw "Failed to configure session options: $($_.Exception.Message)"
    }

    # Create a folder in the _ParentSquare/Processed directory for today's date for the processed files if it doesn't exist
    $todayDate = Get-Date -Format "yyyy-MM-dd"
    $todayProcessedFolder = Join-Path -Path $uploadSourceFolder -ChildPath $todayDate
    if (-not (Test-Path -Path $todayProcessedFolder)) {
        New-Item -ItemType Directory -Path $todayProcessedFolder -Force -ErrorAction Stop | Out-Null
        Write-Log "Created folder for today's processed files: $todayProcessedFolder" -Level Info
    }


    # --- Connect and Upload ---
    Write-Log "Connecting to SFTP server: $hostName..." -Level Info
    $session = New-Object WinSCP.Session
    $session.ExecutablePath = $winscpExePath

    try {
        try {
            $session.Open($sessionOptions)
            Write-Log "Connected to $hostName successfully" -Level Success
        }
        catch {
            throw "Failed to connect to SFTP server '$hostName': $($_.Exception.Message)"
        }

        $filesCount    = $files.Count
        $uploadedCount = 0
        $failedFiles   = @()

        foreach ($file in $files) {
            $localFile = $file.FullName
            Write-Log "Uploading: $localFile" -Level Info

            try {
                $transferOptions = New-Object WinSCP.TransferOptions
                $transferOptions.TransferMode = [WinSCP.TransferMode]::Binary
                $transferOptions.ResumeSupport.State = [WinSCP.TransferResumeSupportState]::Off
                $transferOptions.PreserveTimestamp = $False
                $transferOptions.FilePermissions = $null

                $transferResult = $session.PutFiles($localFile, $remotePath, $False, $transferOptions)
                $transferResult.Check()

                foreach ($transfer in $transferResult.Transfers) {
                    Write-Log "Uploaded successfully: $($transfer.FileName)" -Level Success
                    
                    #move the file to the today's processed folder
                    $destinationPath = Join-Path -Path $todayProcessedFolder -ChildPath $file.Name
                    Move-Item -Path $localFile -Destination $destinationPath -Force -ErrorAction Stop
                    Write-Log "Moved file to processed folder: $destinationPath" -Level Success
                    
                    $uploadedCount++
                }
            }
            catch {
                Write-Log "Failed to upload file '$($file.Name)': $($_.Exception.Message)" -Level Error
                $failedFiles += $file.Name
            }
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

    if ($failedFiles.Count -gt 0) {
        Write-Log "Upload completed with errors. Uploaded $uploadedCount of $filesCount file(s)." -Level Warning
        Write-Log "Failed files: $($failedFiles -join ', ')" -Level Warning
        Save-LogFile | Out-Null
    }
    else {
        Write-Log "All $uploadedCount of $filesCount file(s) uploaded successfully!" -Level Success
        Write-Log "=== UPLOAD COMPLETED SUCCESSFULLY ===" -Level Success
        Save-LogFile | Out-Null
    }

    #Read-Host "Press Enter to exit"
    exit 0
}
catch {
    Write-Log "=== CRITICAL ERROR ===" -Level Error
    Write-Log "Error Message: $($_.Exception.Message)" -Level Error
    Write-Log "Error Location: $($_.InvocationInfo.ScriptName):$($_.InvocationInfo.ScriptLineNumber)" -Level Error
    Write-Log "Stack Trace: $($_.ScriptStackTrace)" -Level Error
    Write-Log "Script failed. Please review the error details above." -Level Error

    Save-AndOpenLogOnError

    #Read-Host "Press Enter to exit"
    exit 1
}