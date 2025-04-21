<#
.SYNOPSIS
    Removes the "Mark of the Web" (Zone.Identifier) from Excel (.xlsm) files in %TEMP%\MyExcelTemp.
.DESCRIPTION
    This script checks for and removes the MOTW (Zone.Identifier alternate data stream) from all .xlsm files in the temp project folder, then opens imarc.xlsm and exits silently.
.NOTES
    Requires PowerShell 3.0 or later.
#>

# Use the temp folder path
$folderPath = "$env:TEMP\MyExcelTemp"

# Get all .xlsm files in the folder
$excelFiles = Get-ChildItem -Path $folderPath -Filter "*.xlsm" -File

# Loop through each file and remove the Zone.Identifier stream
foreach ($file in $excelFiles) {
    $zoneIdStream = "$($file.FullName):Zone.Identifier"
    
    if (Get-Item -Path $zoneIdStream -ErrorAction SilentlyContinue) {
        try {
            Remove-Item -Path $zoneIdStream -Force
        }
        catch {
            # Silently fail if unable to remove
        }
    }
}

# ---- OPEN THE XLSM FILE ----
Start-Sleep -Seconds 1
$xlsmPath = "$folderPath\imarc.xlsm"
Start-Process "excel.exe" -ArgumentList "`"$xlsmPath`""

# Exit silently
exit