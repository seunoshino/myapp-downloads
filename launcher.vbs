Set objShell = CreateObject("Wscript.Shell")
ps1 = """" & CreateObject("Scripting.FileSystemObject").GetSpecialFolder(2) & "\MyExcelTemp\script.ps1" & """"
cmd = "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File " & ps1
objShell.Run cmd, 0, False
