Set WshShell = CreateObject("WScript.Shell")
' The 0 at the end is the magic command that tells Windows to hide the window completely
WshShell.Run chr(34) & "C:\Users\Kasa\Documents\DormsMaster\ProgramStarter\start.bat" & Chr(34), 0
Set WshShell = Nothing