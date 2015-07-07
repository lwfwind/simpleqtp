Dim WshShell
Set WshShell = CreateObject("WScript.Shell")
WshShell.run("cscript " & Chr(34) & WshShell.CurrentDirectory & Chr(34) & "\ScheduleAOM.vbs")
Set WshShell = nothing
