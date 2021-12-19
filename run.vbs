Dim WinScriptHost
Set WinScriptHost = CreateObject("WScript.Shell")
WinScriptHost.Run Chr(34) & "Enable Smart Screen.bat" & Chr(34), 0
Set WinScriptHost = Nothing