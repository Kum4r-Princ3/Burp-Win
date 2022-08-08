Set WshShell = CreateObject("WScript.Shell")
		WshShell.Run chr(34) & "D:\BURPSUITE_PRO\burp.bat" & Chr(34), 0
		Set WshShell = Nothing