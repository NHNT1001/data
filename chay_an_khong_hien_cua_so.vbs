Set WshShell = CreateObject("WScript.Shell")
WshShell.CurrentDirectory = "D:\Antigravity Tai\Github\data"
WshShell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File ""D:\Antigravity Tai\Github\data\auto_push.ps1""", 0, False
