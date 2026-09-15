Set WshShell = CreateObject("WScript.Shell")
WshShell.CurrentDirectory = "D:\data"
WshShell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File ""D:\data\auto_push.ps1""", 0, False
