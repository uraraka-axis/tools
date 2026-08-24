' morning_startup_hidden.vbs - launch morning_startup.ps1 fully hidden (no console flash).
' Called by the scheduled task "MorningTools" (At log on + Daily). No admin needed.
Set sh = CreateObject("WScript.Shell")
sh.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File ""C:\Users\ssasa\tools\morning-startup\morning_startup.ps1""", 0, False
