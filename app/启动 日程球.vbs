' 日程球 launcher - starts the WPF orb without a console window
Set sh = CreateObject("WScript.Shell")
sh.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File ""D:\DeepSeek\workspace\schedule-orb\orb.ps1""", 0, False
