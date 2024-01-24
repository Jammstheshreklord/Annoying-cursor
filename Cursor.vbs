Set objShell = CreateObject("WScript.Shell")

Do
    objShell.Run "rundll32.exe user32.dll,SetCursorPos 0,0", 0, True
Loop