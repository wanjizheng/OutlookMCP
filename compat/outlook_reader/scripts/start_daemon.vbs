' start_daemon.vbs
' Launches outlook_daemon.ps1 in a fully hidden window. No console
' window is created and the calling process returns immediately.
' Usage (from WSL bash):
'   cscript.exe //NoLogo //B "C:\path\start_daemon.vbs" "<path-to-outlook_daemon.ps1>" [args...]
'
' WindowStyle = 0  -> SW_HIDE
' bWaitOnReturn = False -> fire and forget
'
' IMPORTANT: do NOT append shell-style `>` / `2>` redirects to `cmd`. WshShell.Run
' calls CreateProcess directly, so those tokens are passed as positional args
' to powershell.exe and the daemon silently exits with a "file not found" error.
' The daemon script already writes its own logs to %TEMP%.

Option Explicit

Dim WshShell
Dim fso
Dim scriptPath
Dim daemonScript
Dim cmd

If WScript.Arguments.Count < 1 Then
    WScript.Echo "Usage: start_daemon.vbs <outlook_daemon.ps1> [args]"
    WScript.Quit 2
End If

daemonScript = WScript.Arguments(0)

' Build the powershell command line. Pass -WindowStyle Hidden so the daemon
' never opens a console window, even when cscript is launched from a console.
cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & daemonScript & """"
Dim i
For i = 1 To WScript.Arguments.Count - 1
    cmd = cmd & " """ & WScript.Arguments(i) & """"
Next

Set WshShell = CreateObject("WScript.Shell")
' 0 = hidden window. False = do not wait.
WshShell.Run cmd, 0, False

' cscript/wscript is itself windowless; no flash here.
