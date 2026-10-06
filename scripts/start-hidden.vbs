' 由安装程序生成：隐藏窗口启动托盘服务（供登录自启与快捷方式调用）
Option Explicit
Dim shell, fso, base, script, cmd, logFile, f
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
base = fso.GetParentFolderName(WScript.ScriptFullName)
script = base & "\tools\tray.ps1"
If Not fso.FileExists(script) Then script = base & "\tray.ps1"
logFile = base & "\logs\launcher.log"
If Not fso.FileExists(script) Then
  On Error Resume Next
  If Not fso.FolderExists(base & "\logs") Then fso.CreateFolder(base & "\logs")
  Set f = fso.OpenTextFile(logFile, 8, True)
  f.WriteLine Now & " 找不到托盘脚本: " & script
  f.Close
  On Error Goto 0
  WScript.Quit 1
End If
cmd = "powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & script & """"
shell.Run cmd, 0, False
