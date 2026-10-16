' ============================================================
'  BatchRename launcher  (no console window)
'  Double-click this file to start the tool.
' ============================================================
Option Explicit
Dim fso, sh, base, cmd
Set fso = CreateObject("Scripting.FileSystemObject")
Set sh  = CreateObject("WScript.Shell")
base = fso.GetParentFolderName(WScript.ScriptFullName)
sh.CurrentDirectory = base
cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File """ & base & "\BatchRename.ps1"""
sh.Run cmd, 0, False