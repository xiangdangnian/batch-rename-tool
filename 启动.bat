@echo off
rem ??????? - ????
start "" powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "%~dp0BatchRename.ps1"
exit /b
