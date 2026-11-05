@echo off
chcp 65001 >nul
if "%~1"=="" (
  powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File "%~dp0Word_ReadOnly_Open_V1_1.ps1"
) else (
  powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File "%~dp0Word_ReadOnly_Open_V1_1.ps1" -FilePath "%~1"
  if errorlevel 1 pause
)
