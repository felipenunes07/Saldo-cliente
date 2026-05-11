@echo off
if "%~1"=="" (
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0PrepararEnvioWeChat.ps1"
) else (
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0PrepararEnvioWeChat.ps1" -WorkbookPath "%~1"
)
if errorlevel 1 pause
