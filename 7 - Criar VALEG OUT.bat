@echo off
if "%~1"=="" (
  powershell.exe -Sta -NoProfile -ExecutionPolicy Bypass -File "%~dp0CriarValegOutGerado.ps1"
) else (
  powershell.exe -Sta -NoProfile -ExecutionPolicy Bypass -File "%~dp0CriarValegOutGerado.ps1" -WorkbookPath "%~1"
)
if errorlevel 1 pause
