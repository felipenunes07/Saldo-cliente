@echo off
if "%~1"=="" (
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0ColarResumoGrupoWeChat.ps1" -AutoStart -NoFinalMessage
) else (
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0ColarResumoGrupoWeChat.ps1" -QueuePath "%~1" -AutoStart -NoFinalMessage
)
if errorlevel 1 pause
