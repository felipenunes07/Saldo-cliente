@echo off
if "%~1"=="" (
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0ColarWeChatRascunhos.ps1"
) else (
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0ColarWeChatRascunhos.ps1" -QueuePath "%~1"
)
if errorlevel 1 pause
