@echo off
set "BASE=%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%BASE%RodarResumoGrupoWeChat.ps1" %*
if errorlevel 1 pause
