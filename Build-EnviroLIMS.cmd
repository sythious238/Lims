@echo off
rem Double-click to build EnviroLIMS.accdb (requires Microsoft Access).
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Build-EnviroLIMS.ps1" %*
pause
