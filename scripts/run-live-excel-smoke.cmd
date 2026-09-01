@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
set "LOG_PATH=%SCRIPT_DIR%..\artifacts\live-excel-task.log"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%live-excel-smoke.ps1" > "%LOG_PATH%" 2>&1
exit /b %ERRORLEVEL%
