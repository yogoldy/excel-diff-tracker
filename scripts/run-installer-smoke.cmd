@echo off
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\Dev\excel-diff-tracker\scripts\test-installer.ps1 -InstallerPath C:\Dev\excel-diff-tracker\artifacts\release\ExcelDiffTracker-Setup-arm64.exe -RequireNoDotnet > C:\Dev\excel-diff-tracker\artifacts\installer-smoke.log 2>&1
exit /b %errorlevel%
