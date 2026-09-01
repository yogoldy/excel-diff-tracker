@echo off
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\Dev\excel-diff-tracker\scripts\create-large-excel-fixture.ps1 -OutputPath C:\Dev\excel-diff-tracker\artifacts\large-500k.xlsx > C:\Dev\excel-diff-tracker\artifacts\large-fixture.log 2>&1
exit /b %errorlevel%
