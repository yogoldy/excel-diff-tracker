@echo off
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\Dev\excel-diff-tracker\scripts\create-encrypted-excel-fixture.ps1 -OutputPath C:\Dev\excel-diff-tracker\artifacts\encrypted-acceptance.xlsx > C:\Dev\excel-diff-tracker\artifacts\encrypted-fixture.log 2>&1
exit /b %errorlevel%
