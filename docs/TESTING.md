# Testing

## Automated suite

Run from a Windows checkout with the .NET 10 SDK:

```powershell
dotnet build .\ExcelDiffTracker.slnx -c Release
dotnet test .\ExcelDiffTracker.slnx -c Release --no-build
```

The suite covers value/formula/type transitions, formula-result-only changes, stable sheet identity operations, `.xlsx`, `.xlsm`, rejected legacy formats, silent baseline behavior, hash deduplication, corrupt-package baseline preservation, zero-delta saves, transactional history, future-schema rejection, report truncation and full export, Markdown injection resistance, pending-report recovery, rapid sequential saves, restart reconciliation, independent multi-workbook capture during locked-file retries, atomic replacements, and purge/capture coordination.

## Live Excel smoke test

The optional smoke test creates and saves a real workbook through Excel, then checks that the tracker records one version with two expected cell changes and writes Markdown. Excel must be installed and the test must run in an interactive signed-in Windows account.

```powershell
.\scripts\live-excel-smoke.ps1 -RepositoryRoot (Get-Location).Path
```

## Installer acceptance

Build the release, then run:

```powershell
.\scripts\test-installer.ps1 -InstallerPath .\artifacts\release\ExcelDiffTracker-Setup-arm64.exe
```

This performs a silent per-user install, verifies the self-contained payload and startup entry, launches the app, and silently uninstalls it. User data is deliberately retained.

The exact release candidate has also been exercised with real Excel, a real macro-bearing workbook, a real password-encrypted workbook, and a generated 500,000-cell workbook on Windows 11 ARM64. See [VALIDATION.md](VALIDATION.md) for the recorded results.

Manual visual acceptance remains necessary for onboarding, tray behavior, light/dark/system and high-contrast themes, keyboard navigation, text truncation, and 200% scaling. It was deliberately not automated through visual VM control.
