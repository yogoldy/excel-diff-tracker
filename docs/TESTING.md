# Testing

## Automated suite

Run from a Windows checkout with the .NET 10 SDK:

```powershell
dotnet build .\ExcelDiffTracker.slnx -c Release
dotnet test .\ExcelDiffTracker.slnx -c Release --no-build
```

The suite covers value/formula/type transitions, formula-result-only changes, stable sheet identity operations, `.xlsx`, `.xlsm`, rejected legacy formats, silent baseline behavior, hash deduplication, corrupt-package baseline preservation, zero-delta saves, transactional history, future-schema rejection, report truncation and full export, Markdown injection resistance, pending-report recovery, rapid sequential saves, restart reconciliation, independent multi-workbook capture during locked-file retries, atomic replacements, and purge/capture coordination.

For supporting tests on a development PC, set `EXCEL_DIFF_TRACKER_TEST_DATA_DIRECTORY` to a fresh directory under `TestResults` in the test process. Do not launch the installed app or execute installed acceptance there. The WPF regression constructs unshown windows and template peers; it checks title lookup, distinct repeated action IDs, focusability, and full-path accessibility.

The isolated PowerShell regressions run under Windows PowerShell 5.1 and do not install, launch, or uninstall the product, drive Excel, or write to the user's registry/history:

```powershell
powershell.exe -NoProfile -ExecutionPolicy RemoteSigned -File .\tests\acceptance\Test-LifecycleGateRegressions.ps1
powershell.exe -NoProfile -ExecutionPolicy RemoteSigned -File .\tests\Test-AcceptanceSemanticRecoveryRegressions.ps1
powershell.exe -NoProfile -ExecutionPolicy RemoteSigned -File .\scripts\acceptance\tests\Test-BenchmarkAndSoakValidators.ps1
powershell.exe -NoProfile -ExecutionPolicy RemoteSigned -File .\tests\acceptance\Test-SingleFilePayloadRegressions.ps1
```

The execution policy above applies only to each child process. These tests inspect selected script functions with synthetic inputs; passing them does not satisfy the installed release contract.

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

Run this only in a clean test VM or Windows account with no existing Excel Diff Tracker installation. The script refuses to continue if the app is installed or running because it finishes by uninstalling the test copy. It performs a silent per-user install, verifies the self-contained payload and startup entry, launches the app with isolated first-run data, requires the onboarding window to open without a startup error, and silently uninstalls it. Existing history and reports are never read or changed.

The exact release candidate has also been exercised with real Excel, a real macro-bearing workbook, a real password-encrypted workbook, and a generated 500,000-cell workbook on Windows 11 ARM64. See [VALIDATION.md](VALIDATION.md) for the recorded results.

Manual visual acceptance remains necessary for onboarding, tray behavior, light/dark/system and high-contrast themes, keyboard navigation, text truncation, and 200% scaling. It was deliberately not automated through visual VM control.
