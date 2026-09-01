# Excel Diff Tracker

![Excel Diff Tracker icon](assets/app-icon.svg)

Excel Diff Tracker quietly watches your Excel workbooks and writes a readable Markdown change report after every distinct save. It is a standalone Windows 11 ARM64 desktop application: no account, server, API, AI, Python, virtual environment, Excel add-in, or separate .NET installation is required.

Everything stays on your computer.

## What it tracks

- Cell values added, changed, or cleared
- Formula text added, changed, or removed
- Formula results that changed while the stored formula stayed the same
- Cell data-type changes
- Sheets added, removed, renamed, reordered, hidden, or unhidden
- Saves that contain only ignored changes, recorded as “no tracked changes”

Version 0.1 intentionally ignores styles and colors, comments, charts, names, and other workbook metadata.

## Install

Download `ExcelDiffTracker-Setup-arm64.exe` from the project’s [GitHub Releases](https://github.com/yogoldy/excel-diff-tracker/releases), run it, and follow the short first-run setup. The release is self-contained and installs only for your Windows user; administrator access is not required.

The initial unsigned release can display a Windows “unknown publisher” or SmartScreen warning. Verify the installer against `SHA256SUMS.txt` from the same release before running it.

## Use

1. Choose the folder where Markdown reports should be saved.
2. Add an `.xlsx` or `.xlsm` workbook.
3. Leave Excel Diff Tracker running in the system tray.
4. Save the workbook normally in Excel.

The first capture is a silent baseline. Each later distinct save creates a version in the local history database and one Markdown report. Macro-enabled workbooks are inspected as files; macros are never executed by the tracker.

History is stored at `%LocalAppData%\Excel Diff Tracker\history.db`. Stopping tracking keeps that history. Permanent deletion is a separate confirmed action in the app. Uninstalling removes the application, shortcuts, and startup entry but deliberately preserves user history and reports; purge them in the app first if you want those records removed too.

## Supported files and limits

Version 0.1 supports `.xlsx` and `.xlsm`. Encrypted workbooks, legacy `.xls`, binary `.xlsb`, corrupt packages, and unsafe workbook packages are rejected and never replace the last valid baseline.

Automatic reports include up to 5,000 detailed changes. Every change remains in SQLite, and the History page can export a full Markdown report.

## Privacy

Excel Diff Tracker has no telemetry and makes no network requests. Workbook contents, snapshots, change history, and reports remain local. See [Privacy](docs/PRIVACY.md).

## Development

Development requires Windows 11, the .NET 10 SDK, and PowerShell. There is no Python environment.

```powershell
dotnet restore .\ExcelDiffTracker.slnx
dotnet build .\ExcelDiffTracker.slnx -c Release
dotnet test .\ExcelDiffTracker.slnx -c Release --no-build
```

To produce the self-contained ARM64 application and installer, install Inno Setup 7 and run:

```powershell
.\scripts\build-release.ps1 -Version 0.1.1
```

Architecture and test details are in [Architecture](docs/ARCHITECTURE.md) and [Testing](docs/TESTING.md). The build-once promotion process is documented in [Releasing](docs/RELEASING.md).

## License and trademarks

Excel Diff Tracker is available under the [MIT License](LICENSE). Required acknowledgments are in [Third-party notices](THIRD-PARTY-NOTICES.md).

Microsoft and Excel are trademarks of the Microsoft group of companies. Excel Diff Tracker is an independent project and is not affiliated with, endorsed by, or sponsored by Microsoft.
