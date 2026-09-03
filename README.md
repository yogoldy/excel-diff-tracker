# Excel Scenario Analysis Tool

![Excel Scenario Analysis Tool icon](assets/app-icon.svg)

Excel Scenario Analysis Tool quietly watches Excel workbooks, preserves readable save-by-save evidence, and lets the user compare the current workbook with the previous save, the original baseline, or a chosen saved scan. It is a standalone Windows 11 ARM64 desktop application: no account, server, API, AI, Python, virtual environment, Excel add-in, or separate .NET installation is required.

Everything stays on your computer.

Development status: 0.2.0 is an unqualified local candidate on the `excel-scenario-analysis-tool` branch. Version 0.1.2 remains the installed everyday-use baseline. No 0.2.0 public release is authorized. See [Project status](docs/PROJECT_STATUS.md) for the current state and [Validation](docs/VALIDATION.md) for completed historical checks.

## What it tracks

- Cell values added, changed, or cleared
- Formula text added, changed, or removed
- Formula results that changed while the stored formula stayed the same
- Cell data-type changes
- Sheets added, removed, renamed, reordered, hidden, or unhidden
- Saves that contain only ignored changes, recorded as “no tracked changes”

Version 0.2 intentionally ignores styles and colors, comments, charts, names, and other workbook metadata.

## Install

No public 0.2.0 installer is currently approved. A local candidate build produces `ExcelScenarioAnalysisTool-Setup-arm64.exe`; it is self-contained and installs only for the current Windows user.

The initial unsigned release can display a Windows “unknown publisher” or SmartScreen warning. Verify the installer against `SHA256SUMS.txt` from the same release before running it.

## Use

1. Choose the folder where Markdown reports should be saved.
2. Add an `.xlsx` or `.xlsm` workbook.
3. Leave Excel Scenario Analysis Tool running in the system tray.
4. Save the workbook normally in Excel.

The first capture is silent scan 0. Each later distinct save creates an immutable chronological scan in the local history database and one Markdown report. Selecting a different comparison baseline changes only the derived view; it never rewrites the saved sequence or automatic reports. Macro-enabled workbooks are inspected as files; macros are never executed by the tracker.

New 0.2 history is stored at `%LocalAppData%\Excel Scenario Analysis Tool\history.db`. The renamed app starts fresh and offers to delete only the three old SQLite database files after explicit confirmation. Existing Markdown reports and workbooks are never touched by that transition. Stopping tracking keeps history; permanent per-workbook purge remains a separate confirmed action.

## Supported files and limits

Version 0.2 supports `.xlsx` and `.xlsm`. Encrypted workbooks, legacy `.xls`, binary `.xlsb`, corrupt packages, and unsafe workbook packages are rejected and never replace the last valid baseline.

Automatic reports include up to 5,000 detailed changes. Every change remains in SQLite, and the History page can export a full Markdown report.

## Privacy

Excel Scenario Analysis Tool has no telemetry and makes no network requests. Workbook contents, snapshots, change history, comparisons, and reports remain local. See [Privacy](docs/PRIVACY.md).

## Development

Development requires Windows 11, the .NET 10 SDK, and PowerShell. There is no Python environment.

```powershell
dotnet restore .\ExcelDiffTracker.slnx
dotnet build .\ExcelDiffTracker.slnx -c Release
dotnet test .\ExcelDiffTracker.slnx -c Release --no-build
```

To produce the self-contained ARM64 application and installer, install Inno Setup 7 and run:

```powershell
.\scripts\build-release.ps1 -Version 0.2.0
```

Architecture and test details are in [Architecture](docs/ARCHITECTURE.md) and [Testing](docs/TESTING.md). The build-once promotion process is documented in [Releasing](docs/RELEASING.md).

## License and trademarks

Excel Scenario Analysis Tool is available under the [MIT License](LICENSE). Required acknowledgments are in [Third-party notices](THIRD-PARTY-NOTICES.md).

Microsoft and Excel are trademarks of the Microsoft group of companies. Excel Scenario Analysis Tool is an independent project and is not affiliated with, endorsed by, or sponsored by Microsoft.
