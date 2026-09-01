# Excel Diff Tracker 0.1.0

Initial Windows 11 ARM64 release.

## Included

- Quiet multi-workbook save tracking from the system tray
- Values, formulas, cached results, data types, and sheet-structure changes
- Silent first baseline and SHA-256 duplicate-save suppression
- Local SQLite history with one Markdown report per accepted save
- Full Markdown export for reports truncated above 5,000 detailed changes
- `.xlsx` and `.xlsm` inspection without macro execution
- Light, dark, system, and high-contrast-aware themes
- Self-contained per-user ARM64 installer with no .NET prerequisite

## Validation

- 19 automated tests passed on Windows 11 ARM64.
- Real Excel save tracking passed with the expected value and formula changes.
- A real VBA-bearing `.xlsm`, a real encrypted workbook, and a generated 500,000-cell workbook were exercised successfully; source workbooks remained unchanged, and extraction runs off the UI thread.
- The self-contained installer passed silent install, startup, launch, and uninstall checks with the development .NET installation unavailable.

## Known limitations

- Styles, colors, comments, charts, names, and other workbook metadata are intentionally ignored.
- `.xls`, `.xlsb`, encrypted, corrupt, and unsafe workbooks are rejected.
- The installer is unsigned and Windows may show an unknown-publisher or SmartScreen warning.
- Windows x64 and automatic updates are not included in this release.
