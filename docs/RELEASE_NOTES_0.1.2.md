# Excel Diff Tracker 0.1.2

Open-workbook reliability and visual-accessibility update for Windows 11 ARM64.

## Fixed

- Allows safe read-only capture while Excel retains a compatible write handle on an open `.xlsx` or `.xlsm` workbook.
- Adds byte-for-byte source verification across a second quiet interval before a temporary copy is accepted.
- Recovers automatically after a locked-file timeout without requiring another save or advancing the last valid baseline.
- Corrects light/dark text inheritance and themed ComboBox rendering that could produce light-on-light or dark-on-dark content.
- Adds responsive scrolling, wrapping, and card layouts for onboarding and main pages at constrained window sizes and higher display scaling.
- Adds stable Windows UI Automation identifiers throughout onboarding and the primary application views.
- Gives repeated workbook/history actions distinct accessible identifiers, exposes complete truncated paths, and improves onboarding warning contrast.

## Release qualification

- The source includes a fail-closed black-box acceptance contract and external UI Automation/SQLite evidence tools.
- Application and acceptance-probe code are bundled into their hashed executables so loose runtime DLLs cannot escape candidate identity checks.
- Publishing requires two clean installed-product runs, the required visual/scaling matrix, hashed evidence, and independent functional, visual, and release attestations for the exact installer SHA-256.
- Final validation measurements and installer hash will be recorded after the frozen candidate completes that gate.

## Known limitations

- Styles, colors, comments, charts, names, and other workbook metadata are intentionally ignored.
- `.xls`, `.xlsb`, encrypted, corrupt, and unsafe workbooks are rejected.
- The installer is unsigned and Windows may show an unknown-publisher or SmartScreen warning.
- Windows x64 and automatic updates are not included in this release.
