# Excel Diff Tracker 0.1.1

First-run startup fix for Windows 11 ARM64.

## Fixed

- Corrected the onboarding progress indicator binding that caused version 0.1.0 to show a `Startup problem` dialog on a new installation.
- Added a Windows WPF regression test that constructs the first-run window and verifies its progress binding is one-way.
- Strengthened installer acceptance so it requires the real first-run welcome window, rather than only checking that the process remains running.

## Validation

- Release build completed on Windows 11 ARM64 with zero warnings and zero errors.
- 20 automated tests passed, including the new first-run WPF regression test.
- The self-contained installer opened first-run onboarding successfully using isolated test data and passed silent per-user installation and uninstall checks.
- A real Excel save produced one version, the two expected cell changes, no processing errors, and the expected Markdown report.

## Known limitations

- Styles, colors, comments, charts, names, and other workbook metadata are intentionally ignored.
- `.xls`, `.xlsb`, encrypted, corrupt, and unsafe workbooks are rejected.
- The installer is unsigned and Windows may show an unknown-publisher or SmartScreen warning.
- Windows x64 and automatic updates are not included in this release.
