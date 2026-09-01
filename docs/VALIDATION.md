# Validation record

Validation was performed on 2026-08-31 and finalized on 2026-09-01 in a Windows 11 ARM64 Parallels VM using .NET SDK 10.0.400 and runtime 10.0.11. The release candidate is a self-contained ARM64 application.

## Completed checks

- Release build: succeeded with zero warnings and zero errors.
- Automated suite: 19 of 19 tests passed.
- Real Excel save: a silent baseline followed by a value and formula edit produced one version, two cell changes, no processing errors, and a Markdown report.
- Macro safety: the MIT-licensed CLISC `Docs/To_Strict.xlsm` fixture contained `xl/vbaProject.bin` and ActiveX parts. Extraction found three sheets and 19 populated cells, did not execute VBA, and left the source SHA-256 unchanged.
- Encrypted input: a workbook encrypted by real Excel was rejected as corrupt or encrypted and did not advance a baseline.
- Large workbook: a generated workbook containing 500,000 populated cells was extracted successfully in approximately 3.7 seconds. Peak process memory during this stress check was approximately 735 MiB; extraction runs off the WPF UI thread.
- Tracking resilience: automated coverage passed for rapid saves, restart reconciliation, atomic file replacement, multiple independently tracked workbooks, locked-file retry, corrupt input, hash deduplication, and capture/purge coordination.
- Markdown safety: untrusted workbook text was prevented from creating active Markdown links or raw HTML in reports.
- Installer: silent per-user installation, Start-menu integration, startup entry, launch, silent uninstall, and installed-file cleanup passed with the development .NET directory temporarily unavailable. User history is retained by design.
- Release contents: the installer bundles the applicable .NET, WPF, Windows SDK for .NET, Open XML SDK, SQLitePCLRaw, and SQLite notices and license texts.
- WinGet 1.29.290: the three-file manifest validated without warnings against the public GitHub release. WinGet downloaded the installer and independently verified its exact SHA-256.

## Deliberately outstanding

- Manual visual and accessibility acceptance of onboarding, tray states, themes, keyboard navigation, high contrast, text truncation, and 200% scaling. Visual VM control was not used.
- End-to-end WinGet local-manifest installation still needs manual SmartScreen approval because version 0.1.0 is unsigned. The automated CLI test was stopped at that prompt rather than bypassing Windows security. Submission to `microsoft/winget-pkgs` should wait until this and the manual visual acceptance are complete.
- Code signing. Version 0.1.0 is unsigned and may trigger an unknown-publisher or SmartScreen warning.
