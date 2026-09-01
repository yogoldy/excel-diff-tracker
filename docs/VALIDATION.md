# Validation record

Validation was performed on 2026-08-31 and finalized on 2026-09-01 in a Windows 11 ARM64 Parallels VM using .NET SDK 10.0.400 and runtime 10.0.11. The release candidate is a self-contained ARM64 application.

## Initial 0.1.0 checks

- Release build: succeeded with zero warnings and zero errors.
- Automated suite: 19 of 19 tests passed.
- Real Excel save: a silent baseline followed by a value and formula edit produced one version, two cell changes, no processing errors, and a Markdown report.
- Macro safety: the MIT-licensed CLISC `Docs/To_Strict.xlsm` fixture contained `xl/vbaProject.bin` and ActiveX parts. Extraction found three sheets and 19 populated cells, did not execute VBA, and left the source SHA-256 unchanged.
- Encrypted input: a workbook encrypted by real Excel was rejected as corrupt or encrypted and did not advance a baseline.
- Large workbook: a generated workbook containing 500,000 populated cells was extracted successfully in approximately 3.7 seconds. Peak process memory during this stress check was approximately 735 MiB; extraction runs off the WPF UI thread.
- Tracking resilience: automated coverage passed for rapid saves, restart reconciliation, atomic file replacement, multiple independently tracked workbooks, locked-file retry, corrupt input, hash deduplication, and capture/purge coordination.
- Markdown safety: untrusted workbook text was prevented from creating active Markdown links or raw HTML in reports.
- Installer: silent per-user installation, Start-menu integration, startup entry, process launch, silent uninstall, and installed-file cleanup passed with the development .NET directory temporarily unavailable. The original process-liveness assertion did not inspect the first-run window and therefore missed the onboarding binding failure later reported by a user.
- Release contents: the installer bundles the applicable .NET, WPF, Windows SDK for .NET, Open XML SDK, SQLitePCLRaw, and SQLite notices and license texts.
- WinGet 1.29.290: the three-file manifest validated without warnings against the public GitHub release. WinGet downloaded the installer and independently verified its exact SHA-256.

## 0.1.1 startup-fix validation

Validation was repeated on 2026-09-01 in the same Windows 11 ARM64 VM.

- Root cause: `ProgressBar.Value` binds two-way by default, while the onboarding `StepNumber` source is intentionally read-only. The binding is now explicitly one-way.
- Independent review found no other read-only source bound to a two-way-default target in the current WPF views.
- Release build: succeeded with zero warnings and zero errors.
- Automated suite: 20 of 20 tests passed, including a Windows STA regression test that constructs the onboarding window and asserts the one-way progress binding.
- Installer: the exact self-contained 0.1.1 candidate passed silent per-user installation, Start-menu and startup-entry checks, opened `Welcome to Excel Diff Tracker` against isolated first-run data, and uninstalled cleanly with the development .NET installation unavailable.
- The installer smoke test now requires the expected first-run window title, so a modal `Startup problem` dialog can no longer produce a false pass merely by keeping the process alive.
- Real Excel save: repeated against the 0.1.1 source and passed with one captured version, two expected cell changes, no processing errors, and a Markdown report containing the value and formula changes.
- Core workbook extraction, diffing, persistence, tracking, and reporting code did not change in 0.1.1; the macro, encrypted-file, and 500,000-cell checks were not repeated for this UI-only patch.

## Deliberately outstanding

- Manual visual and accessibility acceptance of onboarding, tray states, themes, keyboard navigation, high contrast, text truncation, and 200% scaling. Visual VM control was not used.
- End-to-end WinGet local-manifest installation still needs manual SmartScreen approval because version 0.1.1 is unsigned. The automated CLI test will not bypass Windows security. Submission to `microsoft/winget-pkgs` should wait until this and the manual visual acceptance are complete.
- Code signing. Version 0.1.1 is unsigned and may trigger an unknown-publisher or SmartScreen warning.
