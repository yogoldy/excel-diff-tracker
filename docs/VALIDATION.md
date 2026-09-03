# Validation record

## 0.2.0 local candidate and maintainer trial (2026-09-03 UTC)

The frozen 0.2.0 build source is `845f20283182c5b529c74370918cef3419ca7631`. The candidate installer SHA-256 is `3562301F1EA6375A1726981A263E443BB794627755F1746304396D33167A2144`, and the application SHA-256 is `9244736A10ED36F38F4088A5DF6A6369DCC664417AE8D4FF62F3CD04C3A182CE`.

- The source suite passed 30 tests.
- Targeted minimum-size visual review passed.
- Two independent Sol-medium reviews accepted the local candidate within their stated scope.
- The maintainer installed and launched the candidate and positively assessed the redesigned visual hierarchy, workbook action menus, and compact/expanded views.
- The prior and new processes were simultaneously active during the first launch. Optional legacy cleanup could not remove a locked SQLite write-ahead-log file and reported the failure. It did not touch a workbook or Markdown report.

This establishes source-level checks, artifact identity, successful launch, selected visual acceptance, and safe cleanup failure. It does not establish clean replacement, reliable process handoff, complete legacy cleanup, sustained everyday use, or public-release qualification. The structured maintainer intake and a later stabilization pass remain required. Private installed-state evidence remains local and untracked.

## 0.1.2 existing-profile upgrade and sign-in (2026-09-03 UTC)

Current phase: installed candidate entering maintainer-led everyday use. See [Project status](PROJECT_STATUS.md) for working preferences and next steps. Source through `3fb6304e6d95f0b1b83180d065f3e8be8ce5cfc7` is on `main`; this is not a published or fully qualified 0.1.2 release.

- Frozen build: zero warnings/errors, 26/26 automated C# tests passed, 33 PowerShell files parsed, 112 isolated PowerShell regression checks passed, and 48 numeric palette contrast checks passed. Numeric contrast checks do not establish visual acceptance.
- Candidate installer SHA-256: `9536373FCC47119D2C10A75C8E6D8F1BFC2BF38FDF08639DD27790D95422133A`.
- Installed application SHA-256: `0518A0F02696D698F36F1118960557C805369B815A0941F0A5B75D8304A5D6D8`.
- With the maintainer's authorization, the prior app exited through its tray menu and 499 installation/data files were backed up and hash-verified before upgrading the everyday profile.
- Candidate installation exited successfully, preserved the database byte-for-byte, passed the single-file payload check, and corrected an existing startup entry that pointed to the wrong system-profile installation directory.
- After the maintainer performed a real Windows sign-out/sign-in, a new candidate process was observed in the new session with `--background`, matching the frozen executable hash. No visible app window appeared before interaction. The tray reopened the app and History showed both existing versions.
- Database integrity passed before and after the upgrade/sign-in. All nine tables retained identical full-row fingerprints, including the existing workbook, two versions, and one preexisting capture-error record. No new capture-error record was added during these checks.
- Local evidence and backups are retained under `TestResults/release-preflight/`; personal data is not committed. No uninstall, clean-profile acceptance, or full installed semantic/recovery/performance matrix was performed in this session.

These checks establish upgrade/history preservation and quiet startup on this existing profile. They do not establish sustained real-workflow capture reliability or satisfy the complete public-release gate. No release tag or public asset was created.

## 0.1.2 release preflight (2026-09-03 UTC)

The continued audit found defects in prior-release UI selectors, PowerShell nullable/missing-property handling, populated-page accessibility, warning contrast, recovery semantic checks, benchmark heartbeat coverage, soak completion evidence, and executable identity coverage. The corresponding source/harness corrections have isolated regression coverage. The candidate application and AcceptanceProbe now publish as single-file executables including native dependencies so their hashes cover the executable payload.

During the initial source preflight, the user's installed application and profile were not exercised. The subsequent authorized existing-profile checks are recorded above. Two clean critical-path runs, the full disposable lifecycle/upgrade/uninstall gate, real Excel semantic/recovery/benchmark/soak matrices, the complete visual matrix, explicit human visual approval, aggregate evidence validation, and final independent attestations remain outstanding. Source tests and a compiled installer do not approve release 0.1.2.

The earlier 0.1.0/0.1.1 validation below was performed on 2026-08-31 and finalized on 2026-09-01 in a Windows 11 ARM64 Parallels VM using .NET SDK 10.0.400 and runtime 10.0.11.

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
