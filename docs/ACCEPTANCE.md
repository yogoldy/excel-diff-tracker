# Release acceptance contract

Excel Diff Tracker is approved for release only when the exact installer candidate passes this contract. Source-level tests, process liveness, and direct extractor benchmarks are supporting checks; none substitutes for the installed-product gates below.

## Fail-closed rules

- Freeze the source commit and installer SHA-256 before acceptance. Any code, packaging, or installer change invalidates every earlier run.
- The installer must contain `BUILD-IDENTITY.json` and `BUILD-SOURCE-SHA256SUMS.txt`. The latter must exactly match every tracked byte in the clean release checkout except the one generated WinGet installer-hash manifest. The embedded build-source commit must be an ancestor of the release commit, and that manifest must be their only changed path.
- Run the critical path twice from named, clean Windows 11 ARM64 VM snapshots using a standard non-administrator account, real desktop Microsoft Excel (recording whether its executable is ARM64 or x64-under-emulation), normal `%LocalAppData%`, and no development .NET on `PATH`.
- Generate a UTC cutoff immediately before the first run and pass the same explicit `Z`/`+00:00` value to aggregate validation and publishing. Every outer run, assertion, and installed subgate must fall within its linked fresh run; all outer/subgate GUIDs and whole-run manifests must be distinct.
- Keep the first failure and its evidence. A rerun may diagnose a failure but cannot erase it from the candidate record.
- A skipped, missing, flaky, or inconclusive mandatory assertion is a failure.
- The installed `ExcelDiffTracker.exe` and its public UI are the system under test. Product assemblies and `ExcelDiffTracker.Smoke.exe` are not valid black-box substitutes.
- Primary Excel save acceptance uses visible Excel, keyboard entry, and `Ctrl+S`. COM may create fixtures, disable macros, and open Excel, but it may not perform the acceptance edit or save.
- The acceptance harness must run under stock Windows PowerShell 5.1 on the clean VM; PowerShell 7 is not a prerequisite.

## Mandatory gates

### Candidate identity and environment

Record the source commit/tag; installer and installed-executable hashes; VM snapshot; Windows and Excel builds/architecture; account/admin status; VM CPU/RAM; display resolution/scaling; theme; and test-run UTC timestamps. The exact installer must be retained outside the VM.

### Install, onboarding, tray, and lifecycle

1. Install the exact candidate for the current standard user.
2. Launch the Start-menu shortcut and require `Welcome to Excel Diff Tracker`. Any `Startup problem`, Windows Error Reporting event, unexpected dialog, or early process exit fails.
3. Traverse all five onboarding steps through Windows UI Automation, exercise the folder picker, choose a real fixture through the workbook picker, create the silent baseline, and reach the populated dashboard.
4. Close the main window, reopen it from the tray, exit from the tray, relaunch, and verify onboarding does not repeat.
5. With startup enabled, log off/on and require a quiet tray start plus a working dashboard.
6. Verify in-place upgrade preserves synthetic history, then verify uninstall removes binaries, shortcuts, and startup registration while following the documented data-retention behavior.

The release-level lifecycle requirement is a separate two-phase installed gate using the real prior public installer pinned in `packaging/release-baselines.json` and a strictly newer candidate. Run [the installed lifecycle and in-place-upgrade procedure](INSTALLED_LIFECYCLE_UPGRADE.md), perform the required Windows logoff/sign-in outside the scripts, and retain exactly one validated `installed-lifecycle-upgrade.json` in the aggregate evidence tree. The aggregate validator independently binds it to the policy-pinned prior installer/application/version, candidate installer/application/version, and acceptance probe. The per-run tray/relaunch/background/repair checks remain useful supporting coverage but cannot replace this exact upgrade-and-new-logon proof.

### Real Excel `.xlsx` and `.xlsm`

For both formats, add the fixture through the installed UI while it is already open in visible Excel. Leave Excel open for at least 90 seconds after each save.

1. Select a known empty cell, type `test`, and press `Ctrl+S`.
2. Require capture within 20 seconds, exactly one new sequence, the exact empty-to-`test` delta in the UI/SQLite/Markdown, a Ready report, and no capture error.
3. Without closing Excel, save `test2`, then clear the cell. Require exactly three ordered versions and exact reports; no duplicate or collapsed stable state is allowed.
4. Exercise formula add/edit/delete, formula-result-only change, type change, sheet add/rename/reorder/hide/remove, and a style-only save that produces one `No tracked changes` version.
5. For `.xlsm`, disable macros before opening, hash `xl/vbaProject.bin` before/after, and require the hash to remain unchanged.

### Watcher and recovery matrix

- Exclusive lock for five seconds, then release: capture within 30 seconds with no error.
- Exclusive lock beyond 60 seconds: show one actionable warning, release without another save, then recover automatically within 20 seconds and return the workbook to Active without advancing from a corrupt/incomplete state.
- Atomic replacement, Save As over the tracked path, AutoSave/write burst, three distinct stable saves, two workbooks where one is locked, missing-file restoration, and a save made while the app is stopped followed by restart reconciliation.
- Make the report folder temporarily unwritable and interrupt a large capture. Restore/restart and require pending-report recovery, no duplicate, and preservation of the last valid baseline.
- Corrupt and encrypted workbooks must never advance the baseline.

Every clean critical-path run must execute the automated exclusive-lock recovery gate. The gate writes a valid changed workbook while its tracked path is held with `FileShare.None`, retains that lock until the installed product records its 60-second warning, proves the sequence and baseline hash did not advance, renders the actionable warning in History, then releases the handle without another save. Recovery must capture the one exact delta and return to Active within 20 seconds; a reconciliation settle check must still show sequence 1 with unchanged source bytes and no duplicate warning or version. `recovery/recovery.json` and its raw probe outputs are mandatory, checksummed run evidence.

Each clean run must also execute `Invoke-InstalledRecoveryMatrix.ps1` and pass `Test-InstalledRecoveryMatrixResult.ps1`. Its eleven named scenarios and exact phase sets are mandatory; missing, duplicated, reordered, or inexact scenario evidence fails the run. See [the installed recovery-matrix procedure](INSTALLED_RECOVERY_MATRIX.md).

### 500,000-cell product benchmark

This is an installed-product test, not direct extraction. Add a seeded 500,000-populated-cell fixture through the UI, open it in visible Excel, and make three saves touching cells near the beginning, middle, and end.

- Baseline and each capture complete within 60 seconds on the recorded VM.
- The WPF window never reports `Not Responding`; UI Automation heartbeat has no stall over two seconds.
- Peak private working set stays below 1.5 GiB.
- UI, SQLite, and Markdown contain the exact expected deltas once each.
- A 10,000-cell change produces the policy-truncated automatic report and a complete full export.

### Visual and accessibility matrix

Capture all onboarding steps and Dashboard, Workbooks, History, Settings, About, tray/menu, pickers, dialogs, toasts, empty/populated/processing/warning/error states, and deliberately long paths.

Required configurations include light, dark, system-light, system-dark, and Windows contrast themes; 100%, 125%, 150%, and 200% scaling; minimum/default/resized windows; 1280x720 and 1920x1080-class displays.

- Normal text contrast is at least 4.5:1; large text, focus indicators, and meaningful control boundaries are at least 3:1.
- No text/control is clipped, overlapped, or unintentionally off-screen. Intentional path ellipsis requires the complete accessible name and tooltip.
- Every action is keyboard reachable with visible focus and correct UI Automation name, role, state, and stable `AutomationId`.
- Theme changes apply immediately and persist; Windows contrast/theme changes while running are reflected.

`Capture-VisualMatrix.ps1` is phased because onboarding and a populated dashboard cannot truthfully exist at the same time in one clean profile. Use one new GUID as `-CaptureSessionId` for the entire candidate, and pass `-InstallerPath` plus the frozen installer and installed-application SHA-256 values on every invocation. The script hashes the installer path and the running installed executable instead of trusting their labels. It is append-only: it refuses a duplicate state, a different session/candidate, or a requested scale/contrast label that does not match Windows' measured state.

The onboarding phase requires a real workbook whose absolute path is at least 80 characters. It captures all five steps, the empty and populated workbook step, both native pickers, accessible long-path text, UI Automation trees, geometry, roles/names/IDs, and a screenshot at every unique Tab stop. The main phase requires a populated workbook plus an existing warning and captures every page in Light, Dark, and System themes, the file/folder pickers, and the non-destructive side of the purge confirmation. Run the main phase at minimum, default, and resized window sizes. Arrange each transient tray-menu, toast, processing, warning, and error state and capture it immediately with the supplemental phase; `-UseDesktopRoot` is available for the tray menu.

Do not infer display configuration from folder names. The capture records per-window DPI from Windows, derives scaling from that DPI, reads actual Windows app/system theme and high-contrast state, and reads the primary physical display mode. Collect the complete Light/Dark/System five-page main matrix at each measured 100%, 125%, 150%, and 200% scale; both measured Windows light and dark modes; a five-page System-theme matrix under a measured contrast theme; and physical 1280x720-class and 1920x1080-class modes. A mismatch fails before screenshots are accepted.

A representative invocation is:

```powershell
$session = [Guid]::NewGuid().ToString()
$common = @{
  EvidenceRoot = '<visual-evidence>'
  CaptureSessionId = $session
  InstallerPath = '<exact-installer.exe>'
  ExpectedInstallerSha256 = '<installer-sha256>'
  ExpectedApplicationSha256 = '<installed-exe-sha256>'
  ExpectedScalePercent = '100'
  OperatorDisplayLabel = 'operator note only; not used as evidence'
}
scripts\acceptance\Capture-VisualMatrix.ps1 @common -CapturePhase Onboarding -WorkbookPath '<80-plus-character-fixture-path>'
scripts\acceptance\Capture-VisualMatrix.ps1 @common -CapturePhase Main -ExpectedLongPath '<same-full-path>' -ExpectedWarningText '<visible-warning-text>'
```

Repeat the main phase after changing the real Windows display/theme state, update only `ExpectedScalePercent`, and use the same session/candidate identity. Use `-WindowSizeMode Minimum`, `Default`, and `Resized` across the set, and use `-WindowsContrastTheme` only while Windows actually reports a contrast theme. Supplemental calls use `-CapturePhase Supplemental -StateName <state>` after arranging the transient state; use `-UseDesktopRoot` for the open tray menu. Processing capture requires an exact visible product status of `Processing`, warning requires an exact `Warning` status, and error requires one recognized exact capture-error category/stage on the active History page; supply that exact value through `-ExpectedStateText`. Toast capture requires a visible named `StatusToast`; tray capture requires the three product-owned menu items; and long-path capture requires the accessible full path and HelpText. A state label or arbitrary product text is never accepted as evidence. Never reuse an evidence directory after a failed or partial candidate run.

Create deterministic palette evidence from the canonical `ThemeManager.cs` in the acceptance checkout. The helper rejects any alternate path, copies and hashes the canonical source, independently calculates WCAG relative luminance and ratios for every reviewed foreground/background contract in both fixed themes, and fails below 4.5:1 for normal text or 3:1 for large text and essential UI. The validator separately resolves that canonical repository path relative to its own script and requires byte-for-byte hash equality, so caller-supplied candidate labels cannot substitute a different palette:

```powershell
scripts\acceptance\New-VisualContrastEvidence.ps1 `
  -EvidenceRoot <visual-evidence> `
  -ThemeManagerSourcePath <candidate-checkout>\src\ExcelDiffTracker.App\Services\ThemeManager.cs `
  -ExpectedInstallerSha256 <installer-sha256> `
  -ExpectedApplicationSha256 <installed-exe-sha256> `
  -CaptureSessionId $session
```

This numeric artifact covers only opaque Light/Dark resource colors. It cannot prove final rendered pixels after opacity, antialiasing, font rendering, images, overlays, or compositing. Windows high-contrast colors come from live `SystemColors` and are intentionally not replaced with made-up fixed ratios. Screenshot review remains mandatory for those render-level facts.

On the disposable acceptance VM, run the lifecycle capture after the matrix while the application is open and Windows is in the declared initial state. It selects Light, Dark, and System in turn and gives each one a real executable stop/start with a new process identity. It then leaves System selected and waits for the operator to change the real Windows app theme and high-contrast setting while recording one unchanged product PID:

```powershell
scripts\acceptance\Capture-VisualLifecycleEvidence.ps1 `
  -EvidenceRoot <visual-evidence> `
  -InstallerPath <exact-installer.exe> `
  -ExpectedInstallerSha256 <installer-sha256> `
  -ExpectedApplicationSha256 <installed-exe-sha256> `
  -CaptureSessionId $session `
  -ExpectedInitialWindowsAppTheme Light `
  -ExpectedChangedWindowsAppTheme Dark `
  -ExpectedInitialHighContrast $false `
  -ExpectedChangedHighContrast $true
```

The Light/Dark direction may be reversed, but high contrast must start off and then be enabled; otherwise the preceding app-theme change is not a meaningful rendered transition. Do not run this process-restart harness on a workstation or irreplaceable session: use the resettable acceptance VM. The restart proof uses `Stop-Process` followed by `Start-Process` against the same hashed installed executable, so it proves persisted settings across new operating-system processes but not graceful tray shutdown; the separate tray lifecycle test covers graceful Exit. For the live Windows transitions, the machine proves the registry/high-contrast change, the unchanged PID/start time, System selection, and hashes the before/after screenshot and UIA tree. UI Automation does not expose WPF brush pixels, so an independent human must confirm from the screenshot pairs that the running product actually redrew.

After all phases, run:

```powershell
scripts\acceptance\Test-VisualMatrixBundle.ps1 `
  -EvidenceRoot <visual-evidence> `
  -ExpectedInstallerSha256 <installer-sha256> `
  -ExpectedApplicationSha256 <installed-exe-sha256>
```

This validator re-hashes every screenshot, UIA tree, focus manifest, focus screenshot, palette artifact, source copy, restart state, and transition state. It independently reparses palette colors and recalculates every ratio, requires new PIDs for all three theme restarts, requires both Windows transitions against one unchanged System-themed PID, and fails on missing, overwritten, mixed-session, mislabeled, unreferenced, or machine-failed evidence. It never treats `Pending` as approval.

The independent visual/accessibility reviewer must inspect every screenshot and focus sequence for rendered contrast, clipping/overlap/off-screen content, path ellipsis/tooltips, a clearly visible keyboard focus indicator, restart persistence, and visible redraw across both live Windows transitions. Approval is explicit and fail-closed:

```powershell
scripts\acceptance\Approve-VisualMatrixBundle.ps1 `
  -EvidenceRoot <visual-evidence> `
  -ExpectedInstallerSha256 <installer-sha256> `
  -ExpectedApplicationSha256 <installed-exe-sha256> `
  -Reviewer '<reviewer name>' `
  -Notes '<what was inspected and any finding disposition>' `
  -ApproveContrast -ApproveClippingAndOverlap -ApproveKeyboardAndFocus -ApproveThemeTransitions
```

Omitting any approval, reviewer identity, or substantive notes fails. Numeric palette success alone cannot substitute for rendered screenshot review. The helper first validates the complete machine bundle, then records one approval time after every contrast/lifecycle capture and binds every matrix to the exact capture session plus SHA-256 of the contrast and lifecycle artifacts. It validates the bundle again with human approval required. Replacing or regenerating either artifact invalidates the binding and requires a fresh human review; any later matrix capture also resets its matrix to pending.

### Soak

Perform at least 20 alternating real Excel saves across two workbooks over ten minutes. There must be zero missed stable hashes, duplicates, crashes, or workbooks stuck in Processing/Warning.

## Evidence bundle

Each run is stored under `artifacts/acceptance/<version>/<run-id>/` and copied off the disposable VM. It contains `acceptance.json`, raw command/automation logs, JUnit/TRX-style assertions, environment and candidate manifests, screenshots, UI Automation trees, focus/contrast/bounds output, short recordings, fixture and macro-part hashes, generated reports, a read-only database copy and query results, timing/memory telemetry, Windows Application/.NET/WER exports, and `SHA256SUMS.txt` covering every artifact.

## Independent approval

The candidate requires three independent attestations. Each attestation must identify the exact build-source commit, release commit, source-manifest hash, installer, installed executable, and common evidence digest; include substantive notes and a valid review timestamp after all machine/visual evidence completed. `approval.json` records the SHA-256 of each attestation and their combined set so later byte replacement is detectable:

1. A functional verifier who did not implement the candidate reruns the installed-app/Excel critical path from a reset snapshot.
2. A visual/accessibility reviewer checks the complete screenshot/video matrix and machine geometry/contrast results.
3. A release custodian verifies both clean runs, evidence hashes, installer identity, and no open P0-P2 findings.

After all machine and visual evidence is complete, run the aggregate validator once without attestations. It intentionally stops after writing `EVIDENCE_SHA256SUMS.txt`; its own SHA-256 is the `evidenceSha256` each reviewer must inspect and record. Each file under `attestations/` uses this shape, with the appropriate exact role and distinct reviewer:

```json
{
  "status": "Approved",
  "reviewer": "Reviewer name",
  "role": "Functional verifier",
  "sourceCommit": "40-character build-source commit",
  "releaseCommit": "40-character release commit",
  "sourceManifestSha256": "64-character installed source-manifest hash",
  "installerSha256": "64-character installer hash",
  "installedApplicationSha256": "64-character installed executable hash",
  "evidenceSha256": "SHA-256 of EVIDENCE_SHA256SUMS.txt",
  "reviewedUtc": "YYYY-MM-DDTHH:MM:SSZ",
  "notes": "Substantive description of the evidence personally reviewed."
}
```

Rerun aggregate validation only after all three reviewers have completed their files. The visual attestation reviewer must exactly match the person recorded by the visual-bundle approval helper. The builder cannot approve their own work. Release publishing must refuse to proceed unless the machine-readable acceptance summary is `Approved`, both clean critical-path runs are present, every evidence hash verifies, all attestation bytes still match their approval hashes, and all attestations name the same candidate identities.

These local hashes provide integrity and review traceability; they are not digital signatures and do not prove a human identity against a hostile user who can rewrite the entire local evidence tree. Reviewer independence is an operational control until signed attestations or an external immutable evidence service is introduced.
