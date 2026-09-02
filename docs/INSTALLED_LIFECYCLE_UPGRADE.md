# Installed lifecycle and in-place-upgrade gate

This is a fail-closed, two-phase acceptance gate for a disposable Windows VM. It installs the prior public per-user release, completes first-run onboarding with a real Excel-created `.xlsx`, records an exact visible keyboard/`Ctrl+S` change, exits through the tray, and installs the candidate over the prior version without uninstalling. It then proves the candidate payload hash, version transition, unchanged closed database bytes, preserved sequence-1 history, no repeated onboarding, a second exact real-Excel save, the exact HKCU startup value, and close-to-tray behavior.

The first script stops in `PendingExternalLogoffLogon` with the candidate installed and running only in the tray. It never invokes `logoff`, `shutdown`, `Restart-Computer`, or any equivalent. External VM control or a human must perform an actual Windows user logoff and sign back into the same account.

The second script refuses the old logon. It requires a different token logon SID, a new Explorer process started after the pre-logoff phase, and an already-running candidate process that began in the new shell interval before the script performs any product launch. It proves quiet startup (no Welcome or main window), the actual notification-area icon, native tray double-click reopen, exact dashboard/history/probe state, and tray Exit. It then uninstalls and requires binaries, Start-menu shortcut, startup value, and uninstall registration to be gone while the byte-identical local history database and exact external-probe result remain.

## Safety and prerequisites

- Use a resettable disposable Windows VM, stock 64-bit Windows PowerShell 5.1 Desktop, and a standard non-administrator account. Do not use the original VM or a production Windows profile.
- Install desktop Microsoft Excel. Both synthetic versions are made in a visible Excel window using keyboard navigation and `Ctrl+S`.
- Begin with no Excel Diff Tracker process, install directory, local data directory, default report directory, shortcut, startup value, or uninstall registration.
- Keep the prior installer, candidate installer, and separately built `ExcelDiffTracker.AcceptanceProbe.exe` on local VM paths. Freeze both installer hashes, both installed-executable hashes, the probe hash, and both versions before starting. For candidate 0.1.2, `packaging/release-baselines.json` requires the actual public v0.1.1 GitHub asset (`491F37DA829C4A167B5C1C4CD1BD908503E6A31CFC4B75819B45F0C5426FEFF6`) and its installed executable (`D9901C2709C1759A96BFCFADA71C42018575CD6781C8FF717ADA6D76A255DAD8`); aggregate release approval rejects any substitute.
- Use a fresh local evidence directory. A failed or partial directory is not reusable; reset the VM and start with a new directory.

## Phase 1: install, onboard, upgrade, and stop pending logoff

```powershell
$common = @{
  PriorInstallerPath = 'C:\Candidate\ExcelDiffTracker-Setup-0.1.1-arm64.exe'
  ExpectedPriorInstallerSha256 = '491F37DA829C4A167B5C1C4CD1BD908503E6A31CFC4B75819B45F0C5426FEFF6'
  ExpectedPriorApplicationSha256 = 'D9901C2709C1759A96BFCFADA71C42018575CD6781C8FF717ADA6D76A255DAD8'
  PriorVersion = '0.1.1'
  CandidateInstallerPath = 'C:\Candidate\ExcelDiffTracker-Setup-0.1.2-arm64.exe'
  ExpectedCandidateInstallerSha256 = '<candidate-installer-sha256>'
  ExpectedCandidateApplicationSha256 = '<candidate-installed-exe-sha256>'
  CandidateVersion = '0.1.2'
  ProbePath = 'C:\Candidate\acceptance-tools\ExcelDiffTracker.AcceptanceProbe.exe'
  ExpectedProbeSha256 = '<probe-sha256>'
  EvidenceDirectory = 'C:\Acceptance\installed-lifecycle-upgrade-run-1'
}

& C:\Repo\scripts\acceptance\Start-InstalledLifecycleUpgradeGate.ps1 @common `
  -VmSnapshotName 'edt-clean-clone-1' `
  -VmSnapshotId '<immutable-snapshot-id>' `
  -ConfirmDisposableCleanVm
```

Proceed only after `INSTALLED_LIFECYCLE_UPGRADE_PENDING_EXTERNAL_LOGOFF_LOGON` is printed. Close that PowerShell process, perform an actual Windows user logoff externally, and sign back into the same account. Do not merely lock the desktop, disconnect RDP, restart Explorer, kill the app, or run the candidate manually.

## Phase 2: prove logon startup, exit, uninstall, and retain history

Run this from a new interactive PowerShell after signing back in. Do not launch Excel Diff Tracker first.

```powershell
& C:\Repo\scripts\acceptance\Complete-InstalledLifecycleUpgradeGate.ps1 @common `
  -ConfirmDisposableVmAfterExternalLogoffLogon
```

The pass marker is `INSTALLED_LIFECYCLE_UPGRADE_PASS`. The phase writes `installed-lifecycle-upgrade.json` plus `SHA256SUMS.txt`; the checksum manifest covers every other file in the evidence tree.

## Independent validation

Validation requires the exact retained external installers and probe, not filenames or copied labels:

```powershell
& C:\Repo\scripts\acceptance\Test-InstalledLifecycleUpgradeResult.ps1 `
  -ResultPath 'C:\Acceptance\installed-lifecycle-upgrade-run-1\installed-lifecycle-upgrade.json' `
  -PriorInstallerPath $common.PriorInstallerPath `
  -ExpectedPriorApplicationSha256 $common.ExpectedPriorApplicationSha256 `
  -PriorVersion $common.PriorVersion `
  -CandidateInstallerPath $common.CandidateInstallerPath `
  -ExpectedCandidateApplicationSha256 $common.ExpectedCandidateApplicationSha256 `
  -CandidateVersion $common.CandidateVersion `
  -ProbePath $common.ProbePath
```

`INSTALLED_LIFECYCLE_UPGRADE_VALID` is emitted only when the validator recomputes the complete evidence checksum set and independently accepts the phase links, different logon identity/timestamps, distinct increasing version and payload identities, unchanged database during the in-place installer transition, exact raw probes, UI/tray checkpoints, and complete uninstall-with-retained-history proof. Missing files, a same-logon continuation, same-version repair, hash substitutions, incomplete cleanup, or modified evidence fail validation.
