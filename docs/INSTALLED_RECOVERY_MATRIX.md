# Installed watcher/recovery matrix gate

`Invoke-InstalledRecoveryMatrix.ps1` is a fail-closed, installed-app acceptance gate for watcher semantics and recovery behavior. It must run only in an interactive, disposable clean Windows VM. It drives the installed executable through UI Automation, creates and saves fixtures with real desktop Excel, and queries the installed SQLite history exclusively through the external `ExcelDiffTracker.AcceptanceProbe.exe`.

The matrix covers exactly these independent scenarios:

1. a compatible read handle and an approximately five-second exclusive lock, including recovery after release;
2. same-path atomic replacement;
3. Excel **Save As**, proving the original tracked path does not follow the new file;
4. an ordered five-state AutoSave-equivalent atomic write burst, proving final-state deduplication and order;
5. two tracked workbooks with one exclusively locked, proving the other remains independent;
6. a missing source held beyond the 60-second stable-copy timeout and later restoration;
7. a file changed while the app is stopped and captured by startup reconciliation;
8. a committed version whose report target is deliberately unusable, followed by pending-report completion after restoration;
9. an interrupted large-workbook capture held beyond the stable-copy timeout, followed by recovery;
10. corrupt Open XML package rejection without sequence or baseline-hash advance; and
11. password-protected package rejection without sequence or baseline-hash advance.

Every probe phase retains the raw external-probe JSON, a portable source-workbook copy when readable, a portable Markdown report when ready, and byte-counted SHA-256 records for copied database/WAL/SHM evidence. Each scenario also retains a desktop screenshot and UI Automation tree. The evidence directory must not exist before the run and must be local to the VM.

## Safety and prerequisites

- Use interactive 64-bit Windows PowerShell 5.1 Desktop in the logged-in desktop session.
- Use a disposable clean clone or snapshot. The runner adds test workbooks to the installed database, stops and restarts the installed process once, and creates only new fixtures/evidence beneath the supplied fresh evidence directory.
- Complete installed-app onboarding before the run so the main window and normal workbook controls are available.
- Exit Excel Diff Tracker before starting. The runner refuses to begin if the installed executable is already running.
- Install desktop Microsoft Excel. The runner uses Excel COM only to exercise real workbook/package saves; installed-app registration and report-folder changes use UI Automation.
- Freeze and supply the installer and installed-executable SHA-256 values. Any mismatch stops the run.
- Supply the separately built `ExcelDiffTracker.AcceptanceProbe.exe`; do not use an in-process or runner-owned database query.
- Do not run against the original Windows VM, a user's production database, a UNC path, or a Parallels shared folder.

Excel normally creates the encrypted fixture safely with `Workbook.SaveAs(..., Password)`. If that operation is unavailable in the acceptance image, the runner fails closed. Create a harmless password-protected `.xlsx` separately and pass it with `-PasswordProtectedFixture`; the runner records its exact SHA-256 and provenance. It never needs or opens that fixture's password.

## Run and independently validate

```powershell
$installer = 'C:\Candidate\ExcelDiffTracker-Setup-arm64.exe'
$application = "$env:LOCALAPPDATA\Programs\Excel Diff Tracker\ExcelDiffTracker.exe"
$probe = 'C:\Candidate\acceptance-tools\ExcelDiffTracker.AcceptanceProbe.exe'
$installerHash = (Get-FileHash $installer -Algorithm SHA256).Hash
$applicationHash = (Get-FileHash $application -Algorithm SHA256).Hash

& C:\Repo\scripts\acceptance\Invoke-InstalledRecoveryMatrix.ps1 `
  -InstallerPath $installer `
  -ExpectedInstallerSha256 $installerHash `
  -ExpectedApplicationSha256 $applicationHash `
  -ProbePath $probe `
  -EvidenceDirectory C:\Acceptance\installed-recovery-matrix-run-1 `
  -ConfirmDisposableCleanVm

& C:\Repo\scripts\acceptance\Test-InstalledRecoveryMatrixResult.ps1 `
  -ResultPath C:\Acceptance\installed-recovery-matrix-run-1\installed-recovery-matrix.json `
  -InstallerPath $installer `
  -ExpectedApplicationSha256 $applicationHash `
  -ProbePath $probe
```

The runner prints `INSTALLED_RECOVERY_MATRIX_PASS` only after all eleven scenarios pass. The independent validator recomputes every candidate and retained-evidence hash, reparses every raw probe, rejects missing/extra/duplicated scenarios or phases, checks exact sequence/version/error/status/report-state transitions, and prints `INSTALLED_RECOVERY_MATRIX_VALID` only for a complete bundle. Neither marker is sufficient without the result JSON and all referenced evidence files.
