# Installed semantic matrix gate

`Invoke-InstalledSemanticMatrix.ps1` is a fail-closed black-box gate for the exact installed ARM64 candidate. It is intentionally separate from unit tests and from fixture-only performance tests.

The runner creates a deterministic `.xlsx` workbook and prepares a deterministic macro-bearing `.xlsm` workbook in visible desktop Excel. It registers both through the installed application's UI and performs the complete matrix independently against each format. Every post-baseline mutation uses keyboard input or a UI Automation dialog followed by `Ctrl+S`. Excel COM is used only to create or prepare the pre-tracking fixtures, keep Excel visible, and close the fixtures at the end. It is not used for a post-baseline mutation or save.

The matrix requires these exact distinct saves:

1. formula add;
2. formula edit;
3. formula delete;
4. unchanged formula text with a changed cached result, plus its causal literal delta;
5. literal and cell-type transition;
6. style-only change with a `No tracked changes` report;
7. sheet add;
8. sheet rename;
9. sheet reorder;
10. sheet hide;
11. sheet unhide; and
12. sheet remove.

The 12-save matrix runs once for `.xlsx` and once for `.xlsm`, producing 26 total phases including the two silent baselines. Each phase retains a byte-identical workbook copy, SHA-256, external AcceptanceProbe JSON, Markdown report, desktop screenshot, and installed-app UIA tree. The evidence directory must not exist before the run. The independent validator recomputes candidate and evidence hashes and rejects missing, duplicated, reordered, or semantically inexact phases.

The supplied `.xlsm` fixture must contain exactly one worksheet and a genuine deterministic `xl/vbaProject.bin`. Macros are disabled before Excel opens it. The runner hashes the source VBA part, requires fixture preparation to preserve it, then re-extracts and compares `vbaProject.bin` from the retained workbook bytes after the baseline and every one of the 12 saves. A single changed or missing macro-part byte fails the run. The validator independently repeats those checks against the supplied source fixture and every portable `.xlsm` phase workbook. It also verifies that no `.xlsx` phase contains a VBA project.

## Runtime assumptions

- Run from interactive 64-bit Windows PowerShell 5.1 in the logged-in Windows desktop session.
- Microsoft Excel is installed and exposes the standard English desktop ribbon and dialogs used by the acceptance VM (`Move or Copy`, `Unhide`, and the standard workbook picker).
- No unrelated workbook is modified. The runner creates and uses only `fixtures\Installed Semantic Matrix.xlsx` and `fixtures\Installed Semantic Matrix.xlsm` under its fresh evidence directory.
- The source `.xlsm` fixture has exactly one worksheet. Its values and styles may be replaced during pre-tracking setup, but its `vbaProject.bin` must remain byte-identical.
- Excel Diff Tracker is already installed and onboarding is complete. Its executable, installer, database, and external AcceptanceProbe paths are supplied explicitly.
- The candidate hashes are frozen before the run. A mismatch stops the gate before workbook registration.
- The evidence path is local to the Windows VM. A UNC or Parallels shared-folder path can add watcher and Office behavior that this gate is not intended to approve.
- The runner leaves the installed application running so a larger acceptance orchestrator can continue. It closes only the Excel instance and workbook it created.

## Run and validate

```powershell
$installerHash = (Get-FileHash C:\Candidate\ExcelDiffTracker-Setup-arm64.exe -Algorithm SHA256).Hash
$applicationHash = (Get-FileHash "$env:LOCALAPPDATA\Programs\Excel Diff Tracker\ExcelDiffTracker.exe" -Algorithm SHA256).Hash

& C:\Repo\scripts\acceptance\Invoke-InstalledSemanticMatrix.ps1 `
  -InstallerPath C:\Candidate\ExcelDiffTracker-Setup-arm64.exe `
  -ExpectedInstallerSha256 $installerHash `
  -ExpectedApplicationSha256 $applicationHash `
  -ProbePath C:\Candidate\acceptance-tools\ExcelDiffTracker.AcceptanceProbe.exe `
  -XlsmFixture C:\Candidate\fixtures\DeterministicMacro.xlsm `
  -EvidenceDirectory C:\Acceptance\semantic-matrix-run-1 `
  -ConfirmInstalledCandidate

& C:\Repo\scripts\acceptance\Test-InstalledSemanticMatrixResult.ps1 `
  -ResultPath C:\Acceptance\semantic-matrix-run-1\installed-semantic-matrix.json `
  -InstallerPath C:\Candidate\ExcelDiffTracker-Setup-arm64.exe `
  -ExpectedApplicationSha256 $applicationHash `
  -ProbePath C:\Candidate\acceptance-tools\ExcelDiffTracker.AcceptanceProbe.exe `
  -XlsmFixture C:\Candidate\fixtures\DeterministicMacro.xlsm
```

A passing runner prints `INSTALLED_SEMANTIC_MATRIX_PASS`. Independent validation prints `INSTALLED_SEMANTIC_MATRIX_VALID`. Neither string is sufficient without the retained JSON and every hashed evidence file.
