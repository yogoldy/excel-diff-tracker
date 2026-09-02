[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string] $InstallerPath,
    [Parameter(Mandatory)] [string] $ProbePath,
    [Parameter(Mandatory)] [string] $XlsxFixture,
    [Parameter(Mandatory)] [string] $XlsmFixture,
    [Parameter(Mandatory)] [ValidatePattern('^[A-Fa-f0-9]{40}$')] [string] $ExpectedSourceCommit,
    [Parameter(Mandatory)] [ValidatePattern('^[A-Fa-f0-9]{64}$')] [string] $ExpectedInstallerSha256,
    [Parameter(Mandatory)] [string] $VmSnapshotName,
    [Parameter(Mandatory)] [string] $VmSnapshotId,
    [ValidatePattern('^\d+\.\d+\.\d+$')] [string] $Version = '0.1.2',
    [ValidateRange(1, 2)] [int] $RunNumber = 1,
    [string] $EvidenceRoot,
    [switch] $ConfirmCleanSnapshot
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
Import-Module (Join-Path $PSScriptRoot 'UiAutomation.psm1') -Force

if (-not $ConfirmCleanSnapshot) {
    throw 'Acceptance is fail-closed. Rerun from a named clean snapshot with -ConfirmCleanSnapshot.'
}

$repositoryRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$installer = (Resolve-Path $InstallerPath).Path
$probe = (Resolve-Path $ProbePath).Path
$sourceXlsx = (Resolve-Path $XlsxFixture).Path
$sourceXlsm = (Resolve-Path $XlsmFixture).Path
$largeBenchmarkRunner = (Resolve-Path (Join-Path $PSScriptRoot 'Invoke-LargeWorkbookBenchmark.ps1')).Path
$largeBenchmarkValidator = (Resolve-Path (Join-Path $PSScriptRoot 'Test-LargeWorkbookBenchmarkResult.ps1')).Path
$soakRunner = (Resolve-Path (Join-Path $PSScriptRoot 'Invoke-RealExcelSoak.ps1')).Path
$soakValidator = (Resolve-Path (Join-Path $PSScriptRoot 'Test-RealExcelSoakResult.ps1')).Path
$semanticMatrixRunner = (Resolve-Path (Join-Path $PSScriptRoot 'Invoke-InstalledSemanticMatrix.ps1')).Path
$semanticMatrixValidator = (Resolve-Path (Join-Path $PSScriptRoot 'Test-InstalledSemanticMatrixResult.ps1')).Path
$sourceCommit = (& git -C $repositoryRoot rev-parse HEAD 2>$null).Trim()
$runStartedUtc = [DateTime]::UtcNow
$runId = "run-$RunNumber-$($runStartedUtc.ToString('yyyyMMddTHHmmssZ'))"
if (-not $EvidenceRoot) {
    $EvidenceRoot = Join-Path $repositoryRoot "artifacts\acceptance\$Version\$runId"
}
$evidence = [System.IO.Path]::GetFullPath($EvidenceRoot)
$screenshots = Join-Path $evidence 'screenshots'
$uia = Join-Path $evidence 'uia'
$logs = Join-Path $evidence 'logs'
$fixtures = Join-Path $evidence 'fixtures'
$probeResults = Join-Path $evidence 'probe'
$recoveryResults = Join-Path $evidence 'recovery'
$lifecycleResults = Join-Path $evidence 'lifecycle'
if (Test-Path $evidence) {
    throw "Acceptance evidence directory already exists; use a fresh path so failures cannot be overwritten: $evidence"
}
New-Item -ItemType Directory -Path $evidence, $screenshots, $uia, $logs, $fixtures, $probeResults, $recoveryResults, $lifecycleResults -Force | Out-Null
Start-Transcript -Path (Join-Path $logs 'acceptance-transcript.txt') -Force | Out-Null

$assertions = [System.Collections.Generic.List[object]]::new()
$failed = $false
$installedApplicationHash = $null
$appProcess = $null
$excel = $null
$xlsxWorkbook = $null
$xlsmWorkbook = $null
$installDirectory = Join-Path $env:LOCALAPPDATA 'Programs\Excel Diff Tracker'
$application = Join-Path $installDirectory 'ExcelDiffTracker.exe'
$dataDirectory = Join-Path $env:LOCALAPPDATA 'Excel Diff Tracker'
$databasePath = Join-Path $dataDirectory 'history.db'
$startMenuShortcut = Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs\Excel Diff Tracker\Excel Diff Tracker.lnk'
$reportDirectory = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'Excel Diff Tracker Reports'
$xlsx = Join-Path $fixtures 'Acceptance.xlsx'
$xlsm = Join-Path $fixtures 'Acceptance Macro.xlsm'
$recovery = Join-Path $fixtures 'Recovery Lock.xlsx'
$recoveryChanged = Join-Path $fixtures 'Recovery Lock Changed.xlsx'
$xlsxCell = 'Z1000'
$xlsmCell = 'Z1000'
$recoveryCell = 'Y999'
Copy-Item $sourceXlsx $xlsx -Force
Copy-Item $sourceXlsm $xlsm -Force
Copy-Item $sourceXlsx $recovery -Force

function Add-Assertion {
    param([string] $Name, [bool] $Passed, [string] $EvidencePath, [string] $Detail)
    $script:assertions.Add([pscustomobject]@{
        name = $Name
        passed = $Passed
        evidence = $EvidencePath
        detail = $Detail
        utc = [DateTime]::UtcNow.ToString('O')
    })
    if (-not $Passed) { $script:failed = $true }
}

function Invoke-Step {
    param([string] $Name, [scriptblock] $Action, [string] $EvidencePath = '')
    try {
        & $Action
        Add-Assertion $Name $true $EvidencePath 'Passed'
    }
    catch {
        Add-Assertion $Name $false $EvidencePath $_.Exception.Message
        throw
    }
}

function Save-UiState {
    param([string] $Name, [System.Windows.Automation.AutomationElement] $Window)
    $png = Join-Path $screenshots "$Name.png"
    $json = Join-Path $uia "$Name.json"
    Save-DesktopScreenshot -Path $png
    Export-UiaTree -Root $Window -Path $json
}

function Get-ProductWindow {
    param([string] $Title)
    $startupProblem = Find-UiaElement -Root ([System.Windows.Automation.AutomationElement]::RootElement) -Name 'Startup problem' -Optional
    if ($startupProblem) { throw 'Excel Diff Tracker displayed the Startup problem dialog.' }
    $processId = if ($null -ne $script:appProcess -and -not $script:appProcess.HasExited) { $script:appProcess.Id } else { 0 }
    Find-UiaWindow -Title $Title -ProcessId $processId -TimeoutSeconds 20
}

function Test-ProductWindowVisible {
    $root = [System.Windows.Automation.AutomationElement]::RootElement
    $windows = $root.FindAll([System.Windows.Automation.TreeScope]::Children, [System.Windows.Automation.Condition]::TrueCondition)
    @($windows | Where-Object { $_.Current.Name -eq 'Excel Diff Tracker' -and $_.Current.ProcessId -eq $script:appProcess.Id }).Count -eq 1
}

function Get-ProductTrayIcon {
    $root = [System.Windows.Automation.AutomationElement]::RootElement
    $findIcon = {
        $items = $root.FindAll([System.Windows.Automation.TreeScope]::Descendants, [System.Windows.Automation.Condition]::TrueCondition)
        @($items | Where-Object {
            $_.Current.AutomationId -eq 'NotifyItemIcon' -and $_.Current.Name -like 'Excel Diff Tracker*'
        }) | Select-Object -First 1
    }
    $icon = & $findIcon
    if (-not $icon) {
        $hiddenIcons = Find-UiaElement -Root $root -AutomationId 'SystemTrayIcon' -Name 'Show Hidden Icons'
        Invoke-UiaElement -Element $hiddenIcons
        Start-Sleep -Milliseconds 500
        $icon = & $findIcon
    }
    if (-not $icon) { throw 'Excel Diff Tracker tray icon was not found in the notification area.' }
    $icon
}

function Close-ProductWindowToTray {
    $window = Get-ProductWindow 'Excel Diff Tracker'
    $pattern = $null
    if (-not $window.TryGetCurrentPattern([System.Windows.Automation.WindowPattern]::Pattern, [ref]$pattern)) {
        throw 'The main window does not expose WindowPattern.'
    }
    ([System.Windows.Automation.WindowPattern]$pattern).Close()
    Wait-AcceptanceCondition -TimeoutSeconds 5 -FailureMessage 'The main window did not hide after Close.' -Condition { -not (Test-ProductWindowVisible) }
    if ($script:appProcess.HasExited) { throw 'Closing the main window exited the tray application.' }
}

function Exit-ProductThroughTray {
    $icon = Get-ProductTrayIcon
    Invoke-UiaMouseClick -Element $icon -Button Right
    $root = [System.Windows.Automation.AutomationElement]::RootElement
    $items = $root.FindAll([System.Windows.Automation.TreeScope]::Descendants, [System.Windows.Automation.Condition]::TrueCondition)
    $exit = @($items | Where-Object {
        $_.Current.Name -eq 'Exit' -and $_.Current.ControlType -eq [System.Windows.Automation.ControlType]::MenuItem
    }) | Select-Object -First 1
    if (-not $exit) { throw 'The tray Exit menu item was not found.' }
    Invoke-UiaElement -Element $exit
    Wait-AcceptanceCondition -TimeoutSeconds 10 -FailureMessage 'The app did not exit from the tray menu.' -Condition {
        $script:appProcess.Refresh()
        $script:appProcess.HasExited
    }
}

function Click-ProductControl {
    param([string] $WindowTitle, [string] $AutomationId)
    $window = Get-ProductWindow $WindowTitle
    $element = Find-UiaElement -Root $window -AutomationId $AutomationId
    Invoke-UiaElement -Element $element
}

function Choose-FileFromDialog {
    param([string] $DialogTitle, [string] $Path)
    $dialog = Find-UiaWindow -Title $DialogTitle -TimeoutSeconds 15
    Save-UiState 'workbook-picker' $dialog
    $fileName = Find-UiaElement -Root $dialog -AutomationId '1148' -Optional
    if (-not $fileName) { $fileName = Find-UiaElement -Root $dialog -Name 'File name:' }
    Set-UiaValue -Element $fileName -Value $Path
    $open = Find-UiaElement -Root $dialog -AutomationId '1' -Optional
    if (-not $open) { $open = Find-UiaElement -Root $dialog -Name 'Open' }
    Invoke-UiaElement -Element $open
}

function Exercise-FolderPicker {
    $dialog = Find-UiaWindow -Title 'Choose where Markdown reports should be saved' -TimeoutSeconds 15
    Save-UiState 'report-folder-picker' $dialog
    $cancel = Find-UiaElement -Root $dialog -Name 'Cancel'
    Invoke-UiaElement -Element $cancel
}

function Focus-ExcelCell {
    param([object] $Workbook, [string] $Address)
    $window = Get-UiaWindowFromHandle -Handle ([long]$script:excel.Hwnd)
    Set-UiaForeground -Window $window
    [System.Windows.Forms.SendKeys]::SendWait('^g')
    Start-Sleep -Milliseconds 200
    [System.Windows.Forms.SendKeys]::SendWait($Address)
    [System.Windows.Forms.SendKeys]::SendWait('{ENTER}')
    Start-Sleep -Milliseconds 200
}

function Save-ExcelLiteral {
    param([string] $Address, [AllowEmptyString()] [string] $Value, [switch] $Clear)
    Focus-ExcelCell -Workbook $null -Address $Address
    if ($Clear) { [System.Windows.Forms.SendKeys]::SendWait('{DELETE}') }
    else { [System.Windows.Forms.SendKeys]::SendWait($Value) }
    [System.Windows.Forms.SendKeys]::SendWait('{ENTER}')
    [System.Windows.Forms.SendKeys]::SendWait('^s')
    Start-Sleep -Milliseconds 500
}

function Invoke-Probe {
    param(
        [string] $Workbook,
        [long] $Sequence,
        [string] $Address,
        [string] $ExpectedValue,
        [string] $ExpectedBeforeValue,
        [string] $ExpectedKind,
        [switch] $ExpectCleared,
        [switch] $RequireNoErrors,
        [int] $TimeoutSeconds = 20
    )
    $outputPath = Join-Path $probeResults "$(Split-Path $Workbook -Leaf)-sequence-$Sequence.json"
    $deadline = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
    $lastOutput = ''
    do {
        $arguments = @(
            '--database', $databasePath,
            '--workbook', $Workbook,
            '--expected-sequence', $Sequence,
            '--address', $Address,
            '--expected-kind', $ExpectedKind,
            '--expected-version-count', $Sequence,
            '--require-unique-version-hashes',
            '--require-source-hash-match',
            '--expected-cell-change-count', '1',
            '--expected-sheet-change-count', '0',
            '--require-active',
            '--require-no-last-error',
            '--require-ready-report',
            '--report-contains', $Address)
        if ($RequireNoErrors) { $arguments += '--require-no-errors' }
        if ($Sequence -eq 1) { $arguments += '--expect-before-missing' }
        elseif ($null -ne $ExpectedBeforeValue) { $arguments += @('--expected-before-value', $ExpectedBeforeValue) }
        if ($ExpectCleared) { $arguments += '--expect-cleared' }
        else { $arguments += @('--expected-value', $ExpectedValue) }
        $lastOutput = & $probe @arguments 2>&1 | Out-String
        $exitCode = $LASTEXITCODE
        if ($exitCode -eq 0) {
            [System.IO.File]::WriteAllText($outputPath, $lastOutput, [System.Text.UTF8Encoding]::new($false))
            return
        }
        Start-Sleep -Milliseconds 500
    } while ([DateTime]::UtcNow -lt $deadline)
    [System.IO.File]::WriteAllText($outputPath, $lastOutput, [System.Text.UTF8Encoding]::new($false))
    throw "Acceptance probe did not pass within $TimeoutSeconds seconds. See $outputPath"
}

function Wait-ProbeResult {
    param(
        [Parameter(Mandatory)] [string[]] $ProbeArguments,
        [Parameter(Mandatory)] [string] $OutputPath,
        [int] $TimeoutSeconds = 20
    )
    $deadline = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
    $lastOutput = ''
    do {
        $lastOutput = & $probe @ProbeArguments 2>&1 | Out-String
        $exitCode = $LASTEXITCODE
        if ($exitCode -eq 0) {
            [System.IO.File]::WriteAllText($OutputPath, $lastOutput, [System.Text.UTF8Encoding]::new($false))
            return ($lastOutput | ConvertFrom-Json)
        }
        Start-Sleep -Milliseconds 200
    } while ([DateTime]::UtcNow -lt $deadline)
    [System.IO.File]::WriteAllText($OutputPath, $lastOutput, [System.Text.UTF8Encoding]::new($false))
    throw "Acceptance probe did not pass within $TimeoutSeconds seconds. See $OutputPath"
}

function Prepare-RecoveryFixtures {
    $baselineWorkbook = $null
    $changedWorkbook = $null
    try {
        $baselineWorkbook = $excel.Workbooks.Open($recovery)
        $baselineWorkbook.Worksheets.Item(1).Range($recoveryCell).ClearContents()
        $baselineWorkbook.Save()
        $baselineWorkbook.Close($false)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($baselineWorkbook)
        $baselineWorkbook = $null

        Copy-Item $recovery $recoveryChanged -Force
        $changedWorkbook = $excel.Workbooks.Open($recoveryChanged)
        $changedWorkbook.Worksheets.Item(1).Range($recoveryCell).Value2 = 'locked-recovery'
        $changedWorkbook.Save()
        $changedWorkbook.Close($false)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($changedWorkbook)
        $changedWorkbook = $null
    }
    finally {
        if ($changedWorkbook) {
            try { $changedWorkbook.Close($false) } catch { }
            try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($changedWorkbook) } catch { }
        }
        if ($baselineWorkbook) {
            try { $baselineWorkbook.Close($false) } catch { }
            try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($baselineWorkbook) } catch { }
        }
    }
}

function Invoke-LockedRecoveryGate {
    $resultPath = Join-Path $recoveryResults 'recovery.json'
    $lockStream = $null
    $result = [ordered]@{
        scenarioId = 'held-open-exclusive-lock-over-60s'
        status = 'Failed'
        workbook = $recovery
        changedCandidate = $recoveryChanged
        address = $recoveryCell
        expectedValue = 'locked-recovery'
        baseline = $null
        warning = $null
        recovered = $null
        settled = $null
        hashes = [ordered]@{
            candidate = $null
            sourceAfterRelease = $null
            sourceAfterRecovery = $null
        }
        timing = [ordered]@{
            lockAcquiredUtc = $null
            lockedWriteCompletedUtc = $null
            warningObservedUtc = $null
            lockReleasedUtc = $null
            recoveredUtc = $null
            lockedDurationSeconds = $null
            recoverySeconds = $null
        }
        warningUi = [ordered]@{
            categoryCount = 0
            messageCount = 0
            workbookPathCount = 0
        }
        checks = [ordered]@{
            baselineAtSequenceZero = $false
            lockHeldBeyond60Seconds = $false
            warningRecordedExactlyOnce = $false
            baselinePreservedDuringWarning = $false
            actionableWarningRenderedExactlyOnce = $false
            lockedBytesMatchChangedCandidate = $false
            recoveredWithin20Seconds = $false
            exactDeltaCaptured = $false
            returnedToActive = $false
            noDuplicateAfterReconciliation = $false
            noFileMutationAfterRelease = $false
        }
        failure = $null
    }

    try {
        Set-UiaForeground -Window $main
        Click-ProductControl 'Excel Diff Tracker' 'DashboardAddWorkbookButton'
        Choose-FileFromDialog 'Choose a workbook to track' $recovery
        $baselinePath = Join-Path $recoveryResults 'baseline.json'
        $baseline = Wait-ProbeResult -ProbeArguments @(
            '--database', $databasePath,
            '--workbook', $recovery,
            '--expected-sequence', '0',
            '--require-active',
            '--require-no-errors',
            '--require-no-last-error') -OutputPath $baselinePath -TimeoutSeconds 60
        $result.baseline = $baseline
        $result.checks.baselineAtSequenceZero = $baseline.passed -and $baseline.currentSequence -eq 0 -and $baseline.errorCount -eq 0
        if (-not $result.checks.baselineAtSequenceZero) {
            throw 'Recovery fixture did not establish a clean sequence-zero baseline.'
        }

        $changedBytes = [System.IO.File]::ReadAllBytes($recoveryChanged)
        $result.hashes.candidate = Get-AcceptanceFileSha256 -Path $recoveryChanged
        $lockStream = [System.IO.FileStream]::new(
            $recovery,
            [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::ReadWrite,
            [System.IO.FileShare]::None)
        $lockAcquired = [DateTime]::UtcNow
        $result.timing.lockAcquiredUtc = $lockAcquired.ToString('O')
        $lockStream.SetLength(0)
        $lockStream.Write($changedBytes, 0, $changedBytes.Length)
        $lockStream.Flush($true)
        $lockedWriteCompleted = [DateTime]::UtcNow
        $result.timing.lockedWriteCompletedUtc = $lockedWriteCompleted.ToString('O')

        $warningPath = Join-Path $recoveryResults 'warning.json'
        $warning = Wait-ProbeResult -ProbeArguments @(
            '--database', $databasePath,
            '--workbook', $recovery,
            '--expected-sequence', '0',
            '--minimum-errors', '1') -OutputPath $warningPath -TimeoutSeconds 80
        $warningObserved = [DateTime]::UtcNow
        $result.warning = $warning
        $result.timing.warningObservedUtc = $warningObserved.ToString('O')
        $result.checks.warningRecordedExactlyOnce = $warning.errorCount -eq 1
        $result.checks.baselinePreservedDuringWarning =
            $warning.currentSequence -eq 0 -and
            $warning.currentHash -eq $baseline.currentHash

        Click-ProductControl 'Excel Diff Tracker' 'HistoryNavigationButton'
        $main = Get-ProductWindow 'Excel Diff Tracker'
        Wait-AcceptanceCondition -TimeoutSeconds 5 -FailureMessage 'The recovery warning was recorded but did not render in History.' -Condition {
            $nodes = $main.FindAll(
                [System.Windows.Automation.TreeScope]::Descendants,
                [System.Windows.Automation.Condition]::TrueCondition)
            @($nodes | Where-Object { $_.Current.Name -like '*Workbook temporarily unavailable*Waiting for a stable save*' }).Count -eq 1
        }
        Save-UiState 'locked-recovery-warning' $main
        $warningNodes = $main.FindAll(
            [System.Windows.Automation.TreeScope]::Descendants,
            [System.Windows.Automation.Condition]::TrueCondition)
        $result.warningUi.categoryCount = @($warningNodes | Where-Object {
            $_.Current.Name -like '*Workbook temporarily unavailable*Waiting for a stable save*'
        }).Count
        $result.warningUi.messageCount = @($warningNodes | Where-Object {
            $_.Current.Name -like '*did not become readable and stable within 60 seconds*'
        }).Count
        $result.warningUi.workbookPathCount = @($warningNodes | Where-Object {
            $_.Current.Name -eq $recovery
        }).Count
        $result.checks.actionableWarningRenderedExactlyOnce =
            $result.warningUi.categoryCount -eq 1 -and
            $result.warningUi.messageCount -ge 1 -and
            $result.warningUi.workbookPathCount -ge 1

        $lockStream.Dispose()
        $lockStream = $null
        $lockReleased = [DateTime]::UtcNow
        $result.timing.lockReleasedUtc = $lockReleased.ToString('O')
        $result.timing.lockedDurationSeconds = ($lockReleased - $lockAcquired).TotalSeconds
        $result.checks.lockHeldBeyond60Seconds = $result.timing.lockedDurationSeconds -ge 60
        $result.hashes.sourceAfterRelease = Get-AcceptanceFileSha256 -Path $recovery
        $result.checks.lockedBytesMatchChangedCandidate = $result.hashes.sourceAfterRelease -eq $result.hashes.candidate

        $recoveredPath = Join-Path $recoveryResults 'recovered.json'
        $recovered = Wait-ProbeResult -ProbeArguments @(
            '--database', $databasePath,
            '--workbook', $recovery,
            '--expected-sequence', '1',
            '--address', $recoveryCell,
            '--expected-value', 'locked-recovery',
            '--expected-kind', 'LiteralAdded',
            '--report-contains', $recoveryCell,
            '--require-active',
            '--require-no-last-error',
            '--require-ready-report',
            '--minimum-errors', '1') -OutputPath $recoveredPath -TimeoutSeconds 20
        $recoveredUtc = [DateTime]::UtcNow
        $result.recovered = $recovered
        $result.timing.recoveredUtc = $recoveredUtc.ToString('O')
        $result.timing.recoverySeconds = ($recoveredUtc - $lockReleased).TotalSeconds
        $result.checks.recoveredWithin20Seconds = $result.timing.recoverySeconds -le 20
        $result.checks.exactDeltaCaptured =
            $recovered.currentSequence -eq 1 -and
            $recovered.latestVersion.cellChangeCount -eq 1 -and
            $recovered.latestVersion.sheetChangeCount -eq 0 -and
            $recovered.cellChange.address -eq $recoveryCell -and
            $recovered.cellChange.kinds.Split(',') -contains 'LiteralAdded'
        $result.checks.returnedToActive =
            $recovered.workbookStatus -eq 'Active' -and
            [string]::IsNullOrWhiteSpace($recovered.lastError) -and
            $recovered.errorCount -eq 1

        $result.hashes.sourceAfterRecovery = Get-AcceptanceFileSha256 -Path $recovery
        Start-Sleep -Seconds 12
        $settledPath = Join-Path $recoveryResults 'settled.json'
        $settled = Wait-ProbeResult -ProbeArguments @(
            '--database', $databasePath,
            '--workbook', $recovery,
            '--expected-sequence', '1',
            '--require-active',
            '--require-no-last-error',
            '--require-ready-report',
            '--minimum-errors', '1') -OutputPath $settledPath -TimeoutSeconds 5
        $result.settled = $settled
        $result.checks.noDuplicateAfterReconciliation =
            $settled.currentSequence -eq 1 -and
            $settled.currentHash -eq $recovered.currentHash -and
            $settled.errorCount -eq 1
        $result.checks.noFileMutationAfterRelease =
            $result.hashes.sourceAfterRelease -eq $result.hashes.sourceAfterRecovery -and
            $result.hashes.sourceAfterRecovery -eq $settled.currentHash

        $failedChecks = @($result.checks.GetEnumerator() | Where-Object { -not $_.Value })
        if ($failedChecks.Count -ne 0) {
            throw "Locked recovery checks failed: $($failedChecks.Key -join ', ')"
        }
        $result.status = 'Passed'
        Click-ProductControl 'Excel Diff Tracker' 'DashboardNavigationButton'
    }
    catch {
        $result.status = 'Failed'
        $result.failure = $_.Exception.ToString()
        throw
    }
    finally {
        if ($lockStream) {
            try { $lockStream.Dispose() } catch { }
        }
        Write-AcceptanceUtf8File -Path $resultPath -Content ($result | ConvertTo-Json -Depth 12)
    }
}

function Invoke-LifecycleGate {
    $resultPath = Join-Path $lifecycleResults 'lifecycle.json'
    $checks = [ordered]@{
        closeKeepsTrayProcessAlive = $false
        actualTrayIconReopensWindow = $false
        trayExitStopsProcess = $false
        backgroundLaunchIsQuiet = $false
        secondLaunchActivatesWindow = $false
        onboardingDoesNotRepeat = $false
        repairInstallPreservesHistory = $false
        startupRegistrationIsExact = $false
    }
    $failure = $null
    try {
        Close-ProductWindowToTray
        $checks.closeKeepsTrayProcessAlive = -not $script:appProcess.HasExited
        $icon = Get-ProductTrayIcon
        Save-DesktopScreenshot -Path (Join-Path $screenshots 'tray-hidden-main.png')
        Invoke-UiaMouseClick -Element $icon -Button DoubleLeft
        $script:main = Get-ProductWindow 'Excel Diff Tracker'
        $checks.actualTrayIconReopensWindow = Test-ProductWindowVisible
        Save-UiState 'tray-reopened-main' $script:main

        Close-ProductWindowToTray
        Save-DesktopScreenshot -Path (Join-Path $screenshots 'tray-exit-menu.png')
        Exit-ProductThroughTray
        $checks.trayExitStopsProcess = $script:appProcess.HasExited

        $script:appProcess = Start-Process $application -ArgumentList '--background' -PassThru
        Wait-AcceptanceCondition -TimeoutSeconds 10 -FailureMessage 'Background startup did not create the tray process.' -Condition {
            try { -not $script:appProcess.HasExited -and $null -ne (Get-ProductTrayIcon) } catch { $false }
        }
        Start-Sleep -Seconds 2
        $checks.backgroundLaunchIsQuiet = -not (Test-ProductWindowVisible)
        if (-not $checks.backgroundLaunchIsQuiet) { throw 'The --background startup displayed the main window.' }

        $activation = Start-Process $application -PassThru
        $null = $activation.WaitForExit(5000)
        $script:main = Get-ProductWindow 'Excel Diff Tracker'
        $checks.secondLaunchActivatesWindow = Test-ProductWindowVisible
        $welcome = Find-UiaElement -Root ([System.Windows.Automation.AutomationElement]::RootElement) -Name 'Welcome to Excel Diff Tracker' -Optional
        $checks.onboardingDoesNotRepeat = $null -eq $welcome
        if (-not $checks.onboardingDoesNotRepeat) { throw 'Onboarding repeated after relaunch.' }

        $startupValue = (Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run' -Name ExcelDiffTracker -ErrorAction Stop).ExcelDiffTracker
        $expectedStartup = '"' + $application + '" --background'
        $checks.startupRegistrationIsExact = [string]::Equals($startupValue, $expectedStartup, [StringComparison]::OrdinalIgnoreCase)
        if (-not $checks.startupRegistrationIsExact) { throw "Startup registration is wrong: $startupValue" }

        Close-ProductWindowToTray
        Exit-ProductThroughTray
        $repair = Start-Process $installer -ArgumentList '/VERYSILENT','/CURRENTUSER','/SUPPRESSMSGBOXES','/NORESTART','/TASKS="startup"' -Wait -PassThru
        if ($repair.ExitCode -ne 0) { throw "Repair install exited with $($repair.ExitCode)." }
        $script:appProcess = Start-Process $application -PassThru
        $script:main = Get-ProductWindow 'Excel Diff Tracker'
        $repairProbePath = Join-Path $lifecycleResults 'repair-history.json'
        $repairProbe = Wait-ProbeResult -ProbeArguments @(
            '--database', $databasePath,
            '--workbook', $xlsx,
            '--expected-sequence', '3',
            '--expected-version-count', '3',
            '--require-unique-version-hashes',
            '--require-active',
            '--require-no-errors',
            '--require-no-last-error') -OutputPath $repairProbePath -TimeoutSeconds 20
        $checks.repairInstallPreservesHistory = $repairProbe.currentSequence -eq 3 -and $repairProbe.versionCount -eq 3

        $failedChecks = @($checks.GetEnumerator() | Where-Object { -not $_.Value })
        if ($failedChecks.Count -ne 0) { throw "Lifecycle checks failed: $($failedChecks.Name -join ', ')" }
    }
    catch {
        $failure = $_.Exception.ToString()
        throw
    }
    finally {
        $lifecycle = [ordered]@{
            schemaVersion = 1
            gate = 'installed-app-lifecycle'
            status = if ($null -eq $failure -and @($checks.GetEnumerator() | Where-Object { -not $_.Value }).Count -eq 0) { 'Passed' } else { 'Failed' }
            checks = $checks
            startupValue = (Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run' -Name ExcelDiffTracker -ErrorAction SilentlyContinue).ExcelDiffTracker
            failure = $failure
        }
        Write-AcceptanceUtf8File -Path $resultPath -Content ($lifecycle | ConvertTo-Json -Depth 8)
    }
}

function Get-ZipEntrySha256 {
    param([string] $Path, [string] $EntryName)
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $entry = $archive.GetEntry($EntryName)
        if (-not $entry) { throw "Archive entry not found: $EntryName" }
        $stream = $entry.Open()
        try {
            Get-AcceptanceStreamSha256 -Stream $stream
        }
        finally { $stream.Dispose() }
    }
    finally { $archive.Dispose() }
}

function Get-PeMachine {
    param([string] $Path)
    $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    $reader = New-Object System.IO.BinaryReader -ArgumentList $stream
    try {
        $stream.Position = 0x3c
        $peOffset = $reader.ReadInt32()
        $stream.Position = $peOffset + 4
        '0x{0:X4}' -f $reader.ReadUInt16()
    }
    finally { $reader.Dispose(); $stream.Dispose() }
}

function Write-EnvironmentManifest {
    $excelPath = (Get-Command excel.exe -ErrorAction SilentlyContinue).Source
    $installedHash = if (Test-Path $application -PathType Leaf) { Get-AcceptanceFileSha256 -Path $application } else { $null }
    $excelDetails = if ($null -ne $script:excel) {
        [ordered]@{
            version = [string]$script:excel.Version
            build = [string]$script:excel.Build
            operatingSystem = [string]$script:excel.OperatingSystem
        }
    } else { $null }
    $excelExecutable = if ($null -ne $script:excel) { Join-Path ([string]$script:excel.Path) 'EXCEL.EXE' } else { $excelPath }
    $appliedDpi = (Get-ItemProperty 'HKCU:\Control Panel\Desktop\WindowMetrics' -Name AppliedDPI -ErrorAction SilentlyContinue).AppliedDPI
    $manifest = [ordered]@{
        runId = $runId
        runNumber = $RunNumber
        startedUtc = $runStartedUtc.ToString('O')
        vmSnapshotName = $VmSnapshotName
        vmSnapshotId = $VmSnapshotId
        computerName = $env:COMPUTERNAME
        user = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
        isAdministrator = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        os = Get-CimInstance Win32_OperatingSystem | Select-Object Caption, Version, BuildNumber, OSArchitecture
        computer = Get-CimInstance Win32_ComputerSystem | Select-Object Manufacturer, Model, NumberOfLogicalProcessors, TotalPhysicalMemory
        video = Get-CimInstance Win32_VideoController | Select-Object Name, CurrentHorizontalResolution, CurrentVerticalResolution
        scalePercent = if ($appliedDpi) { [int][math]::Round(([double]$appliedDpi / 96.0) * 100.0) } else { $null }
        highContrast = [System.Windows.Forms.SystemInformation]::HighContrast
        dotnetOnPath = [bool](Get-Command dotnet -ErrorAction SilentlyContinue)
        windowsPowerShell = $PSVersionTable.PSVersion.ToString()
        excelPath = $excelPath
        excelExecutable = $excelExecutable
        excelPeMachine = if ($excelExecutable -and (Test-Path $excelExecutable)) { Get-PeMachine $excelExecutable } else { $null }
        excel = $excelDetails
        installerPath = $installer
        installerSha256 = Get-AcceptanceFileSha256 -Path $installer
        installedApplicationPath = $application
        installedApplicationSha256 = $installedHash
        sourceCommit = $sourceCommit
    }
    Write-AcceptanceUtf8File -Path (Join-Path $evidence 'environment.json') -Content ($manifest | ConvertTo-Json -Depth 8)
}

try {
    Write-EnvironmentManifest
    $installerHash = Get-AcceptanceFileSha256 -Path $installer
    if ($installerHash -ne $ExpectedInstallerSha256.ToUpperInvariant()) {
        throw "Installer SHA-256 $installerHash does not match expected $ExpectedInstallerSha256."
    }
    if ($sourceCommit -ne $ExpectedSourceCommit.ToLowerInvariant()) {
        throw "Source commit $sourceCommit does not match expected $ExpectedSourceCommit."
    }
    $dirtySource = @(& git -C $repositoryRoot status --porcelain --untracked-files=all 2>$null)
    if ($dirtySource.Count -ne 0) {
        throw 'The candidate source tree is dirty; freeze and commit the source before acceptance.'
    }
    if ([string]::IsNullOrWhiteSpace($VmSnapshotName) -or [string]::IsNullOrWhiteSpace($VmSnapshotId)) {
        throw 'A non-empty clean VM snapshot name and ID are required.'
    }
    $isAdministrator = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if ($isAdministrator) { throw 'Acceptance must run as a standard non-administrator user.' }
    $operatingSystem = Get-CimInstance Win32_OperatingSystem
    if ($operatingSystem.OSArchitecture -notlike 'ARM*') { throw "Acceptance requires Windows ARM64; found $($operatingSystem.OSArchitecture)." }
    if ($PSVersionTable.PSVersion.Major -ne 5) { throw "Acceptance requires stock Windows PowerShell 5.1; found $($PSVersionTable.PSVersion)." }
    if (Get-Command dotnet -ErrorAction SilentlyContinue) {
        throw 'dotnet is available on PATH; this clean self-contained acceptance run is invalid.'
    }
    if (Test-Path $installDirectory) { throw "Existing installation found: $installDirectory" }
    if (Test-Path $dataDirectory) { throw "Existing app data found: $dataDirectory" }

    Invoke-Step 'silent per-user install' {
        $process = Start-Process $installer -ArgumentList '/VERYSILENT','/CURRENTUSER','/SUPPRESSMSGBOXES','/NORESTART','/TASKS="startup"' -Wait -PassThru
        if ($process.ExitCode -ne 0) { throw "Installer exited with $($process.ExitCode)." }
        if (-not (Test-Path $application)) { throw 'Installed executable is missing.' }
        if (-not (Test-Path $startMenuShortcut)) { throw 'Start-menu shortcut is missing.' }
    } 'environment.json'
    $installedApplicationHash = Get-AcceptanceFileSha256 -Path $application

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3
    $excelExecutable = Join-Path ([string]$excel.Path) 'EXCEL.EXE'
    $excelMachine = Get-PeMachine $excelExecutable
    if ($excelMachine -notin @('0xAA64', '0x8664')) { throw "Acceptance requires desktop ARM64 or x64 Excel on Windows ARM64; PE machine is $excelMachine at $excelExecutable." }
    $xlsxWorkbook = $excel.Workbooks.Open($xlsx)
    $xlsxWorkbook.Worksheets.Item(1).Range($xlsxCell).ClearContents()
    $xlsxWorkbook.Save()
    Prepare-RecoveryFixtures
    Write-EnvironmentManifest

    $appProcess = Start-Process $application -PassThru
    $onboarding = Get-ProductWindow 'Welcome to Excel Diff Tracker'
    Save-UiState 'onboarding-step-1' $onboarding
    Invoke-Step 'onboarding step 1 renders' { Find-UiaElement -Root $onboarding -AutomationId 'OnboardingNextButton' | Out-Null } 'screenshots/onboarding-step-1.png'
    Click-ProductControl 'Welcome to Excel Diff Tracker' 'OnboardingNextButton'

    $onboarding = Get-ProductWindow 'Welcome to Excel Diff Tracker'
    Save-UiState 'onboarding-step-2' $onboarding
    Invoke-Step 'report folder picker opens' {
        Click-ProductControl 'Welcome to Excel Diff Tracker' 'ChooseReportFolderButton'
        Exercise-FolderPicker
    } 'screenshots/report-folder-picker.png'
    Click-ProductControl 'Welcome to Excel Diff Tracker' 'OnboardingNextButton'

    $onboarding = Get-ProductWindow 'Welcome to Excel Diff Tracker'
    Save-UiState 'onboarding-step-3' $onboarding
    Invoke-Step 'open xlsx is selected through workbook picker' {
        Click-ProductControl 'Welcome to Excel Diff Tracker' 'ChooseWorkbookButton'
        Choose-FileFromDialog 'Choose a workbook to track' $xlsx
    } 'screenshots/workbook-picker.png'
    Click-ProductControl 'Welcome to Excel Diff Tracker' 'OnboardingNextButton'

    $onboarding = Get-ProductWindow 'Welcome to Excel Diff Tracker'
    Save-UiState 'onboarding-step-4' $onboarding
    Click-ProductControl 'Welcome to Excel Diff Tracker' 'OnboardingNextButton'
    Wait-AcceptanceCondition -TimeoutSeconds 60 -FailureMessage 'Onboarding did not create the baseline.' -Condition {
        try {
            $window = Get-ProductWindow 'Welcome to Excel Diff Tracker'
            [bool](Find-UiaElement -Root $window -Name 'Tracking is active' -Optional)
        } catch { $false }
    }
    $onboarding = Get-ProductWindow 'Welcome to Excel Diff Tracker'
    Save-UiState 'onboarding-step-5' $onboarding
    Click-ProductControl 'Welcome to Excel Diff Tracker' 'OnboardingNextButton'

    $main = Get-ProductWindow 'Excel Diff Tracker'
    Save-UiState 'dashboard-baseline' $main
    Add-Assertion 'all onboarding steps completed' $true 'screenshots/onboarding-step-5.png' 'Reached installed dashboard.'

    Invoke-Step 'exclusive-lock timeout recovers without another save' {
        Invoke-LockedRecoveryGate
    } 'recovery/recovery.json'

    Invoke-Step 'xlsx save 1 while Excel remains open' {
        Save-ExcelLiteral -Address $xlsxCell -Value 'test'
        Invoke-Probe -Workbook $xlsx -Sequence 1 -Address $xlsxCell -ExpectedValue 'test' -ExpectedKind 'LiteralAdded' -RequireNoErrors
    } 'probe/Acceptance.xlsx-sequence-1.json'
    Invoke-Step 'xlsx save 2 while Excel remains open' {
        Save-ExcelLiteral -Address $xlsxCell -Value 'test2'
        Invoke-Probe -Workbook $xlsx -Sequence 2 -Address $xlsxCell -ExpectedValue 'test2' -ExpectedBeforeValue 'test' -ExpectedKind 'LiteralChanged' -RequireNoErrors
    } 'probe/Acceptance.xlsx-sequence-2.json'
    Invoke-Step 'xlsx save 3 clear while Excel remains open' {
        Save-ExcelLiteral -Address $xlsxCell -Value '' -Clear
        Invoke-Probe -Workbook $xlsx -Sequence 3 -Address $xlsxCell -ExpectedValue '' -ExpectedBeforeValue 'test2' -ExpectedKind 'LiteralCleared' -ExpectCleared -RequireNoErrors
    } 'probe/Acceptance.xlsx-sequence-3.json'

    $xlsmWorkbook = $excel.Workbooks.Open($xlsm)
    $xlsmWorkbook.Worksheets.Item(1).Range($xlsmCell).ClearContents()
    $xlsmWorkbook.Save()
    $macroHashBefore = Get-ZipEntrySha256 $xlsm 'xl/vbaProject.bin'
    Set-UiaForeground -Window $main
    Invoke-Step 'open xlsm is added through installed app' {
        Click-ProductControl 'Excel Diff Tracker' 'DashboardAddWorkbookButton'
        Choose-FileFromDialog 'Choose a workbook to track' $xlsm
        Wait-AcceptanceCondition -TimeoutSeconds 60 -FailureMessage 'The open xlsm baseline was not added.' -Condition {
            $output = & $probe '--database' $databasePath '--workbook' $xlsm '--expected-sequence' '0' '--require-active' '--require-no-errors' '--require-no-last-error' 2>&1
            $LASTEXITCODE -eq 0
        }
    } 'screenshots/workbook-picker.png'

    Invoke-Step 'xlsm save 1 while Excel remains open' {
        Save-ExcelLiteral -Address $xlsmCell -Value 'test'
        Invoke-Probe -Workbook $xlsm -Sequence 1 -Address $xlsmCell -ExpectedValue 'test' -ExpectedKind 'LiteralAdded' -RequireNoErrors
    } 'probe/Acceptance Macro.xlsm-sequence-1.json'
    Invoke-Step 'xlsm save 2 while Excel remains open' {
        Save-ExcelLiteral -Address $xlsmCell -Value 'test2'
        Invoke-Probe -Workbook $xlsm -Sequence 2 -Address $xlsmCell -ExpectedValue 'test2' -ExpectedBeforeValue 'test' -ExpectedKind 'LiteralChanged' -RequireNoErrors
    } 'probe/Acceptance Macro.xlsm-sequence-2.json'
    Invoke-Step 'xlsm save 3 clear while Excel remains open' {
        Save-ExcelLiteral -Address $xlsmCell -Value '' -Clear
        Invoke-Probe -Workbook $xlsm -Sequence 3 -Address $xlsmCell -ExpectedValue '' -ExpectedBeforeValue 'test2' -ExpectedKind 'LiteralCleared' -ExpectCleared -RequireNoErrors
    } 'probe/Acceptance Macro.xlsm-sequence-3.json'
    $macroHashAfter = Get-ZipEntrySha256 $xlsm 'xl/vbaProject.bin'
    Add-Assertion 'xlsm VBA part unchanged' ($macroHashBefore -eq $macroHashAfter) 'macro-hashes.json' "$macroHashBefore -> $macroHashAfter"
    Write-AcceptanceUtf8File -Path (Join-Path $evidence 'macro-hashes.json') -Content (@{ before = $macroHashBefore; after = $macroHashAfter } | ConvertTo-Json)

    Click-ProductControl 'Excel Diff Tracker' 'HistoryNavigationButton'
    $main = Get-ProductWindow 'Excel Diff Tracker'
    Save-UiState 'history-after-real-excel-saves' $main

    Invoke-Step 'installed app tray, relaunch, startup, and repair lifecycle' {
        Invoke-LifecycleGate
        $main = Get-ProductWindow 'Excel Diff Tracker'
    } 'lifecycle/lifecycle.json'

    Invoke-Step 'installed real Excel semantic matrix' {
        $semanticEvidence = Join-Path $evidence 'semantic-matrix'
        $null = & $semanticMatrixRunner `
            -InstallerPath $installer `
            -ExpectedInstallerSha256 $installerHash `
            -ExpectedApplicationSha256 $installedApplicationHash `
            -ProbePath $probe `
            -EvidenceDirectory $semanticEvidence `
            -ConfirmInstalledCandidate
        $semanticResult = Join-Path $semanticEvidence 'installed-semantic-matrix.json'
        $null = & $semanticMatrixValidator `
            -ResultPath $semanticResult `
            -InstallerPath $installer `
            -ExpectedApplicationSha256 $installedApplicationHash `
            -ProbePath $probe
    } 'semantic-matrix/installed-semantic-matrix.json'

    Invoke-Step 'installed-product 500,000-cell benchmark' {
        $largeEvidence = Join-Path $evidence 'large-workbook'
        $null = & $largeBenchmarkRunner `
            -InstallerPath $installer `
            -ExpectedInstallerSha256 $installerHash `
            -ExpectedApplicationSha256 $installedApplicationHash `
            -ProbePath $probe `
            -EvidenceDirectory $largeEvidence `
            -ConfirmInstalledCandidate
        $largeResult = Join-Path $largeEvidence 'large-workbook-benchmark.json'
        $null = & $largeBenchmarkValidator `
            -ResultPath $largeResult `
            -InstallerPath $installer `
            -ExpectedApplicationSha256 $installedApplicationHash
    } 'large-workbook/large-workbook-benchmark.json'

    Invoke-Step 'twenty-save real Excel soak over ten minutes' {
        $soakEvidence = Join-Path $evidence 'soak'
        $null = & $soakRunner `
            -InstallerPath $installer `
            -ExpectedInstallerSha256 $installerHash `
            -ExpectedApplicationSha256 $installedApplicationHash `
            -ProbePath $probe `
            -XlsxFixture $sourceXlsx `
            -XlsmFixture $sourceXlsm `
            -EvidenceDirectory $soakEvidence `
            -ConfirmInstalledCandidate
        $soakResult = Join-Path $soakEvidence 'real-excel-soak.json'
        $null = & $soakValidator `
            -ResultPath $soakResult `
            -InstallerPath $installer `
            -ExpectedApplicationSha256 $installedApplicationHash
    } 'soak/real-excel-soak.json'
}
catch {
    $failed = $true
    Add-Assertion 'unhandled acceptance step' $false 'logs/acceptance-transcript.txt' $_.Exception.ToString()
    try { Save-DesktopScreenshot -Path (Join-Path $screenshots 'failure.png') } catch { }
}
finally {
    try { if ($xlsmWorkbook) { $xlsmWorkbook.Close($false) } } catch { }
    try { if ($xlsxWorkbook) { $xlsxWorkbook.Close($false) } } catch { }
    try { if ($excel) { $excel.Quit() } } catch { }
    foreach ($value in @($xlsmWorkbook, $xlsxWorkbook, $excel)) {
        if ($value) { try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value) } catch { } }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    if ($appProcess -and -not $appProcess.HasExited) {
        try { Stop-Process -Id $appProcess.Id -Force -ErrorAction Stop } catch { }
    }
    Start-Sleep -Milliseconds 500
    if (Test-Path $databasePath) {
        Copy-Item $databasePath (Join-Path $evidence 'history.db') -Force
    }
    if (Test-Path $reportDirectory) {
        Copy-Item $reportDirectory (Join-Path $evidence 'reports') -Recurse -Force
    }

    try {
        $windowsErrors = Get-WinEvent -FilterHashtable @{ LogName = 'Application'; StartTime = $runStartedUtc } -ErrorAction Stop |
            Where-Object ProviderName -in @('Application Error', '.NET Runtime', 'Windows Error Reporting') |
            Select-Object TimeCreated, ProviderName, Id, LevelDisplayName, Message
        $windowsErrorsJson = ConvertTo-Json -InputObject @($windowsErrors) -Depth 5
        Write-AcceptanceUtf8File -Path (Join-Path $logs 'windows-errors.json') -Content $windowsErrorsJson
        Add-Assertion 'no Application Error, .NET Runtime, or WER events' (@($windowsErrors).Count -eq 0) 'logs/windows-errors.json' "events=$(@($windowsErrors).Count)"
    } catch {
        Add-Assertion 'Windows application error log exported' $false 'logs/windows-errors.json' $_.Exception.Message
    }

    if (Test-Path (Join-Path $installDirectory 'unins000.exe')) {
        $uninstall = Start-Process (Join-Path $installDirectory 'unins000.exe') -ArgumentList '/VERYSILENT','/SUPPRESSMSGBOXES','/NORESTART' -Wait -PassThru
        Add-Assertion 'silent uninstall exits successfully' ($uninstall.ExitCode -eq 0) 'acceptance.json' "Exit code $($uninstall.ExitCode)"
        Add-Assertion 'uninstall removes installed binaries' (-not (Test-Path $installDirectory)) 'acceptance.json' $installDirectory
        Add-Assertion 'uninstall removes Start-menu shortcut' (-not (Test-Path $startMenuShortcut)) 'acceptance.json' $startMenuShortcut
        $startupAfterUninstall = (Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run' -Name ExcelDiffTracker -ErrorAction SilentlyContinue).ExcelDiffTracker
        Add-Assertion 'uninstall removes startup registration' ([string]::IsNullOrWhiteSpace($startupAfterUninstall)) 'acceptance.json' "value=$startupAfterUninstall"
        Add-Assertion 'uninstall retains local history by policy' ((Test-Path $dataDirectory -PathType Container) -and (Test-Path $databasePath -PathType Leaf)) 'acceptance.json' $dataDirectory
    } else {
        Add-Assertion 'uninstaller exists' $false 'acceptance.json' (Join-Path $installDirectory 'unins000.exe')
    }

    $summary = [ordered]@{
        status = if (-not $failed -and @($assertions | Where-Object { -not $_.passed }).Count -eq 0) { 'Passed' } else { 'Failed' }
        version = $Version
        runId = $runId
        runNumber = $RunNumber
        startedUtc = $runStartedUtc.ToString('O')
        finishedUtc = [DateTime]::UtcNow.ToString('O')
        sourceCommit = $sourceCommit
        vmSnapshotName = $VmSnapshotName
        vmSnapshotId = $VmSnapshotId
        installerSha256 = Get-AcceptanceFileSha256 -Path $installer
        installedApplicationSha256 = $installedApplicationHash
        assertions = $assertions
    }
    Write-AcceptanceUtf8File -Path (Join-Path $evidence 'acceptance.json') -Content ($summary | ConvertTo-Json -Depth 10)

    try { Stop-Transcript | Out-Null } catch { }

    $checksumPath = Join-Path $evidence 'SHA256SUMS.txt'
    $checksumLines = Get-ChildItem $evidence -File -Recurse |
        Where-Object FullName -ne $checksumPath |
        Sort-Object FullName |
        ForEach-Object {
            $relative = Get-AcceptanceRelativePath -BasePath $evidence -Path $_.FullName -UseForwardSlash
            "$((Get-AcceptanceFileSha256 -Path $_.FullName).ToLowerInvariant())  $relative"
    }
    Write-AcceptanceUtf8File -Path $checksumPath -Content (($checksumLines -join [Environment]::NewLine) + [Environment]::NewLine)
}

if ($failed -or @($assertions | Where-Object { -not $_.passed }).Count -ne 0) {
    throw "BLACK_BOX_ACCEPTANCE_FAILED|$evidence"
}
Write-Output "BLACK_BOX_ACCEPTANCE_PASS|$evidence"
