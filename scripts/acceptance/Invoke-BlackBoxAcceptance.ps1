[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string] $InstallerPath,
    [Parameter(Mandatory)] [string] $ProbePath,
    [Parameter(Mandatory)] [string] $XlsxFixture,
    [Parameter(Mandatory)] [string] $XlsmFixture,
    [ValidatePattern('^\d+\.\d+\.\d+$')] [string] $Version = '0.1.2',
    [ValidateRange(1, 2)] [int] $RunNumber = 1,
    [string] $EvidenceRoot,
    [string] $ExpectedInstallerSha256,
    [switch] $ConfirmCleanSnapshot,
    [switch] $KeepInstalled
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
New-Item -ItemType Directory -Path $evidence, $screenshots, $uia, $logs, $fixtures, $probeResults, $recoveryResults -Force | Out-Null
Start-Transcript -Path (Join-Path $logs 'acceptance-transcript.txt') -Force | Out-Null

$assertions = [System.Collections.Generic.List[object]]::new()
$failed = $false
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
    Find-UiaWindow -Title $Title -TimeoutSeconds 20
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
    [System.Windows.Forms.SendKeys]::SendWait('^s')
    Start-Sleep -Milliseconds 500
}

function Invoke-Probe {
    param(
        [string] $Workbook,
        [long] $Sequence,
        [string] $Address,
        [string] $ExpectedValue,
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
            '--require-active',
            '--require-no-last-error',
            '--require-ready-report',
            '--report-contains', $Address)
        if ($RequireNoErrors) { $arguments += '--require-no-errors' }
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

function Write-EnvironmentManifest {
    $excelPath = (Get-Command excel.exe -ErrorAction SilentlyContinue).Source
    $manifest = [ordered]@{
        runId = $runId
        runNumber = $RunNumber
        startedUtc = $runStartedUtc.ToString('O')
        computerName = $env:COMPUTERNAME
        user = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
        isAdministrator = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        os = Get-CimInstance Win32_OperatingSystem | Select-Object Caption, Version, BuildNumber, OSArchitecture
        computer = Get-CimInstance Win32_ComputerSystem | Select-Object Manufacturer, Model, NumberOfLogicalProcessors, TotalPhysicalMemory
        video = Get-CimInstance Win32_VideoController | Select-Object Name, CurrentHorizontalResolution, CurrentVerticalResolution
        dotnetOnPath = [bool](Get-Command dotnet -ErrorAction SilentlyContinue)
        excelPath = $excelPath
        installerPath = $installer
        installerSha256 = Get-AcceptanceFileSha256 -Path $installer
        sourceCommit = (& git -C $repositoryRoot rev-parse HEAD 2>$null)
    }
    Write-AcceptanceUtf8File -Path (Join-Path $evidence 'environment.json') -Content ($manifest | ConvertTo-Json -Depth 8)
}

try {
    Write-EnvironmentManifest
    $installerHash = Get-AcceptanceFileSha256 -Path $installer
    if ($ExpectedInstallerSha256 -and $installerHash -ne $ExpectedInstallerSha256.ToUpperInvariant()) {
        throw "Installer SHA-256 $installerHash does not match expected $ExpectedInstallerSha256."
    }
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

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3
    $xlsxWorkbook = $excel.Workbooks.Open($xlsx)
    $xlsxWorkbook.Worksheets.Item(1).Range($xlsxCell).ClearContents()
    $xlsxWorkbook.Save()
    Prepare-RecoveryFixtures

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
        Invoke-Probe -Workbook $xlsx -Sequence 2 -Address $xlsxCell -ExpectedValue 'test2' -ExpectedKind 'LiteralChanged' -RequireNoErrors
    } 'probe/Acceptance.xlsx-sequence-2.json'
    Invoke-Step 'xlsx save 3 clear while Excel remains open' {
        Save-ExcelLiteral -Address $xlsxCell -Value '' -Clear
        Invoke-Probe -Workbook $xlsx -Sequence 3 -Address $xlsxCell -ExpectedValue '' -ExpectedKind 'LiteralCleared' -ExpectCleared -RequireNoErrors
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
        Invoke-Probe -Workbook $xlsm -Sequence 2 -Address $xlsmCell -ExpectedValue 'test2' -ExpectedKind 'LiteralChanged' -RequireNoErrors
    } 'probe/Acceptance Macro.xlsm-sequence-2.json'
    Invoke-Step 'xlsm save 3 clear while Excel remains open' {
        Save-ExcelLiteral -Address $xlsmCell -Value '' -Clear
        Invoke-Probe -Workbook $xlsm -Sequence 3 -Address $xlsmCell -ExpectedValue '' -ExpectedKind 'LiteralCleared' -ExpectCleared -RequireNoErrors
    } 'probe/Acceptance Macro.xlsm-sequence-3.json'
    $macroHashAfter = Get-ZipEntrySha256 $xlsm 'xl/vbaProject.bin'
    Add-Assertion 'xlsm VBA part unchanged' ($macroHashBefore -eq $macroHashAfter) 'macro-hashes.json' "$macroHashBefore -> $macroHashAfter"
    Write-AcceptanceUtf8File -Path (Join-Path $evidence 'macro-hashes.json') -Content (@{ before = $macroHashBefore; after = $macroHashAfter } | ConvertTo-Json)

    Click-ProductControl 'Excel Diff Tracker' 'HistoryNavigationButton'
    $main = Get-ProductWindow 'Excel Diff Tracker'
    Save-UiState 'history-after-real-excel-saves' $main
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
    } catch { }

    if (-not $KeepInstalled -and (Test-Path (Join-Path $installDirectory 'unins000.exe'))) {
        $uninstall = Start-Process (Join-Path $installDirectory 'unins000.exe') -ArgumentList '/VERYSILENT','/SUPPRESSMSGBOXES','/NORESTART' -Wait -PassThru
        Add-Assertion 'silent uninstall' ($uninstall.ExitCode -eq 0 -and -not (Test-Path $installDirectory)) 'acceptance.json' "Exit code $($uninstall.ExitCode)"
    }

    $summary = [ordered]@{
        status = if (-not $failed -and @($assertions | Where-Object { -not $_.passed }).Count -eq 0) { 'Passed' } else { 'Failed' }
        version = $Version
        runId = $runId
        runNumber = $RunNumber
        startedUtc = $runStartedUtc.ToString('O')
        finishedUtc = [DateTime]::UtcNow.ToString('O')
        installerSha256 = Get-AcceptanceFileSha256 -Path $installer
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
