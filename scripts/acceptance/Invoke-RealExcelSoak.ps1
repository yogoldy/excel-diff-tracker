[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string] $InstallerPath,
    [Parameter(Mandatory)] [ValidatePattern('^[A-Fa-f0-9]{64}$')] [string] $ExpectedInstallerSha256,
    [Parameter(Mandatory)] [ValidatePattern('^[A-Fa-f0-9]{64}$')] [string] $ExpectedApplicationSha256,
    [Parameter(Mandatory)] [string] $ProbePath,
    [Parameter(Mandatory)] [string] $XlsxFixture,
    [Parameter(Mandatory)] [string] $XlsmFixture,
    [Parameter(Mandatory)] [string] $EvidenceDirectory,
    [Parameter(Mandatory)] [ValidatePattern('^[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}$')] [string] $OuterRunEvidenceId,
    [string] $ApplicationPath = (Join-Path $env:LOCALAPPDATA 'Programs\Excel Diff Tracker\ExcelDiffTracker.exe'),
    [string] $DatabasePath = (Join-Path $env:LOCALAPPDATA 'Excel Diff Tracker\history.db'),
    [ValidateSet(20)] [int] $SaveCount = 20,
    [ValidateRange(30, 300)] [int] $SaveIntervalSeconds = 32,
    [switch] $ConfirmInstalledCandidate
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

if (-not $ConfirmInstalledCandidate) {
    throw 'The real-Excel soak is fail-closed. Rerun against an installed candidate with -ConfirmInstalledCandidate.'
}
if (($SaveCount % 2) -ne 0) {
    throw 'SaveCount must be even so both workbooks receive the same number of saves.'
}
if ((($SaveCount - 1) * $SaveIntervalSeconds) -lt 600) {
    throw 'The first-to-last save schedule must span at least ten minutes.'
}

$installer = (Resolve-Path $InstallerPath).Path
$application = (Resolve-Path $ApplicationPath).Path
$probe = (Resolve-Path $ProbePath).Path
$sourceXlsx = (Resolve-Path $XlsxFixture).Path
$sourceXlsm = (Resolve-Path $XlsmFixture).Path
$database = [System.IO.Path]::GetFullPath($DatabasePath)
$evidence = [System.IO.Path]::GetFullPath($EvidenceDirectory)
$fixtures = Join-Path $evidence 'fixtures'
$probeResults = Join-Path $evidence 'probe'
$reports = Join-Path $evidence 'reports'
$screenshots = Join-Path $evidence 'screenshots'
$uia = Join-Path $evidence 'uia'
$xlsx = Join-Path $fixtures 'Soak.xlsx'
$xlsm = Join-Path $fixtures 'Soak Macro.xlsm'
$resultPath = Join-Path $evidence 'real-excel-soak.json'
$cell = 'Y1001'

if (Test-Path $evidence) { throw "Soak evidence directory already exists; use a fresh path: $evidence" }
New-Item -ItemType Directory -Path $evidence, $fixtures, $probeResults, $reports, $screenshots, $uia -Force | Out-Null
Copy-Item $sourceXlsx $xlsx -Force
Copy-Item $sourceXlsm $xlsm -Force

Import-Module (Join-Path $PSScriptRoot 'UiAutomation.psm1') -Force

$startedUtc = [DateTime]::UtcNow
$evidenceId = [Guid]::NewGuid().ToString('D')
$assertions = [System.Collections.Generic.List[object]]::new()
$saves = [System.Collections.Generic.List[object]]::new()
$failed = $false
$failure = $null
$excel = $null
$xlsxWorkbook = $null
$xlsmWorkbook = $null
$applicationProcess = $null
$mainWindow = $null
$macroHashBefore = $null
$macroHashAfter = $null
$scheduleClock = $null

function Add-SoakAssertion {
    param([string] $Name, [bool] $Passed, [string] $Detail = '', [string] $Evidence = '')
    $script:assertions.Add([pscustomobject]@{
        name = $Name
        passed = $Passed
        detail = $Detail
        evidence = $Evidence
        utc = [DateTime]::UtcNow.ToString('O')
    })
    if (-not $Passed) { $script:failed = $true }
}

function Assert-Soak {
    param([string] $Name, [bool] $Condition, [string] $Detail = '', [string] $Evidence = '')
    Add-SoakAssertion -Name $Name -Passed $Condition -Detail $Detail -Evidence $Evidence
    if (-not $Condition) { throw "$Name failed. $Detail" }
}

function Get-ZipEntrySha256 {
    param([string] $Path, [string] $EntryName)
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $entry = $archive.GetEntry($EntryName)
        if (-not $entry) { throw "Archive entry not found: $EntryName" }
        $stream = $entry.Open()
        $algorithm = [System.Security.Cryptography.SHA256]::Create()
        try {
            [Convert]::ToBase64String($algorithm.ComputeHash($stream))
        }
        finally {
            $algorithm.Dispose()
            $stream.Dispose()
        }
    }
    finally { $archive.Dispose() }
}

function Choose-FileFromDialog {
    param([string] $Path)
    $dialog = Find-UiaWindow -Title 'Choose a workbook to track' -TimeoutSeconds 15
    $fileName = Find-UiaElement -Root $dialog -AutomationId '1148' -Optional
    if (-not $fileName) { $fileName = Find-UiaElement -Root $dialog -Name 'File name:' }
    Set-UiaValue -Element $fileName -Value $Path
    $open = Find-UiaElement -Root $dialog -AutomationId '1' -Optional
    if (-not $open) { $open = Find-UiaElement -Root $dialog -Name 'Open' }
    Invoke-UiaElement -Element $open
}

function Add-TrackedWorkbook {
    param([string] $Path)
    Set-UiaForeground -Window $mainWindow
    $dashboard = Find-UiaElement -Root $mainWindow -AutomationId 'DashboardNavigationButton'
    Invoke-UiaElement -Element $dashboard
    Start-Sleep -Milliseconds 300
    $script:mainWindow = Find-UiaWindow -Title 'Excel Diff Tracker' -TimeoutSeconds 10
    $add = Find-UiaElement -Root $mainWindow -AutomationId 'DashboardAddWorkbookButton'
    Invoke-UiaElement -Element $add
    Choose-FileFromDialog -Path $Path
}

function Invoke-ProbeUntilPassed {
    param(
        [string] $WorkbookPath,
        [long] $ExpectedSequence,
        [string] $ExpectedValue,
        [string] $ExpectedKind,
        [string] $ExpectedBeforeValue,
        [string] $OutputName,
        [int] $TimeoutSeconds = 20,
        [switch] $Baseline
    )
    $outputPath = Join-Path $probeResults $OutputName
    $deadline = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
    $lastOutput = ''
    do {
        if ($applicationProcess.HasExited) { throw 'Excel Diff Tracker exited during the real-Excel soak.' }
        $applicationProcess.Refresh()
        if (-not $applicationProcess.Responding) { throw 'Excel Diff Tracker reported Not Responding during the real-Excel soak.' }

        $arguments = @(
            '--database', $database,
            '--workbook', $WorkbookPath,
            '--expected-sequence', $ExpectedSequence,
            '--require-active',
            '--require-no-errors',
            '--require-no-last-error',
            '--expected-version-count', $ExpectedSequence,
            '--require-unique-version-hashes')
        if (-not $Baseline) {
            $arguments += @(
                '--address', $cell,
                '--expected-value', $ExpectedValue,
                '--expected-kind', $ExpectedKind,
                '--expected-cell-change-count', '1',
                '--expected-sheet-change-count', '0',
                '--require-source-hash-match',
                '--require-ready-report',
                '--report-contains', $cell)
            if ($ExpectedSequence -eq 1) { $arguments += '--expect-before-missing' }
            else { $arguments += @('--expected-before-value', $ExpectedBeforeValue) }
        }
        $lastOutput = & $probe @arguments 2>&1 | Out-String
        $exitCode = $LASTEXITCODE
        if ($exitCode -eq 0) {
            [System.IO.File]::WriteAllText($outputPath, $lastOutput, [System.Text.UTF8Encoding]::new($false))
            return $lastOutput | ConvertFrom-Json
        }
        Start-Sleep -Milliseconds 300
    } while ([DateTime]::UtcNow -lt $deadline)
    [System.IO.File]::WriteAllText($outputPath, $lastOutput, [System.Text.UTF8Encoding]::new($false))
    throw "Soak probe did not pass within $TimeoutSeconds seconds. See $outputPath"
}

function Activate-AndSaveLiteral {
    param([object] $Workbook, [string] $Value)
    [void]$Workbook.Activate()
    $excelWindow = Get-UiaWindowFromHandle -Handle ([long]$excel.Hwnd)
    Set-UiaForeground -Window $excelWindow
    [System.Windows.Forms.SendKeys]::SendWait('^g')
    Start-Sleep -Milliseconds 200
    [System.Windows.Forms.SendKeys]::SendWait($cell)
    [System.Windows.Forms.SendKeys]::SendWait('{ENTER}')
    Start-Sleep -Milliseconds 200
    [System.Windows.Forms.SendKeys]::SendWait($Value)
    [System.Windows.Forms.SendKeys]::SendWait('{ENTER}')
    [System.Windows.Forms.SendKeys]::SendWait('^s')
    [System.Diagnostics.Stopwatch]::StartNew()
}

function Wait-UntilScheduledSave {
    param([System.Diagnostics.Stopwatch] $Clock, [double] $TargetElapsedSeconds)
    while ($Clock.Elapsed.TotalSeconds -lt $TargetElapsedSeconds) {
        if ($applicationProcess.HasExited) { throw 'Excel Diff Tracker exited between soak saves.' }
        $applicationProcess.Refresh()
        if (-not $applicationProcess.Responding) { throw 'Excel Diff Tracker reported Not Responding between soak saves.' }
        $remaining = ($TargetElapsedSeconds - $Clock.Elapsed.TotalSeconds) * 1000
        Start-Sleep -Milliseconds ([int][math]::Min(500, [math]::Max(20, $remaining)))
    }
}

try {
    $installerHash = (Get-FileHash $installer -Algorithm SHA256).Hash.ToUpperInvariant()
    $applicationHash = (Get-FileHash $application -Algorithm SHA256).Hash.ToUpperInvariant()
    Assert-Soak 'soak uses the frozen installer' ($installerHash -eq $ExpectedInstallerSha256.ToUpperInvariant()) $installerHash
    Assert-Soak 'soak uses the frozen installed executable' ($applicationHash -eq $ExpectedApplicationSha256.ToUpperInvariant()) $applicationHash
    Assert-Soak 'installed app database exists' (Test-Path $database -PathType Leaf) $database

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3
    $xlsxWorkbook = $excel.Workbooks.Open($xlsx)
    $xlsxWorkbook.Worksheets.Item(1).Range($cell).ClearContents()
    $xlsxWorkbook.Save()
    $xlsmWorkbook = $excel.Workbooks.Open($xlsm)
    $xlsmWorkbook.Worksheets.Item(1).Range($cell).ClearContents()
    $xlsmWorkbook.Save()
    $macroHashBefore = Get-ZipEntrySha256 $xlsm 'xl/vbaProject.bin'

    Start-Process $application | Out-Null
    $mainWindow = Find-UiaWindow -Title 'Excel Diff Tracker' -TimeoutSeconds 20
    $applicationProcess = Get-Process -Id $mainWindow.Current.ProcessId -ErrorAction Stop
    Assert-Soak 'soak controls the installed executable' ([string]::Equals($applicationProcess.Path, $application, [StringComparison]::OrdinalIgnoreCase)) $applicationProcess.Path

    Add-TrackedWorkbook $xlsx
    $xlsxBaseline = Invoke-ProbeUntilPassed -WorkbookPath $xlsx -ExpectedSequence 0 -OutputName 'xlsx-baseline.json' -TimeoutSeconds 60 -Baseline
    Add-TrackedWorkbook $xlsm
    $xlsmBaseline = Invoke-ProbeUntilPassed -WorkbookPath $xlsm -ExpectedSequence 0 -OutputName 'xlsm-baseline.json' -TimeoutSeconds 60 -Baseline
    Assert-Soak 'both soak baselines are silent sequence zero' ($xlsxBaseline.currentSequence -eq 0 -and $xlsmBaseline.currentSequence -eq 0) 'xlsx=0 xlsm=0'

    $scheduleStartedUtc = [DateTime]::UtcNow
    $scheduleClock = [System.Diagnostics.Stopwatch]::StartNew()
    $xlsxSequence = 0L
    $xlsmSequence = 0L
    for ($index = 1; $index -le $SaveCount; $index++) {
        $targetElapsedSeconds = [double](($index - 1) * $SaveIntervalSeconds)
        $targetUtc = $scheduleStartedUtc.AddSeconds($targetElapsedSeconds)
        Wait-UntilScheduledSave -Clock $scheduleClock -TargetElapsedSeconds $targetElapsedSeconds
        $isXlsx = ($index % 2) -eq 1
        $workbookPath = if ($isXlsx) { $xlsx } else { $xlsm }
        $workbook = if ($isXlsx) { $xlsxWorkbook } else { $xlsmWorkbook }
        if ($isXlsx) { $xlsxSequence++ } else { $xlsmSequence++ }
        $sequence = if ($isXlsx) { $xlsxSequence } else { $xlsmSequence }
        $format = if ($isXlsx) { 'xlsx' } else { 'xlsm' }
        $value = 'EDT-SOAK-{0:D2}' -f $index
        $previousValue = if ($sequence -gt 1) { 'EDT-SOAK-{0:D2}' -f ($index - 2) } else { $null }
        $kind = if ($sequence -eq 1) { 'LiteralAdded' } else { 'LiteralChanged' }
        $saveStarted = [DateTime]::UtcNow
        $monotonicStartSeconds = $scheduleClock.Elapsed.TotalSeconds
        $captureClock = Activate-AndSaveLiteral -Workbook $workbook -Value $value
        $ctrlSaveUtc = [DateTime]::UtcNow
        $probeName = '{0}-sequence-{1:D2}.json' -f $format, $sequence
        $probeResult = Invoke-ProbeUntilPassed -WorkbookPath $workbookPath -ExpectedSequence $sequence -ExpectedValue $value -ExpectedKind $kind -ExpectedBeforeValue $previousValue -OutputName $probeName
        $captureClock.Stop()
        $saveFinished = [DateTime]::UtcNow
        $fileHash = (Get-FileHash $workbookPath -Algorithm SHA256).Hash.ToUpperInvariant()
        Assert-Soak "soak save $index captured the exact stable hash" ($probeResult.currentHash.ToUpperInvariant() -eq $fileHash -and $probeResult.latestVersion.sha256.ToUpperInvariant() -eq $fileHash) $fileHash "probe/$probeName"
        $reportName = '{0}-sequence-{1:D2}.md' -f $format, $sequence
        $reportCopy = Join-Path $reports $reportName
        Copy-Item $probeResult.latestVersion.reportPath $reportCopy -Force
        $reportText = [System.IO.File]::ReadAllText($reportCopy)
        Assert-Soak "soak save $index Markdown contains the exact address and value" (
            $reportText.IndexOf($cell, [StringComparison]::Ordinal) -ge 0 -and
            $reportText.IndexOf($value, [StringComparison]::Ordinal) -ge 0) $reportName "reports/$reportName"
        $saves.Add([pscustomobject]@{
            index = $index
            format = $format
            workbook = $workbookPath
            sequence = $sequence
            value = $value
            expectedKind = $kind
            scheduledUtc = $targetUtc.ToString('O')
            saveStartedUtc = $saveStarted.ToString('O')
            capturedUtc = $saveFinished.ToString('O')
            monotonicStartSeconds = [math]::Round($monotonicStartSeconds, 3)
            ctrlSaveUtc = $ctrlSaveUtc.ToString('O')
            captureMilliseconds = [math]::Round($captureClock.Elapsed.TotalMilliseconds, 3)
            sha256 = $fileHash
            probe = "probe/$probeName"
            report = "reports/$reportName"
        })
    }

    Start-Sleep -Seconds 12
    $perWorkbookSaves = [long]($SaveCount / 2)
    $xlsxSettled = Invoke-ProbeUntilPassed -WorkbookPath $xlsx -ExpectedSequence $perWorkbookSaves -ExpectedValue ('EDT-SOAK-{0:D2}' -f ($SaveCount - 1)) -ExpectedKind 'LiteralChanged' -ExpectedBeforeValue ('EDT-SOAK-{0:D2}' -f ($SaveCount - 3)) -OutputName 'xlsx-settled.json'
    $xlsmSettled = Invoke-ProbeUntilPassed -WorkbookPath $xlsm -ExpectedSequence $perWorkbookSaves -ExpectedValue ('EDT-SOAK-{0:D2}' -f $SaveCount) -ExpectedKind 'LiteralChanged' -ExpectedBeforeValue ('EDT-SOAK-{0:D2}' -f ($SaveCount - 2)) -OutputName 'xlsm-settled.json'
    $macroHashAfter = Get-ZipEntrySha256 $xlsm 'xl/vbaProject.bin'

    $scheduleClock.Stop()
    $durationSeconds = $scheduleClock.Elapsed.TotalSeconds
    $xlsxHashes = @($saves | Where-Object format -eq 'xlsx' | Select-Object -ExpandProperty sha256 -Unique)
    $xlsmHashes = @($saves | Where-Object format -eq 'xlsm' | Select-Object -ExpandProperty sha256 -Unique)
    Assert-Soak 'twenty real saves span at least ten minutes' ($durationSeconds -ge 600) "$durationSeconds seconds"
    Assert-Soak 'xlsx has ten distinct captured stable hashes' ($xlsxHashes.Count -eq $perWorkbookSaves) "distinct=$($xlsxHashes.Count)"
    Assert-Soak 'xlsm has ten distinct captured stable hashes' ($xlsmHashes.Count -eq $perWorkbookSaves) "distinct=$($xlsmHashes.Count)"
    Assert-Soak 'both workbooks settle Active with exact version counts' (
        $xlsxSettled.workbookStatus -eq 'Active' -and
        $xlsmSettled.workbookStatus -eq 'Active' -and
        $xlsxSettled.versionCount -eq $perWorkbookSaves -and
        $xlsmSettled.versionCount -eq $perWorkbookSaves) "xlsx=$($xlsxSettled.versionCount) xlsm=$($xlsmSettled.versionCount)"
    Assert-Soak 'xlsm VBA project remains byte-identical' ($macroHashBefore -eq $macroHashAfter) "$macroHashBefore -> $macroHashAfter"

    Set-UiaForeground -Window $mainWindow
    $history = Find-UiaElement -Root $mainWindow -AutomationId 'HistoryNavigationButton'
    Invoke-UiaElement -Element $history
    Start-Sleep -Milliseconds 500
    Save-DesktopScreenshot -Path (Join-Path $screenshots 'soak-history.png')
    Export-UiaTree -Root $mainWindow -Path (Join-Path $uia 'soak-history.json')
}
catch {
    $failed = $true
    $failure = $_.Exception.ToString()
    Add-SoakAssertion -Name 'unhandled real-Excel soak step' -Passed $false -Detail $failure -Evidence 'real-excel-soak.json'
    try { Save-DesktopScreenshot -Path (Join-Path $screenshots 'soak-failure.png') } catch { }
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

    $finishedUtc = [DateTime]::UtcNow
    $result = [ordered]@{
        schemaVersion = 2
        evidenceId = $evidenceId
        outerRunEvidenceId = $OuterRunEvidenceId.ToLowerInvariant()
        gate = 'real-excel-ten-minute-soak'
        status = if (-not $failed -and $saves.Count -eq $SaveCount -and @($assertions | Where-Object { -not $_.passed }).Count -eq 0) { 'Passed' } else { 'Failed' }
        startedUtc = $startedUtc.ToString('O')
        finishedUtc = $finishedUtc.ToString('O')
        durationSeconds = [math]::Round(($finishedUtc - $startedUtc).TotalSeconds, 3)
        monotonicDurationSeconds = if ($null -ne $scheduleClock) { [math]::Round($scheduleClock.Elapsed.TotalSeconds, 3) } else { 0 }
        saveCount = $SaveCount
        saveIntervalSeconds = $SaveIntervalSeconds
        candidate = [ordered]@{
            installerSha256 = if (Test-Path $installer) { (Get-FileHash $installer -Algorithm SHA256).Hash.ToUpperInvariant() } else { $null }
            expectedInstallerSha256 = $ExpectedInstallerSha256.ToUpperInvariant()
            applicationSha256 = if (Test-Path $application) { (Get-FileHash $application -Algorithm SHA256).Hash.ToUpperInvariant() } else { $null }
            expectedApplicationSha256 = $ExpectedApplicationSha256.ToUpperInvariant()
        }
        workbooks = @(
            [ordered]@{ format = 'xlsx'; path = $xlsx; finalSequence = $xlsxSequence; macroHashBefore = $null; macroHashAfter = $null },
            [ordered]@{ format = 'xlsm'; path = $xlsm; finalSequence = $xlsmSequence; macroHashBefore = $macroHashBefore; macroHashAfter = $macroHashAfter }
        )
        saves = $saves
        assertions = $assertions
        failure = $failure
    }
    [System.IO.File]::WriteAllText($resultPath, ($result | ConvertTo-Json -Depth 12), [System.Text.UTF8Encoding]::new($false))
}

if ($failed -or $saves.Count -ne $SaveCount -or $result.status -ne 'Passed' -or @($assertions | Where-Object { -not $_.passed }).Count -ne 0) {
    throw "REAL_EXCEL_SOAK_FAILED|$resultPath"
}
Write-Output "REAL_EXCEL_SOAK_PASS|$resultPath"
