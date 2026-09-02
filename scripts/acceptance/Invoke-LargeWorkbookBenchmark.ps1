[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string] $InstallerPath,
    [Parameter(Mandatory)] [ValidatePattern('^[A-Fa-f0-9]{64}$')] [string] $ExpectedInstallerSha256,
    [Parameter(Mandatory)] [ValidatePattern('^[A-Fa-f0-9]{64}$')] [string] $ExpectedApplicationSha256,
    [Parameter(Mandatory)] [string] $ProbePath,
    [Parameter(Mandatory)] [string] $EvidenceDirectory,
    [Parameter(Mandatory)] [ValidatePattern('^[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}$')] [string] $OuterRunEvidenceId,
    [string] $ApplicationPath = (Join-Path $env:LOCALAPPDATA 'Programs\Excel Diff Tracker\ExcelDiffTracker.exe'),
    [string] $DatabasePath = (Join-Path $env:LOCALAPPDATA 'Excel Diff Tracker\history.db'),
    [switch] $ConfirmInstalledCandidate
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$phaseTimeoutSeconds = 60
$maxUiaHeartbeatMilliseconds = 2000
$maxHeartbeatGapMilliseconds = 2000
$maxPrivateBytes = 1610612736L # 1.5 GiB
$rows = 25000
$columns = 20
$populatedCells = $rows * $columns
$bulkChangedCells = 10000
$automaticReportCellLimit = 5000

if (-not $ConfirmInstalledCandidate) {
    throw 'The large-workbook gate is fail-closed. Rerun against an installed candidate with -ConfirmInstalledCandidate.'
}

$repositoryRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$generator = Join-Path $repositoryRoot 'scripts\create-large-excel-fixture.ps1'
$installer = (Resolve-Path $InstallerPath).Path
$application = (Resolve-Path $ApplicationPath).Path
$probe = (Resolve-Path $ProbePath).Path
$database = [System.IO.Path]::GetFullPath($DatabasePath)
$evidence = [System.IO.Path]::GetFullPath($EvidenceDirectory)
$fixtureDirectory = Join-Path $evidence 'fixtures'
$probeDirectory = Join-Path $evidence 'probe'
$reportDirectory = Join-Path $evidence 'reports'
$screenshotDirectory = Join-Path $evidence 'screenshots'
$uiaDirectory = Join-Path $evidence 'uia'
$telemetryDirectory = Join-Path $evidence 'telemetry'
$logDirectory = Join-Path $evidence 'logs'
$fixture = Join-Path $fixtureDirectory 'Acceptance Large 500k.xlsx'
$resultPath = Join-Path $evidence 'large-workbook-benchmark.json'
$telemetryPath = Join-Path $telemetryDirectory 'large-workbook-telemetry.jsonl'
$monitorStopPath = Join-Path $telemetryDirectory 'large-workbook-monitor.stop'
$monitorLogPath = Join-Path $logDirectory 'large-workbook-monitor.txt'
$fullExportPath = Join-Path $reportDirectory 'large-workbook-version-000004-full.md'
$startedUtc = [DateTime]::UtcNow
$evidenceId = [Guid]::NewGuid().ToString('D')

if (Test-Path $evidence) { throw "Benchmark evidence directory already exists; use a fresh path: $evidence" }
New-Item -ItemType Directory -Path $evidence, $fixtureDirectory, $probeDirectory, $reportDirectory, $screenshotDirectory, $uiaDirectory, $telemetryDirectory, $logDirectory -Force | Out-Null
if (-not (Test-Path $generator -PathType Leaf)) { throw "Large fixture generator is missing: $generator" }

Import-Module (Join-Path $PSScriptRoot 'UiAutomation.psm1') -Force

$assertions = [System.Collections.Generic.List[object]]::new()
$phases = [System.Collections.Generic.List[object]]::new()
$failed = $false
$failureMessage = $null
$excel = $null
$workbook = $null
$monitorJob = $null
$mainWindow = $null
$applicationProcess = $null
$generationElapsedMilliseconds = $null
$initialFixtureSha256 = $null
$initialFixtureBytes = $null
$usedRangeRows = $null
$usedRangeColumns = $null
$usedRangeCells = $null
$monitorForcedStop = $false
$monitorStartedUtc = $null
$monitorStoppedUtc = $null
$monitorJobState = $null

function Add-GateAssertion {
    param(
        [Parameter(Mandatory)] [string] $Name,
        [Parameter(Mandatory)] [bool] $Passed,
        [string] $EvidencePath = '',
        [string] $Detail = ''
    )
    $script:assertions.Add([pscustomobject]@{
        name = $Name
        passed = $Passed
        evidence = $EvidencePath
        detail = $Detail
        utc = [DateTime]::UtcNow.ToString('O')
    })
    if (-not $Passed) { $script:failed = $true }
}

function Assert-Gate {
    param(
        [Parameter(Mandatory)] [string] $Name,
        [Parameter(Mandatory)] [bool] $Condition,
        [string] $EvidencePath = '',
        [string] $Detail = ''
    )
    Add-GateAssertion -Name $Name -Passed $Condition -EvidencePath $EvidencePath -Detail $Detail
    if (-not $Condition) { throw "$Name failed. $Detail" }
}

function Get-RelativeEvidencePath {
    param([Parameter(Mandatory)] [string] $Path)
    Get-AcceptanceRelativePath -BasePath $evidence -Path $Path -UseForwardSlash
}

function Choose-FileFromDialog {
    param([Parameter(Mandatory)] [string] $Path)
    $dialog = Find-UiaWindow -Title 'Choose a workbook to track' -TimeoutSeconds 15
    $fileName = Find-UiaElement -Root $dialog -AutomationId '1148' -Optional
    if (-not $fileName) { $fileName = Find-UiaElement -Root $dialog -Name 'File name:' }
    Set-UiaValue -Element $fileName -Value $Path
    $open = Find-UiaElement -Root $dialog -AutomationId '1' -Optional
    if (-not $open) { $open = Find-UiaElement -Root $dialog -Name 'Open' }
    Invoke-UiaElement -Element $open
}

function Save-UiEvidence {
    param(
        [Parameter(Mandatory)] [string] $Name,
        [Parameter(Mandatory)] [System.Windows.Automation.AutomationElement] $Window
    )
    $screenshot = Join-Path $screenshotDirectory "$Name.png"
    $tree = Join-Path $uiaDirectory "$Name.json"
    Save-DesktopScreenshot -Path $screenshot
    Export-UiaTree -Root $Window -Path $tree
    [pscustomobject]@{
        screenshot = Get-RelativeEvidencePath $screenshot
        uiaTree = Get-RelativeEvidencePath $tree
    }
}

function Test-HistoryUi {
    param(
        [Parameter(Mandatory)] [long] $Sequence,
        [Parameter(Mandatory)] [string] $Summary,
        [Parameter(Mandatory)] [string] $EvidenceName
    )
    Set-UiaForeground -Window $mainWindow
    $navigation = Find-UiaElement -Root $mainWindow -AutomationId 'HistoryNavigationButton'
    Invoke-UiaElement -Element $navigation
    $names = @()
    $deadline = [DateTime]::UtcNow.AddSeconds(10)
    do {
        $names = @($mainWindow.FindAll(
            [System.Windows.Automation.TreeScope]::Descendants,
            [System.Windows.Automation.Condition]::TrueCondition) | ForEach-Object { $_.Current.Name })
        if (($names -contains $Summary) -and @($names | Where-Object { $_ -match "^Version $Sequence(?:\s|$)" }).Count -gt 0 -and ($names -contains $fixture)) { break }
        Start-Sleep -Milliseconds 200
    } while ([DateTime]::UtcNow -lt $deadline)
    $summaryVisible = $names -contains $Summary
    $versionVisible = @($names | Where-Object { $_ -match "^Version $Sequence(?:\s|$)" }).Count -gt 0
    $workbookVisible = $names -contains $fixture
    $uiEvidence = Save-UiEvidence -Name $EvidenceName -Window $mainWindow
    Assert-Gate -Name "history UI shows sequence $Sequence" -Condition $versionVisible -EvidencePath $uiEvidence.uiaTree -Detail "Expected a visible Version $Sequence history card."
    Assert-Gate -Name "history UI shows summary for sequence $Sequence" -Condition $summaryVisible -EvidencePath $uiEvidence.uiaTree -Detail "Expected summary '$Summary'."
    Assert-Gate -Name "history UI identifies benchmark workbook at sequence $Sequence" -Condition $workbookVisible -EvidencePath $uiEvidence.uiaTree -Detail $fixture
    $dashboard = Find-UiaElement -Root $mainWindow -AutomationId 'DashboardNavigationButton'
    Invoke-UiaElement -Element $dashboard
    $uiEvidence
}

function Test-BaselineDashboardUi {
    Set-UiaForeground -Window $mainWindow
    $dashboard = Find-UiaElement -Root $mainWindow -AutomationId 'DashboardNavigationButton'
    Invoke-UiaElement -Element $dashboard
    $names = @()
    $deadline = [DateTime]::UtcNow.AddSeconds(10)
    do {
        $names = @($mainWindow.FindAll(
            [System.Windows.Automation.TreeScope]::Descendants,
            [System.Windows.Automation.Condition]::TrueCondition) | ForEach-Object { $_.Current.Name })
        if ($names -contains $fixture) { break }
        Start-Sleep -Milliseconds 200
    } while ([DateTime]::UtcNow -lt $deadline)
    $uiEvidence = Save-UiEvidence -Name 'dashboard-large-baseline' -Window $mainWindow
    Assert-Gate -Name 'dashboard shows the 500,000-cell baseline workbook' -Condition ($names -contains $fixture) -EvidencePath $uiEvidence.uiaTree -Detail $fixture
    $uiEvidence
}

function Invoke-ProbeUntilPassed {
    param(
        [Parameter(Mandatory)] [long] $Sequence,
        [string] $Address,
        [string] $ExpectedValue,
        [string] $ExpectedKind,
        [Parameter(Mandatory)] [string] $OutputName,
        [Parameter(Mandatory)] [System.Diagnostics.Stopwatch] $Stopwatch
    )
    $outputPath = Join-Path $probeDirectory $OutputName
    $lastOutput = ''
    do {
        if ($applicationProcess.HasExited) { throw 'Excel Diff Tracker exited during the large-workbook benchmark.' }
        $applicationProcess.Refresh()
        if (-not $applicationProcess.Responding) { throw 'Excel Diff Tracker reported Not Responding during the large-workbook benchmark.' }

        $arguments = @(
            '--database', $database,
            '--workbook', $fixture,
            '--expected-sequence', $Sequence,
            '--require-active',
            '--require-no-errors',
            '--require-no-last-error')
        if ($Sequence -gt 0) {
            $arguments += @(
                '--address', $Address,
                '--expected-value', $ExpectedValue,
                '--expected-kind', $ExpectedKind,
                '--require-ready-report',
                '--report-contains', $Address)
        }

        $lastOutput = & $probe @arguments 2>&1 | Out-String
        $exitCode = $LASTEXITCODE
        if ($exitCode -eq 0) {
            [System.IO.File]::WriteAllText($outputPath, $lastOutput, [System.Text.UTF8Encoding]::new($false))
            return $lastOutput | ConvertFrom-Json
        }
        Start-Sleep -Milliseconds 300
    } while ($Stopwatch.Elapsed.TotalSeconds -lt $phaseTimeoutSeconds)

    [System.IO.File]::WriteAllText($outputPath, $lastOutput, [System.Text.UTF8Encoding]::new($false))
    throw "Sequence $Sequence did not become Ready within $phaseTimeoutSeconds seconds. See $outputPath"
}

function Copy-AndVerifyAutomaticReport {
    param(
        [Parameter(Mandatory)] [object] $ProbeResult,
        [Parameter(Mandatory)] [long] $Sequence,
        [Parameter(Mandatory)] [string] $Address,
        [Parameter(Mandatory)] [string] $ExpectedValue,
        [Parameter(Mandatory)] [long] $ExpectedCellChanges,
        [Parameter(Mandatory)] [long] $ExpectedHeadingCount,
        [switch] $RequireTruncation
    )
    $source = [string]$ProbeResult.latestVersion.reportPath
    Assert-Gate -Name "sequence $Sequence automatic report exists" -Condition (Test-Path $source -PathType Leaf) -EvidencePath $source -Detail $source
    $copy = Join-Path $reportDirectory ("sequence-{0:D6}-automatic.md" -f $Sequence)
    Copy-Item $source $copy -Force
    $markdown = [System.IO.File]::ReadAllText($copy)
    $escapedAddress = [regex]::Escape("### Performance!$Address")
    $addressCount = [regex]::Matches($markdown, "(?m)^$escapedAddress\r?$").Count
    $valueCount = [regex]::Matches($markdown, [regex]::Escape($ExpectedValue)).Count
    $headingCount = [regex]::Matches($markdown, '(?m)^### Performance!').Count
    $truncationText = "> This automatic report includes the first $($automaticReportCellLimit.ToString('N0')) of $($ExpectedCellChanges.ToString('N0')) cell changes."
    $truncated = $markdown.IndexOf($truncationText, [StringComparison]::Ordinal) -ge 0

    Assert-Gate -Name "sequence $Sequence Markdown contains the exact changed address once" -Condition ($addressCount -eq 1) -EvidencePath (Get-RelativeEvidencePath $copy) -Detail "addressCount=$addressCount"
    Assert-Gate -Name "sequence $Sequence Markdown contains each included new value exactly once" -Condition ($valueCount -eq $ExpectedHeadingCount) -EvidencePath (Get-RelativeEvidencePath $copy) -Detail "expected=$ExpectedHeadingCount valueCount=$valueCount"
    Assert-Gate -Name "sequence $Sequence automatic report has the expected cell detail count" -Condition ($headingCount -eq $ExpectedHeadingCount) -EvidencePath (Get-RelativeEvidencePath $copy) -Detail "expected=$ExpectedHeadingCount actual=$headingCount"
    Assert-Gate -Name "sequence $Sequence automatic report truncation policy" -Condition ($truncated -eq [bool]$RequireTruncation) -EvidencePath (Get-RelativeEvidencePath $copy) -Detail "expected=$([bool]$RequireTruncation) actual=$truncated"
    [pscustomobject]@{
        path = Get-RelativeEvidencePath $copy
        addressOccurrences = $addressCount
        valueOccurrences = $valueCount
        cellHeadingCount = $headingCount
        truncated = $truncated
    }
}

function Invoke-KeyboardSave {
    param(
        [Parameter(Mandatory)] [string] $AddressOrRange,
        [Parameter(Mandatory)] [string] $Value,
        [switch] $FillSelection
    )
    $excelWindow = Get-UiaWindowFromHandle -Handle ([long]$excel.Hwnd)
    Set-UiaForeground -Window $excelWindow
    [System.Windows.Forms.SendKeys]::SendWait('^g')
    Start-Sleep -Milliseconds 200
    [System.Windows.Forms.SendKeys]::SendWait($AddressOrRange)
    [System.Windows.Forms.SendKeys]::SendWait('{ENTER}')
    Start-Sleep -Milliseconds 200
    [System.Windows.Forms.SendKeys]::SendWait($Value)
    if ($FillSelection) { [System.Windows.Forms.SendKeys]::SendWait('^{ENTER}') }
    else { [System.Windows.Forms.SendKeys]::SendWait('{ENTER}') }
    [System.Windows.Forms.SendKeys]::SendWait('^s')
}

function Invoke-CapturePhase {
    param(
        [Parameter(Mandatory)] [string] $Name,
        [Parameter(Mandatory)] [long] $Sequence,
        [Parameter(Mandatory)] [string] $AddressOrRange,
        [Parameter(Mandatory)] [string] $EvidenceAddress,
        [Parameter(Mandatory)] [string] $Value,
        [Parameter(Mandatory)] [long] $ExpectedCellChanges,
        [Parameter(Mandatory)] [long] $ExpectedAutomaticHeadings,
        [switch] $FillSelection,
        [switch] $RequireTruncation
    )
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    Invoke-KeyboardSave -AddressOrRange $AddressOrRange -Value $Value -FillSelection:$FillSelection
    $probeResult = Invoke-ProbeUntilPassed -Sequence $Sequence -Address $EvidenceAddress -ExpectedValue $Value -ExpectedKind 'FormulaRemoved' -OutputName ("sequence-{0:D6}.json" -f $Sequence) -Stopwatch $stopwatch

    Assert-Gate -Name "$Name has exactly $ExpectedCellChanges SQLite cell changes" -Condition ([long]$probeResult.latestVersion.cellChangeCount -eq $ExpectedCellChanges) -EvidencePath ("probe/sequence-{0:D6}.json" -f $Sequence) -Detail "actual=$($probeResult.latestVersion.cellChangeCount)"
    Assert-Gate -Name "$Name has no sheet changes" -Condition ([long]$probeResult.latestVersion.sheetChangeCount -eq 0) -EvidencePath ("probe/sequence-{0:D6}.json" -f $Sequence) -Detail "actual=$($probeResult.latestVersion.sheetChangeCount)"
    Assert-Gate -Name "$Name delta is on Performance!$EvidenceAddress" -Condition ($probeResult.cellChange.sheetName -eq 'Performance' -and $probeResult.cellChange.address -eq $EvidenceAddress) -EvidencePath ("probe/sequence-{0:D6}.json" -f $Sequence) -Detail "$($probeResult.cellChange.sheetName)!$($probeResult.cellChange.address)"

    $fileHash = Get-AcceptanceFileSha256 -Path $fixture
    Assert-Gate -Name "$Name captured hash matches the saved workbook" -Condition ($probeResult.currentHash.ToUpperInvariant() -eq $fileHash -and $probeResult.latestVersion.sha256.ToUpperInvariant() -eq $fileHash) -EvidencePath ("probe/sequence-{0:D6}.json" -f $Sequence) -Detail $fileHash
    $automaticReport = Copy-AndVerifyAutomaticReport -ProbeResult $probeResult -Sequence $Sequence -Address $EvidenceAddress -ExpectedValue $Value -ExpectedCellChanges $ExpectedCellChanges -ExpectedHeadingCount $ExpectedAutomaticHeadings -RequireTruncation:$RequireTruncation
    $uiEvidence = Test-HistoryUi -Sequence $Sequence -Summary "$($ExpectedCellChanges.ToString('N0')) cells, 0 sheet changes" -EvidenceName ("history-sequence-{0:D6}" -f $Sequence)
    $stopwatch.Stop()
    Assert-Gate -Name "$Name completes in SQLite, Markdown, and UI within 60 seconds" -Condition ($stopwatch.Elapsed.TotalSeconds -le $phaseTimeoutSeconds) -EvidencePath ("probe/sequence-{0:D6}.json" -f $Sequence) -Detail "$($stopwatch.Elapsed.TotalMilliseconds) ms"

    $phase = [pscustomobject]@{
        name = $Name
        sequence = $Sequence
        addressOrRange = $AddressOrRange
        evidenceAddress = $EvidenceAddress
        expectedValue = $Value
        expectedKind = 'FormulaRemoved'
        expectedCellChanges = $ExpectedCellChanges
        elapsedMilliseconds = [math]::Round($stopwatch.Elapsed.TotalMilliseconds, 3)
        timeoutSeconds = $phaseTimeoutSeconds
        workbookSha256 = $fileHash
        probe = "probe/sequence-{0:D6}.json" -f $Sequence
        automaticReport = $automaticReport
        uiEvidence = $uiEvidence
    }
    $script:phases.Add($phase)
    $phase
}

function Export-LatestVersionThroughUi {
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    Set-UiaForeground -Window $mainWindow
    $history = Find-UiaElement -Root $mainWindow -AutomationId 'HistoryNavigationButton'
    Invoke-UiaElement -Element $history
    Start-Sleep -Milliseconds 300
    $buttons = $mainWindow.FindAll(
        [System.Windows.Automation.TreeScope]::Descendants,
        [System.Windows.Automation.PropertyCondition]::new(
            [System.Windows.Automation.AutomationElement]::NameProperty,
            'Full export'))
    Assert-Gate -Name 'history exposes Full export for the 10,000-cell version' -Condition ($buttons.Count -gt 0) -EvidencePath 'uia/history-sequence-000004.json' -Detail "buttons=$($buttons.Count)"
    Invoke-UiaElement -Element $buttons.Item(0)

    $dialog = Find-UiaWindow -Title 'Export complete Markdown report' -TimeoutSeconds 15
    $dialogEvidence = Save-UiEvidence -Name 'full-export-dialog-sequence-000004' -Window $dialog
    $fileName = Find-UiaElement -Root $dialog -AutomationId '1001' -Optional
    if (-not $fileName) { $fileName = Find-UiaElement -Root $dialog -AutomationId '1148' -Optional }
    if (-not $fileName) { $fileName = Find-UiaElement -Root $dialog -Name 'File name:' }
    $valuePattern = $null
    $defaultName = if ($fileName.TryGetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern, [ref]$valuePattern)) {
        ([System.Windows.Automation.ValuePattern]$valuePattern).Current.Value
    } else { '' }
    Assert-Gate -Name 'Full export targets the latest sequence' -Condition ($defaultName -match 'version-000004-full\.md$') -EvidencePath 'uia/history-sequence-000004.json' -Detail "defaultName=$defaultName"
    Set-UiaValue -Element $fileName -Value $fullExportPath
    $save = Find-UiaElement -Root $dialog -AutomationId '1' -Optional
    if (-not $save) { $save = Find-UiaElement -Root $dialog -Name 'Save' }
    Invoke-UiaElement -Element $save
    Wait-AcceptanceCondition -TimeoutSeconds 60 -FailureMessage 'Full Markdown export was not written within 60 seconds.' -Condition { Test-Path $fullExportPath -PathType Leaf }

    $markdown = [System.IO.File]::ReadAllText($fullExportPath)
    $headingCount = [regex]::Matches($markdown, '(?m)^### Performance!').Count
    $valueCount = [regex]::Matches($markdown, 'EDT-LARGE-BULK').Count
    $truncated = $markdown.IndexOf('> This automatic report includes the first ', [StringComparison]::Ordinal) -ge 0
    Assert-Gate -Name 'Full export contains all 10,000 changed cells' -Condition ($headingCount -eq $bulkChangedCells) -EvidencePath (Get-RelativeEvidencePath $fullExportPath) -Detail "headings=$headingCount"
    Assert-Gate -Name 'Full export contains every bulk value' -Condition ($valueCount -eq $bulkChangedCells) -EvidencePath (Get-RelativeEvidencePath $fullExportPath) -Detail "values=$valueCount"
    Assert-Gate -Name 'Full export is not truncated' -Condition (-not $truncated) -EvidencePath (Get-RelativeEvidencePath $fullExportPath) -Detail "truncated=$truncated"
    $stopwatch.Stop()
    Assert-Gate -Name 'Full export completes within 60 seconds' -Condition ($stopwatch.Elapsed.TotalSeconds -le $phaseTimeoutSeconds) -EvidencePath (Get-RelativeEvidencePath $fullExportPath) -Detail "$($stopwatch.Elapsed.TotalMilliseconds) ms"
    [pscustomobject]@{
        path = Get-RelativeEvidencePath $fullExportPath
        exportedViaInstalledUi = $true
        cellHeadingCount = $headingCount
        valueOccurrences = $valueCount
        truncated = $truncated
        elapsedMilliseconds = [math]::Round($stopwatch.Elapsed.TotalMilliseconds, 3)
        defaultFileName = $defaultName
        uiEvidence = $dialogEvidence
    }
}

try {
    $installerHash = Get-AcceptanceFileSha256 -Path $installer
    $applicationHash = Get-AcceptanceFileSha256 -Path $application
    Assert-Gate -Name 'benchmark uses the frozen installer' -Condition ($installerHash -eq $ExpectedInstallerSha256.ToUpperInvariant()) -EvidencePath '' -Detail $installerHash
    Assert-Gate -Name 'benchmark uses the frozen installed executable' -Condition ($applicationHash -eq $ExpectedApplicationSha256.ToUpperInvariant()) -EvidencePath '' -Detail $applicationHash
    Assert-Gate -Name 'installed app database exists' -Condition (Test-Path $database -PathType Leaf) -EvidencePath '' -Detail $database

    $generationStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $generatorOutput = & $generator -OutputPath $fixture -Rows $rows -Columns $columns 2>&1 | Out-String
    $generationStopwatch.Stop()
    $generationElapsedMilliseconds = [math]::Round($generationStopwatch.Elapsed.TotalMilliseconds, 3)
    [System.IO.File]::WriteAllText((Join-Path $logDirectory 'large-fixture-generation.txt'), $generatorOutput, [System.Text.UTF8Encoding]::new($false))
    Assert-Gate -Name 'fixture generator reports exactly 500,000 cells' -Condition ($generatorOutput -match 'LARGE_FIXTURE_CREATED\|rows=25000\|columns=20\|cells=500000\|') -EvidencePath 'logs/large-fixture-generation.txt' -Detail $generatorOutput.Trim()
    $initialFixtureSha256 = Get-AcceptanceFileSha256 -Path $fixture
    $initialFixtureBytes = (Get-Item $fixture).Length

    Start-Process $application | Out-Null
    $mainWindow = Find-UiaWindow -Title 'Excel Diff Tracker' -TimeoutSeconds 20
    $welcome = Find-UiaElement -Root ([System.Windows.Automation.AutomationElement]::RootElement) -Name 'Welcome to Excel Diff Tracker' -Optional
    Assert-Gate -Name 'installed candidate is already onboarded' -Condition ($null -eq $welcome) -EvidencePath '' -Detail 'The benchmark must run after real onboarding.'
    $applicationProcess = Get-Process -Id $mainWindow.Current.ProcessId -ErrorAction Stop
    Assert-Gate -Name 'UI belongs to the installed executable' -Condition ($applicationProcess.Path -eq $application) -EvidencePath '' -Detail $applicationProcess.Path

    if (Test-Path $monitorStopPath) { Remove-Item $monitorStopPath -Force }
    $monitorJob = Start-Job -ArgumentList (Join-Path $PSScriptRoot 'UiAutomation.psm1'), $applicationProcess.Id, ([long]$mainWindow.Current.NativeWindowHandle), $telemetryPath, $monitorStopPath -ScriptBlock {
        param($UiModule, $ProcessId, $WindowHandle, $TelemetryPath, $StopPath)
        Set-StrictMode -Version Latest
        $ErrorActionPreference = 'Stop'
        Import-Module $UiModule -Force
        $previousStart = $null
        $lastPrivateWorkingSet = $null
        $lastPrivateWorkingSetUtc = $null
        $sampleIndex = 0
        while (-not (Test-Path $StopPath)) {
            $queryStarted = [DateTime]::UtcNow
            $queryStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
            $uiaOk = $false
            $processFound = $false
            $responding = $false
            $privateBytes = $null
            $workingSetBytes = $null
            $peakWorkingSetBytes = $null
            $errorText = $null
            try {
                $process = Get-Process -Id $ProcessId -ErrorAction Stop
                $process.Refresh()
                $processFound = $true
                $responding = $process.Responding
                $privateBytes = [long]$process.PrivateMemorySize64
                $workingSetBytes = [long]$process.WorkingSet64
                $peakWorkingSetBytes = [long]$process.PeakWorkingSet64
                if ($sampleIndex % 4 -eq 0) {
                    $performance = Get-CimInstance Win32_PerfFormattedData_PerfProc_Process -Filter "IDProcess = $ProcessId" -ErrorAction Stop
                    $lastPrivateWorkingSet = [long]$performance.WorkingSetPrivate
                    $lastPrivateWorkingSetUtc = [DateTime]::UtcNow
                }
                $window = Get-UiaWindowFromHandle -Handle $WindowHandle
                if ($window.Current.Name -ne 'Excel Diff Tracker') { throw "Unexpected window title: $($window.Current.Name)" }
                $heartbeat = Find-UiaElement -Root $window -AutomationId 'DashboardNavigationButton'
                $null = $heartbeat.Current.IsEnabled
                $uiaOk = $true
            }
            catch {
                $errorText = $_.Exception.Message
            }
            $queryStopwatch.Stop()
            $privateWorkingSetAge = if ($null -eq $lastPrivateWorkingSetUtc) { $null } else { ([DateTime]::UtcNow - $lastPrivateWorkingSetUtc).TotalMilliseconds }
            $gap = if ($null -eq $previousStart) { 0.0 } else { ($queryStarted - $previousStart).TotalMilliseconds }
            $previousStart = $queryStarted
            $sample = [ordered]@{
                sample = $sampleIndex
                startedUtc = $queryStarted.ToString('O')
                durationMilliseconds = [math]::Round($queryStopwatch.Elapsed.TotalMilliseconds, 3)
                gapMilliseconds = [math]::Round($gap, 3)
                processFound = $processFound
                responding = $responding
                uiaOk = $uiaOk
                privateBytes = $privateBytes
                workingSetBytes = $workingSetBytes
                peakWorkingSetBytes = $peakWorkingSetBytes
                privateWorkingSetBytes = $lastPrivateWorkingSet
                privateWorkingSetAgeMilliseconds = if ($null -eq $privateWorkingSetAge) { $null } else { [math]::Round($privateWorkingSetAge, 3) }
                error = $errorText
            }
            [System.IO.File]::AppendAllText($TelemetryPath, (($sample | ConvertTo-Json -Compress) + [Environment]::NewLine), [System.Text.UTF8Encoding]::new($false))
            $sampleIndex++
            Start-Sleep -Milliseconds 250
        }
    }
    Wait-AcceptanceCondition -TimeoutSeconds 10 -FailureMessage 'The independent UI/memory monitor did not produce telemetry.' -Condition {
        (Test-Path $telemetryPath -PathType Leaf) -and (Get-Item $telemetryPath).Length -gt 0
    }
    $monitorStartedUtc = [DateTime]::UtcNow

    $baselineStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $dashboardNavigation = Find-UiaElement -Root $mainWindow -AutomationId 'DashboardNavigationButton'
    Invoke-UiaElement -Element $dashboardNavigation
    Start-Sleep -Milliseconds 300
    $mainWindow = Find-UiaWindow -Title 'Excel Diff Tracker' -TimeoutSeconds 10
    $addButton = Find-UiaElement -Root $mainWindow -AutomationId 'DashboardAddWorkbookButton'
    Invoke-UiaElement -Element $addButton
    Choose-FileFromDialog -Path $fixture
    $baselineProbe = Invoke-ProbeUntilPassed -Sequence 0 -OutputName 'sequence-000000-baseline.json' -Stopwatch $baselineStopwatch
    Assert-Gate -Name 'baseline is silent sequence zero' -Condition ([long]$baselineProbe.currentSequence -eq 0 -and $null -eq $baselineProbe.latestVersion) -EvidencePath 'probe/sequence-000000-baseline.json' -Detail "sequence=$($baselineProbe.currentSequence)"
    Assert-Gate -Name 'baseline hash matches generated fixture' -Condition ($baselineProbe.currentHash.ToUpperInvariant() -eq $initialFixtureSha256) -EvidencePath 'probe/sequence-000000-baseline.json' -Detail $baselineProbe.currentHash
    $baselineUiEvidence = Test-BaselineDashboardUi
    $baselineStopwatch.Stop()
    Assert-Gate -Name '500,000-cell baseline completes in SQLite and UI within 60 seconds' -Condition ($baselineStopwatch.Elapsed.TotalSeconds -le $phaseTimeoutSeconds) -EvidencePath 'probe/sequence-000000-baseline.json' -Detail "$($baselineStopwatch.Elapsed.TotalMilliseconds) ms"
    $phases.Add([pscustomobject]@{
        name = 'baseline'
        sequence = 0
        expectedCellChanges = 0
        elapsedMilliseconds = [math]::Round($baselineStopwatch.Elapsed.TotalMilliseconds, 3)
        timeoutSeconds = $phaseTimeoutSeconds
        workbookSha256 = $initialFixtureSha256
        probe = 'probe/sequence-000000-baseline.json'
        uiEvidence = $baselineUiEvidence
    })

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3
    $workbook = $excel.Workbooks.Open($fixture)
    $usedRange = $workbook.Worksheets.Item(1).UsedRange
    $usedRangeRows = [long]$usedRange.Rows.Count
    $usedRangeColumns = [long]$usedRange.Columns.Count
    $usedRangeCells = [long]$usedRange.Cells.CountLarge
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($usedRange)
    Assert-Gate -Name 'visible Excel opened the full 500,000-cell fixture' -Condition ($usedRangeRows -eq $rows -and $usedRangeColumns -eq $columns -and $usedRangeCells -eq $populatedCells) -EvidencePath 'fixtures/Acceptance Large 500k.xlsx' -Detail "rows=$usedRangeRows columns=$usedRangeColumns cells=$usedRangeCells"

    $smallPhases = @(
        @{ Name = 'beginning-cell save'; Sequence = 1L; Address = 'A1'; Value = 'EDT-LARGE-BEGIN' },
        @{ Name = 'middle-cell save'; Sequence = 2L; Address = 'J12500'; Value = 'EDT-LARGE-MIDDLE' },
        @{ Name = 'end-cell save'; Sequence = 3L; Address = 'T25000'; Value = 'EDT-LARGE-END' }
    )
    foreach ($phase in $smallPhases) {
        Invoke-CapturePhase -Name $phase.Name -Sequence $phase.Sequence -AddressOrRange $phase.Address -EvidenceAddress $phase.Address -Value $phase.Value -ExpectedCellChanges 1 -ExpectedAutomaticHeadings 1 | Out-Null
    }

    $bulkPhase = Invoke-CapturePhase -Name '10,000-cell truncation save' -Sequence 4 -AddressOrRange 'A2:T501' -EvidenceAddress 'A2' -Value 'EDT-LARGE-BULK' -ExpectedCellChanges $bulkChangedCells -ExpectedAutomaticHeadings $automaticReportCellLimit -FillSelection -RequireTruncation
    $fullExport = Export-LatestVersionThroughUi
    $bulkPhase | Add-Member -NotePropertyName fullExport -NotePropertyValue $fullExport

    $distinctHashes = @($phases | Select-Object -ExpandProperty workbookSha256 -Unique)
    Assert-Gate -Name 'baseline and four real saves have five distinct stable hashes' -Condition ($distinctHashes.Count -eq 5) -EvidencePath 'large-workbook-benchmark.json' -Detail "distinctHashes=$($distinctHashes.Count)"
}
catch {
    $failed = $true
    $failureMessage = $_.Exception.ToString()
    Add-GateAssertion -Name 'unhandled large-workbook benchmark step' -Passed $false -EvidencePath 'large-workbook-benchmark.json' -Detail $failureMessage
    try { Save-DesktopScreenshot -Path (Join-Path $screenshotDirectory 'large-workbook-failure.png') } catch { }
}
finally {
    try { if ($workbook) { $workbook.Close($false) } } catch { }
    try { if ($excel) { $excel.Quit() } } catch { }
    foreach ($value in @($workbook, $excel)) {
        if ($value) { try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value) } catch { } }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    if ($monitorJob) {
        $monitorStoppedUtc = [DateTime]::UtcNow
        try { [System.IO.File]::WriteAllText($monitorStopPath, 'stop', [System.Text.UTF8Encoding]::new($false)) } catch { }
        try { $null = Wait-Job -Job $monitorJob -Timeout 15 } catch { }
        if ($monitorJob.State -notin @('Completed', 'Failed', 'Stopped')) {
            $monitorForcedStop = $true
            Stop-Job -Job $monitorJob -ErrorAction SilentlyContinue
        }
        $monitorJobState = $monitorJob.State.ToString()
        try {
            $monitorOutput = Receive-Job -Job $monitorJob -ErrorAction SilentlyContinue | Out-String
            [System.IO.File]::WriteAllText($monitorLogPath, $monitorOutput, [System.Text.UTF8Encoding]::new($false))
        } catch { }
        Remove-Job -Job $monitorJob -Force -ErrorAction SilentlyContinue
    }
    Remove-Item $monitorStopPath -Force -ErrorAction SilentlyContinue

    $telemetry = if (Test-Path $telemetryPath -PathType Leaf) {
        @(Get-Content $telemetryPath | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_ | ConvertFrom-Json })
    } else { @() }
    $sampleCount = $telemetry.Count
    $telemetryDurationMilliseconds = if ($sampleCount -gt 1) {
        ([DateTime]::Parse($telemetry[-1].startedUtc).ToUniversalTime() - [DateTime]::Parse($telemetry[0].startedUtc).ToUniversalTime()).TotalMilliseconds
    } else { 0.0 }
    $requiredMonitorDurationMilliseconds = if ($null -ne $monitorStartedUtc -and $null -ne $monitorStoppedUtc) {
        ($monitorStoppedUtc - $monitorStartedUtc).TotalMilliseconds
    } else { 0.0 }
    $initialTelemetryLagMilliseconds = if ($sampleCount -gt 0 -and $null -ne $monitorStartedUtc) {
        [math]::Max(0, ([DateTime]::Parse($telemetry[0].startedUtc).ToUniversalTime() - $monitorStartedUtc).TotalMilliseconds)
    } else { [double]($maxHeartbeatGapMilliseconds + 1) }
    $tailTelemetryLagMilliseconds = if ($sampleCount -gt 0 -and $null -ne $monitorStoppedUtc) {
        ($monitorStoppedUtc - [DateTime]::Parse($telemetry[-1].startedUtc).ToUniversalTime()).TotalMilliseconds
    } else { [double]($maxHeartbeatGapMilliseconds + 1) }
    $maxUiaDuration = if ($sampleCount -gt 0) { ($telemetry | Measure-Object durationMilliseconds -Maximum).Maximum } else { [double]($maxUiaHeartbeatMilliseconds + 1) }
    $maxGap = if ($sampleCount -gt 1) { ($telemetry | Select-Object -Skip 1 | Measure-Object gapMilliseconds -Maximum).Maximum } else { [double]($maxHeartbeatGapMilliseconds + 1) }
    $peakPrivateBytes = if ($sampleCount -gt 0) { ($telemetry | Measure-Object privateBytes -Maximum).Maximum } else { $maxPrivateBytes }
    $peakPrivateWorkingSetBytes = if ($sampleCount -gt 0) { ($telemetry | Where-Object { $null -ne $_.privateWorkingSetBytes } | Measure-Object privateWorkingSetBytes -Maximum).Maximum } else { $null }
    $peakWorkingSetBytes = if ($sampleCount -gt 0) { ($telemetry | Measure-Object peakWorkingSetBytes -Maximum).Maximum } else { $maxPrivateBytes }
    $expectedMinimumSamples = [math]::Max(2, [math]::Floor($requiredMonitorDurationMilliseconds / $maxHeartbeatGapMilliseconds))
    $telemetryFailures = @($telemetry | Where-Object {
        -not $_.processFound -or -not $_.responding -or -not $_.uiaOk -or
        $null -eq $_.privateWorkingSetBytes -or
        [double]$_.privateWorkingSetAgeMilliseconds -gt $maxHeartbeatGapMilliseconds -or
        [double]$_.durationMilliseconds -gt $maxUiaHeartbeatMilliseconds -or
        [double]$_.gapMilliseconds -gt $maxHeartbeatGapMilliseconds
    })

    Add-GateAssertion -Name 'UI/memory monitor stopped cleanly' -Passed (-not $monitorForcedStop -and $monitorJobState -eq 'Completed') -EvidencePath 'logs/large-workbook-monitor.txt' -Detail "forcedStop=$monitorForcedStop state=$monitorJobState"
    Add-GateAssertion -Name 'UI heartbeat telemetry covers the benchmark' -Passed ($sampleCount -ge $expectedMinimumSamples -and $initialTelemetryLagMilliseconds -le $maxHeartbeatGapMilliseconds -and $tailTelemetryLagMilliseconds -le $maxHeartbeatGapMilliseconds) -EvidencePath 'telemetry/large-workbook-telemetry.jsonl' -Detail "samples=$sampleCount expectedMinimum=$expectedMinimumSamples initialLagMs=$initialTelemetryLagMilliseconds tailLagMs=$tailTelemetryLagMilliseconds requiredDurationMs=$requiredMonitorDurationMilliseconds"
    Add-GateAssertion -Name 'WPF remained responsive for every heartbeat' -Passed ($telemetryFailures.Count -eq 0) -EvidencePath 'telemetry/large-workbook-telemetry.jsonl' -Detail "failedSamples=$($telemetryFailures.Count)"
    Add-GateAssertion -Name 'UI Automation heartbeat has no query over two seconds' -Passed ([double]$maxUiaDuration -le $maxUiaHeartbeatMilliseconds) -EvidencePath 'telemetry/large-workbook-telemetry.jsonl' -Detail "maxMs=$maxUiaDuration"
    Add-GateAssertion -Name 'UI Automation heartbeat has no gap over two seconds' -Passed ([double]$maxGap -le $maxHeartbeatGapMilliseconds) -EvidencePath 'telemetry/large-workbook-telemetry.jsonl' -Detail "maxMs=$maxGap"
    Add-GateAssertion -Name 'peak private bytes stay below 1.5 GiB' -Passed ([long]$peakPrivateBytes -lt $maxPrivateBytes) -EvidencePath 'telemetry/large-workbook-telemetry.jsonl' -Detail "peak=$peakPrivateBytes threshold=$maxPrivateBytes"
    Add-GateAssertion -Name 'peak private working set stays below 1.5 GiB' -Passed ($null -ne $peakPrivateWorkingSetBytes -and [long]$peakPrivateWorkingSetBytes -lt $maxPrivateBytes) -EvidencePath 'telemetry/large-workbook-telemetry.jsonl' -Detail "peak=$peakPrivateWorkingSetBytes threshold=$maxPrivateBytes"
    Add-GateAssertion -Name 'peak total working set stays below 1.5 GiB' -Passed ([long]$peakWorkingSetBytes -lt $maxPrivateBytes) -EvidencePath 'telemetry/large-workbook-telemetry.jsonl' -Detail "peak=$peakWorkingSetBytes threshold=$maxPrivateBytes"

    $installerHashForResult = if (Test-Path $installer -PathType Leaf) { Get-AcceptanceFileSha256 -Path $installer } else { $null }
    $applicationHashForResult = if (Test-Path $application -PathType Leaf) { Get-AcceptanceFileSha256 -Path $application } else { $null }
    $status = if (-not $failed -and @($assertions | Where-Object { -not $_.passed }).Count -eq 0 -and $phases.Count -eq 5) { 'Passed' } else { 'Failed' }
    $result = [ordered]@{
        schemaVersion = 2
        evidenceId = $evidenceId
        outerRunEvidenceId = $OuterRunEvidenceId.ToLowerInvariant()
        gate = 'large-workbook-500k'
        status = $status
        startedUtc = $startedUtc.ToString('O')
        finishedUtc = [DateTime]::UtcNow.ToString('O')
        candidate = [ordered]@{
            installerPath = $installer
            installerSha256 = $installerHashForResult
            expectedInstallerSha256 = $ExpectedInstallerSha256.ToUpperInvariant()
            applicationPath = $application
            applicationSha256 = $applicationHashForResult
            expectedApplicationSha256 = $ExpectedApplicationSha256.ToUpperInvariant()
            processId = if ($applicationProcess) { $applicationProcess.Id } else { $null }
        }
        fixture = [ordered]@{
            path = Get-RelativeEvidencePath $fixture
            rows = $rows
            columns = $columns
            populatedCells = $populatedCells
            usedRangeRows = $usedRangeRows
            usedRangeColumns = $usedRangeColumns
            usedRangeCells = $usedRangeCells
            initialSha256 = $initialFixtureSha256
            bytes = $initialFixtureBytes
            generationElapsedMilliseconds = $generationElapsedMilliseconds
        }
        thresholds = [ordered]@{
            phaseTimeoutSeconds = $phaseTimeoutSeconds
            maxUiaHeartbeatMilliseconds = $maxUiaHeartbeatMilliseconds
            maxHeartbeatGapMilliseconds = $maxHeartbeatGapMilliseconds
            maxPrivateBytes = $maxPrivateBytes
            requiredBulkChangedCells = $bulkChangedCells
            automaticReportCellLimit = $automaticReportCellLimit
        }
        telemetry = [ordered]@{
            path = 'telemetry/large-workbook-telemetry.jsonl'
            monitorLog = 'logs/large-workbook-monitor.txt'
            monitorStartedUtc = if ($null -ne $monitorStartedUtc) { $monitorStartedUtc.ToString('O') } else { $null }
            monitorStoppedUtc = if ($null -ne $monitorStoppedUtc) { $monitorStoppedUtc.ToString('O') } else { $null }
            sampleCount = $sampleCount
            expectedMinimumSamples = $expectedMinimumSamples
            durationMilliseconds = [math]::Round($telemetryDurationMilliseconds, 3)
            requiredDurationMilliseconds = [math]::Round($requiredMonitorDurationMilliseconds, 3)
            initialLagMilliseconds = $initialTelemetryLagMilliseconds
            tailLagMilliseconds = $tailTelemetryLagMilliseconds
            failedSamples = $telemetryFailures.Count
            maxUiaDurationMilliseconds = $maxUiaDuration
            maxGapMilliseconds = $maxGap
            peakPrivateBytes = $peakPrivateBytes
            peakPrivateWorkingSetBytes = $peakPrivateWorkingSetBytes
            peakWorkingSetBytes = $peakWorkingSetBytes
            monitorForcedStop = $monitorForcedStop
            monitorJobState = $monitorJobState
        }
        phases = $phases
        assertions = $assertions
        failure = $failureMessage
    }
    Write-AcceptanceUtf8File -Path $resultPath -Content ($result | ConvertTo-Json -Depth 15)
}

if ($failed -or @($assertions | Where-Object { -not $_.passed }).Count -ne 0) {
    throw "LARGE_WORKBOOK_BENCHMARK_FAILED|$resultPath"
}
Write-Output "LARGE_WORKBOOK_BENCHMARK_PASS|$resultPath"
