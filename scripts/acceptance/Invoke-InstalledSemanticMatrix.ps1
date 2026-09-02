[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string] $InstallerPath,
    [Parameter(Mandatory)] [ValidatePattern('^[A-Fa-f0-9]{64}$')] [string] $ExpectedInstallerSha256,
    [Parameter(Mandatory)] [ValidatePattern('^[A-Fa-f0-9]{64}$')] [string] $ExpectedApplicationSha256,
    [Parameter(Mandatory)] [string] $ProbePath,
    [Parameter(Mandatory)] [string] $XlsmFixture,
    [Parameter(Mandatory)] [string] $EvidenceDirectory,
    [string] $ApplicationPath = (Join-Path $env:LOCALAPPDATA 'Programs\Excel Diff Tracker\ExcelDiffTracker.exe'),
    [string] $DatabasePath = (Join-Path $env:LOCALAPPDATA 'Excel Diff Tracker\history.db'),
    [switch] $ConfirmInstalledCandidate
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

if (-not $ConfirmInstalledCandidate) {
    throw 'The installed semantic matrix is fail-closed. Rerun against an installed candidate with -ConfirmInstalledCandidate.'
}

$installer = (Resolve-Path $InstallerPath).Path
$application = (Resolve-Path $ApplicationPath).Path
$probe = (Resolve-Path $ProbePath).Path
$sourceXlsm = (Resolve-Path $XlsmFixture).Path
if ([System.IO.Path]::GetExtension($sourceXlsm) -ne '.xlsm') { throw 'XlsmFixture must use the .xlsm extension.' }
$database = [System.IO.Path]::GetFullPath($DatabasePath)
$evidence = [System.IO.Path]::GetFullPath($EvidenceDirectory)
$fixtures = Join-Path $evidence 'fixtures'
$probeResults = Join-Path $evidence 'probe'
$reports = Join-Path $evidence 'reports'
$workbookStates = Join-Path $evidence 'workbooks'
$screenshots = Join-Path $evidence 'screenshots'
$uia = Join-Path $evidence 'uia'
$logs = Join-Path $evidence 'logs'
$xlsxPath = Join-Path $fixtures 'Installed Semantic Matrix.xlsx'
$xlsmPath = Join-Path $fixtures 'Installed Semantic Matrix.xlsm'
$resultPath = Join-Path $evidence 'installed-semantic-matrix.json'
$transcriptPath = Join-Path $logs 'installed-semantic-matrix-transcript.txt'

if (Test-Path $evidence) {
    throw "Semantic-matrix evidence directory already exists; use a fresh path: $evidence"
}
New-Item -ItemType Directory -Path $evidence, $fixtures, $probeResults, $reports, $workbookStates, $screenshots, $uia, $logs -Force | Out-Null

Import-Module (Join-Path $PSScriptRoot 'UiAutomation.psm1') -Force

$startedUtc = [DateTime]::UtcNow
$evidenceId = [Guid]::NewGuid().ToString('D')
$assertions = [System.Collections.Generic.List[object]]::new()
$phases = [System.Collections.Generic.List[object]]::new()
$failed = $false
$failure = $null
$excel = $null
$xlsxWorkbook = $null
$xlsmWorkbook = $null
$matrixSheet = $null
$activeWorkbook = $null
$activeWorkbookPath = $null
$activeFormat = $null
$activeMacroSha256 = $null
$sourceMacroSha256 = $null
$xlsmMacroSha256 = $null
$applicationProcess = $null
$mainWindow = $null
$transcriptStarted = $false

function Add-MatrixAssertion {
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

function Assert-Matrix {
    param([string] $Name, [bool] $Condition, [string] $Detail = '', [string] $Evidence = '')
    Add-MatrixAssertion -Name $Name -Passed $Condition -Detail $Detail -Evidence $Evidence
    if (-not $Condition) { throw "$Name failed. $Detail" }
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

function Get-ExcelWindow {
    if (-not $excel) { throw 'Excel is not running.' }
    Get-UiaWindowFromHandle -Handle ([long]$excel.Hwnd)
}

function Get-ZipEntrySha256 {
    param([string] $Path, [string] $EntryName)
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $entry = $archive.GetEntry($EntryName)
        if (-not $entry) { throw "Archive entry not found: $EntryName in $Path" }
        $stream = $entry.Open()
        try { Get-AcceptanceStreamSha256 -Stream $stream }
        finally { $stream.Dispose() }
    }
    finally { $archive.Dispose() }
}

function Set-ActiveMatrixWorkbook {
    param([object] $Workbook, [string] $Path, [ValidateSet('xlsx','xlsm')] [string] $Format, [string] $MacroSha256)
    $script:activeWorkbook = $Workbook
    $script:activeWorkbookPath = $Path
    $script:activeFormat = $Format
    $script:activeMacroSha256 = $MacroSha256
    [void]$Workbook.Activate()
}

function Select-ExcelCell {
    param([string] $Address)
    [void]$activeWorkbook.Activate()
    $excelWindow = Get-ExcelWindow
    Send-UiaKeys -Window $excelWindow -Keys '^g'
    [System.Windows.Forms.SendKeys]::SendWait($Address)
    [System.Windows.Forms.SendKeys]::SendWait('{ENTER}')
    Start-Sleep -Milliseconds 200
}

function Send-ExcelRibbonSequence {
    param([string[]] $Keys)
    [void]$activeWorkbook.Activate()
    $excelWindow = Get-ExcelWindow
    Set-UiaForeground -Window $excelWindow
    foreach ($key in $Keys) {
        [System.Windows.Forms.SendKeys]::SendWait($key)
        Start-Sleep -Milliseconds 180
    }
}

function Save-ExcelByKeyboard {
    [void]$activeWorkbook.Activate()
    $excelWindow = Get-ExcelWindow
    Send-UiaKeys -Window $excelWindow -Keys '^s'
    [DateTime]::UtcNow
}

function Invoke-ProbeUntilPassed {
    param(
        [long] $Sequence,
        [long] $CellChangeCount,
        [long] $SheetChangeCount,
        [string] $OutputName,
        [string[]] $AdditionalArguments = @(),
        [int] $TimeoutSeconds = 30,
        [switch] $Baseline
    )
    $outputPath = Join-Path $probeResults $OutputName
    $deadline = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
    $lastOutput = ''
    do {
        if ($applicationProcess.HasExited) { throw 'Excel Diff Tracker exited during the semantic matrix.' }
        $applicationProcess.Refresh()
        if (-not $applicationProcess.Responding) { throw 'Excel Diff Tracker reported Not Responding during the semantic matrix.' }

        $arguments = @(
            '--database', $database,
            '--workbook', $activeWorkbookPath,
            '--expected-sequence', $Sequence,
            '--require-active',
            '--require-no-errors',
            '--require-no-last-error',
            '--expected-error-count', '0',
            '--expected-version-count', $Sequence,
            '--require-unique-version-hashes')
        if (-not $Baseline) {
            $arguments += @(
                '--expected-cell-change-count', $CellChangeCount,
                '--expected-sheet-change-count', $SheetChangeCount,
                '--require-source-hash-match',
                '--require-ready-report')
        }
        $arguments += $AdditionalArguments
        $lastOutput = & $probe @arguments 2>&1 | Out-String
        $exitCode = $LASTEXITCODE
        if ($exitCode -eq 0) {
            Write-AcceptanceUtf8File -Path $outputPath -Content $lastOutput
            return $lastOutput | ConvertFrom-Json
        }
        Start-Sleep -Milliseconds 300
    } while ([DateTime]::UtcNow -lt $deadline)
    Write-AcceptanceUtf8File -Path $outputPath -Content $lastOutput
    throw "Semantic-matrix probe did not pass within $TimeoutSeconds seconds. See $outputPath"
}

function Capture-AppEvidence {
    param([long] $Sequence)
    Set-UiaForeground -Window $mainWindow
    $navigationId = if ($Sequence -eq 0) { 'DashboardNavigationButton' } else { 'HistoryNavigationButton' }
    $navigation = Find-UiaElement -Root $mainWindow -AutomationId $navigationId
    Invoke-UiaElement -Element $navigation
    Start-Sleep -Milliseconds 500
    $script:mainWindow = Find-UiaWindow -Title 'Excel Diff Tracker' -TimeoutSeconds 10
    $screenshotPath = Join-Path $screenshots ('{0}-sequence-{1:D2}.png' -f $activeFormat, $Sequence)
    $uiaPath = Join-Path $uia ('{0}-sequence-{1:D2}.json' -f $activeFormat, $Sequence)
    Save-DesktopScreenshot -Path $screenshotPath
    Export-UiaTree -Root $mainWindow -Path $uiaPath
    [pscustomobject]@{
        screenshot = Get-AcceptanceRelativePath -BasePath $evidence -Path $screenshotPath -UseForwardSlash
        screenshotSha256 = Get-AcceptanceFileSha256 -Path $screenshotPath
        uiaTree = Get-AcceptanceRelativePath -BasePath $evidence -Path $uiaPath -UseForwardSlash
        uiaTreeSha256 = Get-AcceptanceFileSha256 -Path $uiaPath
    }
}

function Complete-MatrixPhase {
    param(
        [long] $Sequence,
        [string] $Name,
        [string] $Action,
        [long] $CellChangeCount,
        [long] $SheetChangeCount,
        [object[]] $Checks,
        [DateTime] $ActionStartedUtc,
        [Nullable[DateTime]] $CtrlSaveUtc,
        [switch] $Baseline
    )
    $phaseClock = [System.Diagnostics.Stopwatch]::StartNew()
    $probePaths = [System.Collections.Generic.List[object]]::new()
    $primary = $null
    for ($index = 0; $index -lt $Checks.Count; $index++) {
        $probeName = '{0}-sequence-{1:D2}-check-{2:D2}.json' -f $activeFormat, $Sequence, ($index + 1)
        $checkArguments = [string[]]@($Checks[$index].arguments)
        $result = Invoke-ProbeUntilPassed -Sequence $Sequence -CellChangeCount $CellChangeCount -SheetChangeCount $SheetChangeCount -OutputName $probeName -AdditionalArguments $checkArguments -TimeoutSeconds $(if ($Baseline) { 60 } else { 30 }) -Baseline:$Baseline
        if ($null -eq $primary) { $primary = $result }
        $probeOutputPath = Join-Path $probeResults $probeName
        $probePaths.Add([pscustomobject]@{
            path = Get-AcceptanceRelativePath -BasePath $evidence -Path $probeOutputPath -UseForwardSlash
            sha256 = Get-AcceptanceFileSha256 -Path $probeOutputPath
        })
    }
    $phaseClock.Stop()
    if ($null -eq $primary) { throw "Phase $Sequence did not define an external probe check." }

    $sourceHash = (Get-FileHash $activeWorkbookPath -Algorithm SHA256).Hash.ToUpperInvariant()
    Assert-Matrix "sequence $Sequence captured the exact saved workbook hash" ($primary.currentHash.ToUpperInvariant() -eq $sourceHash) $sourceHash $probePaths[0].path
    if (-not $Baseline) {
        Assert-Matrix "sequence $Sequence version hash matches the saved workbook" ($primary.latestVersion.sha256.ToUpperInvariant() -eq $sourceHash) $sourceHash $probePaths[0].path
    }

    $workbookCopy = Join-Path $workbookStates ('{0}-sequence-{1:D2}.{0}' -f $activeFormat, $Sequence)
    Copy-Item -LiteralPath $activeWorkbookPath -Destination $workbookCopy
    $workbookCopyHash = Get-AcceptanceFileSha256 -Path $workbookCopy
    Assert-Matrix "sequence $Sequence portable workbook copy is byte-identical" ($workbookCopyHash -eq $sourceHash) $workbookCopyHash

    $report = $null
    if (-not $Baseline) {
        $reportCopy = Join-Path $reports ('{0}-sequence-{1:D2}.md' -f $activeFormat, $Sequence)
        Copy-Item -LiteralPath $primary.latestVersion.reportPath -Destination $reportCopy
        $report = [pscustomobject]@{
            path = Get-AcceptanceRelativePath -BasePath $evidence -Path $reportCopy -UseForwardSlash
            sha256 = Get-AcceptanceFileSha256 -Path $reportCopy
        }
    }

    $uiEvidence = Capture-AppEvidence -Sequence $Sequence
    $capturedUtc = [DateTime]::UtcNow
    $macroSha256 = if ($activeFormat -eq 'xlsm') { Get-ZipEntrySha256 -Path $workbookCopy -EntryName 'xl/vbaProject.bin' } else { $null }
    if ($activeFormat -eq 'xlsm') {
        Assert-Matrix "xlsm sequence $Sequence retains the deterministic VBA project" ($macroSha256 -eq $activeMacroSha256) "$activeMacroSha256 -> $macroSha256" $probePaths[0].path
    }
    $script:phases.Add([pscustomobject]@{
        format = $activeFormat
        sequence = $Sequence
        name = $Name
        action = $Action
        transport = if ($Baseline) { 'Excel COM fixture setup before tracking; installed-app UIA baseline registration' } else { 'visible Excel keyboard/UIA mutation and Ctrl+S' }
        actionStartedUtc = $ActionStartedUtc.ToString('O')
        ctrlSaveUtc = if ($CtrlSaveUtc.HasValue) { $CtrlSaveUtc.Value.ToString('O') } else { $null }
        capturedUtc = $capturedUtc.ToString('O')
        captureMilliseconds = [math]::Round($phaseClock.Elapsed.TotalMilliseconds, 3)
        expectedCellChangeCount = $CellChangeCount
        expectedSheetChangeCount = $SheetChangeCount
        workbookSha256 = $sourceHash
        macroSha256 = $macroSha256
        workbookEvidence = [pscustomobject]@{
            path = Get-AcceptanceRelativePath -BasePath $evidence -Path $workbookCopy -UseForwardSlash
            sha256 = $workbookCopyHash
        }
        probes = $probePaths
        report = $report
        uiEvidence = $uiEvidence
    })
}

function Move-ActiveSheet {
    param([ValidateSet('End','BeforeMatrix')] [string] $Destination)
    Send-ExcelRibbonSequence -Keys @('%h','o','m')
    $moveDialog = Find-UiaWindow -Title 'Move or Copy' -TimeoutSeconds 10
    if ($Destination -eq 'End') {
        $endDestination = Find-UiaElement -Root $moveDialog -Name '(move to end)' -Optional
        if ($endDestination) { Invoke-UiaElement -Element $endDestination }
        else { Send-UiaKeys -Window $moveDialog -Keys '{END}' }
    }
    else {
        $matrixDestination = Find-UiaElement -Root $moveDialog -Name 'Matrix' -Optional
        if ($matrixDestination) { Invoke-UiaElement -Element $matrixDestination }
        else { Send-UiaKeys -Window $moveDialog -Keys '{HOME}' }
    }
    $moveOk = Find-UiaElement -Root $moveDialog -AutomationId '1' -Optional
    if (-not $moveOk) { $moveOk = Find-UiaElement -Root $moveDialog -Name 'OK' }
    Invoke-UiaElement -Element $moveOk
    Start-Sleep -Milliseconds 300
}

function Invoke-MatrixSaves {
    Add-TrackedWorkbook -Path $activeWorkbookPath
    Complete-MatrixPhase -Sequence 0 -Name 'silent baseline' -Action "register deterministic $activeFormat workbook through the installed app" -CellChangeCount 0 -SheetChangeCount 0 -Checks @(
        [pscustomobject]@{ arguments = @() }
    ) -ActionStartedUtc ([DateTime]::UtcNow) -CtrlSaveUtc $null -Baseline

    $actionStarted = [DateTime]::UtcNow
    Select-ExcelCell 'A1'
    [System.Windows.Forms.SendKeys]::SendWait('=1{+}1')
    [System.Windows.Forms.SendKeys]::SendWait('{ENTER}')
    $savedUtc = Save-ExcelByKeyboard
    Complete-MatrixPhase -Sequence 1 -Name 'formula add' -Action 'type =1+1 into Matrix!A1' -CellChangeCount 1 -SheetChangeCount 0 -Checks @(
        [pscustomobject]@{ arguments = @('--address','A1','--expected-kind','FormulaAdded','--expected-formula-text','1+1','--expected-cached-result','2','--expect-before-missing','--report-contains','Matrix!A1') }
    ) -ActionStartedUtc $actionStarted -CtrlSaveUtc $savedUtc

    $actionStarted = [DateTime]::UtcNow
    Select-ExcelCell 'A1'
    [System.Windows.Forms.SendKeys]::SendWait('=1{+}2')
    [System.Windows.Forms.SendKeys]::SendWait('{ENTER}')
    $savedUtc = Save-ExcelByKeyboard
    Complete-MatrixPhase -Sequence 2 -Name 'formula edit' -Action 'replace Matrix!A1 with =1+2' -CellChangeCount 1 -SheetChangeCount 0 -Checks @(
        [pscustomobject]@{ arguments = @('--address','A1','--expected-kind','FormulaChanged','--expected-formula-text','1+2','--expected-cached-result','3','--report-contains','Formula changed') }
    ) -ActionStartedUtc $actionStarted -CtrlSaveUtc $savedUtc

    $actionStarted = [DateTime]::UtcNow
    Select-ExcelCell 'A1'
    [System.Windows.Forms.SendKeys]::SendWait('{DELETE}')
    $savedUtc = Save-ExcelByKeyboard
    Complete-MatrixPhase -Sequence 3 -Name 'formula delete' -Action 'delete Matrix!A1' -CellChangeCount 1 -SheetChangeCount 0 -Checks @(
        [pscustomobject]@{ arguments = @('--address','A1','--expected-kind','FormulaRemoved','--expect-formula-missing','--report-contains','Formula removed') }
    ) -ActionStartedUtc $actionStarted -CtrlSaveUtc $savedUtc

    $actionStarted = [DateTime]::UtcNow
    Select-ExcelCell 'D1'
    [System.Windows.Forms.SendKeys]::SendWait('2')
    [System.Windows.Forms.SendKeys]::SendWait('{ENTER}')
    $savedUtc = Save-ExcelByKeyboard
    Complete-MatrixPhase -Sequence 4 -Name 'formula result only' -Action 'change Matrix!D1 from 1 to 2 so Matrix!C1 recalculates without formula-text change' -CellChangeCount 2 -SheetChangeCount 0 -Checks @(
        [pscustomobject]@{ arguments = @('--address','C1','--expected-kind','FormulaResultChanged','--expected-formula-text','D1','--expected-cached-result','2','--report-contains','Calculated result') },
        [pscustomobject]@{ arguments = @('--address','D1','--expected-kind','LiteralChanged','--expected-before-value','1','--expected-value','2','--report-contains','Matrix!D1') }
    ) -ActionStartedUtc $actionStarted -CtrlSaveUtc $savedUtc

    $actionStarted = [DateTime]::UtcNow
    Select-ExcelCell 'E1'
    [System.Windows.Forms.SendKeys]::SendWait('EDT-TEXT-123')
    [System.Windows.Forms.SendKeys]::SendWait('{ENTER}')
    $savedUtc = Save-ExcelByKeyboard
    Complete-MatrixPhase -Sequence 5 -Name 'literal and type transition' -Action 'replace numeric Matrix!E1 with text EDT-TEXT-123' -CellChangeCount 1 -SheetChangeCount 0 -Checks @(
        [pscustomobject]@{ arguments = @('--address','E1','--expected-kind','CellTypeChanged','--expected-before-value','123','--expected-value','EDT-TEXT-123','--report-contains','Cell type changed') }
    ) -ActionStartedUtc $actionStarted -CtrlSaveUtc $savedUtc

    $actionStarted = [DateTime]::UtcNow
    Select-ExcelCell 'F1'
    [System.Windows.Forms.SendKeys]::SendWait('^b')
    $savedUtc = Save-ExcelByKeyboard
    Complete-MatrixPhase -Sequence 6 -Name 'style only' -Action 'toggle bold on Matrix!F1 without changing its value' -CellChangeCount 0 -SheetChangeCount 0 -Checks @(
        [pscustomobject]@{ arguments = @('--report-contains','No tracked changes') }
    ) -ActionStartedUtc $actionStarted -CtrlSaveUtc $savedUtc

    $actionStarted = [DateTime]::UtcNow
    $excelWindow = Get-ExcelWindow
    Send-UiaKeys -Window $excelWindow -Keys '+{F11}'
    Move-ActiveSheet -Destination End
    Send-ExcelRibbonSequence -Keys @('%h','o','r')
    [System.Windows.Forms.SendKeys]::SendWait('Matrix Added')
    [System.Windows.Forms.SendKeys]::SendWait('{ENTER}')
    $savedUtc = Save-ExcelByKeyboard
    Complete-MatrixPhase -Sequence 7 -Name 'sheet add' -Action 'insert a worksheet, move it to the end, and name it Matrix Added' -CellChangeCount 0 -SheetChangeCount 1 -Checks @(
        [pscustomobject]@{ arguments = @('--expected-sheet-kind','Added','--expected-sheet-name','Matrix Added','--report-contains','Added') }
    ) -ActionStartedUtc $actionStarted -CtrlSaveUtc $savedUtc

    $actionStarted = [DateTime]::UtcNow
    Send-ExcelRibbonSequence -Keys @('%h','o','r')
    [System.Windows.Forms.SendKeys]::SendWait('Matrix Renamed')
    [System.Windows.Forms.SendKeys]::SendWait('{ENTER}')
    $savedUtc = Save-ExcelByKeyboard
    Complete-MatrixPhase -Sequence 8 -Name 'sheet rename' -Action 'rename Matrix Added to Matrix Renamed' -CellChangeCount 0 -SheetChangeCount 1 -Checks @(
        [pscustomobject]@{ arguments = @('--expected-sheet-kind','Renamed','--expected-sheet-name','Matrix Renamed','--report-contains','Renamed') }
    ) -ActionStartedUtc $actionStarted -CtrlSaveUtc $savedUtc

    $actionStarted = [DateTime]::UtcNow
    Move-ActiveSheet -Destination BeforeMatrix
    $savedUtc = Save-ExcelByKeyboard
    Complete-MatrixPhase -Sequence 9 -Name 'sheet reorder' -Action 'move Matrix Renamed before Matrix through the Move or Copy dialog' -CellChangeCount 0 -SheetChangeCount 2 -Checks @(
        [pscustomobject]@{ arguments = @('--expected-sheet-kind','Reordered','--expected-sheet-name','Matrix Renamed','--report-contains','Reordered') }
    ) -ActionStartedUtc $actionStarted -CtrlSaveUtc $savedUtc

    $actionStarted = [DateTime]::UtcNow
    Send-ExcelRibbonSequence -Keys @('%h','o','u','s')
    $savedUtc = Save-ExcelByKeyboard
    Complete-MatrixPhase -Sequence 10 -Name 'sheet hide' -Action 'hide Matrix Renamed from the Home ribbon' -CellChangeCount 0 -SheetChangeCount 1 -Checks @(
        [pscustomobject]@{ arguments = @('--expected-sheet-kind','VisibilityChanged','--expected-sheet-name','Matrix Renamed','--report-contains','Visibility changed') }
    ) -ActionStartedUtc $actionStarted -CtrlSaveUtc $savedUtc

    $actionStarted = [DateTime]::UtcNow
    Send-ExcelRibbonSequence -Keys @('%h','o','u','h')
    $unhideDialog = Find-UiaWindow -Title 'Unhide' -TimeoutSeconds 10
    $hiddenSheet = Find-UiaElement -Root $unhideDialog -Name 'Matrix Renamed' -Optional
    if ($hiddenSheet) { Invoke-UiaElement -Element $hiddenSheet }
    $unhideOk = Find-UiaElement -Root $unhideDialog -AutomationId '1' -Optional
    if (-not $unhideOk) { $unhideOk = Find-UiaElement -Root $unhideDialog -Name 'OK' }
    Invoke-UiaElement -Element $unhideOk
    Start-Sleep -Milliseconds 300
    $savedUtc = Save-ExcelByKeyboard
    Complete-MatrixPhase -Sequence 11 -Name 'sheet unhide' -Action 'unhide Matrix Renamed through the Unhide dialog' -CellChangeCount 0 -SheetChangeCount 1 -Checks @(
        [pscustomobject]@{ arguments = @('--expected-sheet-kind','VisibilityChanged','--expected-sheet-name','Matrix Renamed','--report-contains','Visibility changed') }
    ) -ActionStartedUtc $actionStarted -CtrlSaveUtc $savedUtc

    $actionStarted = [DateTime]::UtcNow
    $excelWindow = Get-ExcelWindow
    Send-UiaKeys -Window $excelWindow -Keys '^{PGUP}'
    Send-ExcelRibbonSequence -Keys @('%h','d','s')
    Start-Sleep -Milliseconds 300
    $savedUtc = Save-ExcelByKeyboard
    Complete-MatrixPhase -Sequence 12 -Name 'sheet remove' -Action 'activate and delete Matrix Renamed from the Home ribbon' -CellChangeCount 0 -SheetChangeCount 2 -Checks @(
        [pscustomobject]@{ arguments = @('--expected-sheet-kind','Removed','--expected-sheet-name','Matrix Renamed','--report-contains','Removed') }
    ) -ActionStartedUtc $actionStarted -CtrlSaveUtc $savedUtc

    $formatPhases = @($phases | Where-Object format -eq $activeFormat)
    $distinctHashes = @($formatPhases | Select-Object -ExpandProperty workbookSha256 -Unique)
    Assert-Matrix "$activeFormat baseline and twelve saves have thirteen distinct stable hashes" ($formatPhases.Count -eq 13 -and $distinctHashes.Count -eq 13) "phases=$($formatPhases.Count) distinct=$($distinctHashes.Count)"
}

try {
    Start-Transcript -Path $transcriptPath -Force | Out-Null
    $transcriptStarted = $true

    $installerHash = Get-AcceptanceFileSha256 -Path $installer
    $applicationHash = Get-AcceptanceFileSha256 -Path $application
    $probeHash = Get-AcceptanceFileSha256 -Path $probe
    Assert-Matrix 'semantic matrix uses the frozen installer' ($installerHash -eq $ExpectedInstallerSha256.ToUpperInvariant()) $installerHash
    Assert-Matrix 'semantic matrix uses the frozen installed executable' ($applicationHash -eq $ExpectedApplicationSha256.ToUpperInvariant()) $applicationHash
    Assert-Matrix 'installed app database exists' (Test-Path $database -PathType Leaf) $database

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3

    $xlsxWorkbook = $excel.Workbooks.Add()
    while ($xlsxWorkbook.Worksheets.Count -gt 1) { $xlsxWorkbook.Worksheets.Item($xlsxWorkbook.Worksheets.Count).Delete() }
    $matrixSheet = $xlsxWorkbook.Worksheets.Item(1)
    $matrixSheet.Name = 'Matrix'
    $matrixSheet.Range('D1').Value2 = 1
    $matrixSheet.Range('C1').Formula = '=D1'
    $matrixSheet.Range('E1').Value2 = 123
    $matrixSheet.Range('F1').Value2 = 'Style sentinel'
    $excel.CalculateFull()
    $xlsxWorkbook.SaveAs($xlsxPath, 51)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($matrixSheet)
    $matrixSheet = $null

    Copy-Item -LiteralPath $sourceXlsm -Destination $xlsmPath
    $sourceMacroSha256 = Get-ZipEntrySha256 -Path $xlsmPath -EntryName 'xl/vbaProject.bin'
    $xlsmWorkbook = $excel.Workbooks.Open($xlsmPath)
    Assert-Matrix 'xlsm fixture contains exactly one sheet' ($xlsmWorkbook.Sheets.Count -eq 1 -and $xlsmWorkbook.Worksheets.Count -eq 1) "sheets=$($xlsmWorkbook.Sheets.Count) worksheets=$($xlsmWorkbook.Worksheets.Count)"
    $matrixSheet = $xlsmWorkbook.Worksheets.Item(1)
    $matrixSheet.Name = 'Matrix'
    $matrixSheet.Cells.Clear()
    $matrixSheet.Range('D1').Value2 = 1
    $matrixSheet.Range('C1').Formula = '=D1'
    $matrixSheet.Range('E1').Value2 = 123
    $matrixSheet.Range('F1').Value2 = 'Style sentinel'
    $excel.CalculateFull()
    $xlsmWorkbook.Save()
    $xlsmMacroSha256 = Get-ZipEntrySha256 -Path $xlsmPath -EntryName 'xl/vbaProject.bin'
    Assert-Matrix 'xlsm fixture preparation retains its deterministic VBA project' ($xlsmMacroSha256 -eq $sourceMacroSha256) "$sourceMacroSha256 -> $xlsmMacroSha256"
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($matrixSheet)
    $matrixSheet = $null

    Start-Process $application | Out-Null
    $mainWindow = Find-UiaWindow -Title 'Excel Diff Tracker' -TimeoutSeconds 20
    $applicationProcess = Get-Process -Id $mainWindow.Current.ProcessId -ErrorAction Stop
    Assert-Matrix 'semantic matrix controls the installed executable' ([string]::Equals($applicationProcess.Path, $application, [StringComparison]::OrdinalIgnoreCase)) $applicationProcess.Path

    Set-ActiveMatrixWorkbook -Workbook $xlsxWorkbook -Path $xlsxPath -Format xlsx -MacroSha256 $null
    Invoke-MatrixSaves
    Set-ActiveMatrixWorkbook -Workbook $xlsmWorkbook -Path $xlsmPath -Format xlsm -MacroSha256 $xlsmMacroSha256
    Invoke-MatrixSaves

}
catch {
    $failed = $true
    $failure = $_.Exception.ToString()
    Add-MatrixAssertion -Name 'unhandled installed semantic-matrix step' -Passed $false -Detail $failure -Evidence 'installed-semantic-matrix.json'
    try { Save-DesktopScreenshot -Path (Join-Path $screenshots 'semantic-matrix-failure.png') } catch { }
}
finally {
    try { if ($xlsmWorkbook) { $xlsmWorkbook.Close($false) } } catch { }
    try { if ($xlsxWorkbook) { $xlsxWorkbook.Close($false) } } catch { }
    try { if ($excel) { $excel.Quit() } } catch { }
    foreach ($value in @($matrixSheet, $xlsmWorkbook, $xlsxWorkbook, $excel)) {
        if ($value) { try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value) } catch { } }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    if ($transcriptStarted) {
        try { Stop-Transcript | Out-Null } catch { }
    }
    $finishedUtc = [DateTime]::UtcNow
    $result = [ordered]@{
        schemaVersion = 1
        gate = 'installed-real-excel-semantic-matrix'
        evidenceId = $evidenceId
        status = if (-not $failed -and $phases.Count -eq 26 -and @($assertions | Where-Object { -not $_.passed }).Count -eq 0) { 'Passed' } else { 'Failed' }
        startedUtc = $startedUtc.ToString('O')
        finishedUtc = $finishedUtc.ToString('O')
        durationSeconds = [math]::Round(($finishedUtc - $startedUtc).TotalSeconds, 3)
        candidate = [ordered]@{
            installerSha256 = if (Test-Path $installer) { Get-AcceptanceFileSha256 -Path $installer } else { $null }
            expectedInstallerSha256 = $ExpectedInstallerSha256.ToUpperInvariant()
            applicationSha256 = if (Test-Path $application) { Get-AcceptanceFileSha256 -Path $application } else { $null }
            expectedApplicationSha256 = $ExpectedApplicationSha256.ToUpperInvariant()
            probeSha256 = if (Test-Path $probe) { Get-AcceptanceFileSha256 -Path $probe } else { $null }
        }
        workbooks = @(
            [ordered]@{ format = 'xlsx'; path = $xlsxPath; finalSequence = 12; sourceMacroSha256 = $null; baselineMacroSha256 = $null },
            [ordered]@{ format = 'xlsm'; path = $xlsmPath; finalSequence = 12; sourceMacroSha256 = $sourceMacroSha256; baselineMacroSha256 = $xlsmMacroSha256 }
        )
        phases = $phases
        assertions = $assertions
        transcript = [ordered]@{
            path = 'logs/installed-semantic-matrix-transcript.txt'
            sha256 = if (Test-Path $transcriptPath) { Get-AcceptanceFileSha256 -Path $transcriptPath } else { $null }
        }
        failure = $failure
    }
    Write-AcceptanceUtf8File -Path $resultPath -Content ($result | ConvertTo-Json -Depth 15)
}

if ($failed -or $phases.Count -ne 26 -or $result.status -ne 'Passed' -or @($assertions | Where-Object { -not $_.passed }).Count -ne 0) {
    throw "INSTALLED_SEMANTIC_MATRIX_FAILED|$resultPath"
}
Write-Output "INSTALLED_SEMANTIC_MATRIX_PASS|$resultPath"
