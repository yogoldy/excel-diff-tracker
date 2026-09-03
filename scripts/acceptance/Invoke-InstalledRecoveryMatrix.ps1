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
    [string] $PasswordProtectedFixture,
    [ValidateRange(5, 15)] [int] $ShortLockSeconds = 5,
    [ValidateRange(65, 120)] [int] $FailureExposureSeconds = 65,
    [ValidateRange(5000, 50000)] [int] $LargeWorkbookRows = 25000,
    [switch] $ConfirmDisposableCleanVm
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

if (-not $ConfirmDisposableCleanVm) {
    throw 'The installed recovery matrix mutates installed-app state and is fail-closed. Run only in a disposable clean Windows VM and pass -ConfirmDisposableCleanVm.'
}
if ($PSVersionTable.PSVersion.Major -ne 5 -or $PSVersionTable.PSEdition -ne 'Desktop') {
    throw "Windows PowerShell 5.1 Desktop is required; found $($PSVersionTable.PSVersion) $($PSVersionTable.PSEdition)."
}

$installer = (Resolve-Path $InstallerPath).Path
$application = (Resolve-Path $ApplicationPath).Path
$probe = (Resolve-Path $ProbePath).Path
$database = [System.IO.Path]::GetFullPath($DatabasePath)
$evidence = [System.IO.Path]::GetFullPath($EvidenceDirectory)
if ($evidence.StartsWith('\\')) { throw 'Use a fresh local evidence directory, not a UNC path.' }
if (Test-Path $evidence) { throw "Recovery-matrix evidence directory already exists; use a fresh path: $evidence" }

$fixtures = Join-Path $evidence 'fixtures'
$probeDirectory = Join-Path $evidence 'probe'
$databaseDirectory = Join-Path $evidence 'database'
$workbookDirectory = Join-Path $evidence 'workbooks'
$reportEvidenceDirectory = Join-Path $evidence 'reports'
$generatedReports = Join-Path $evidence 'generated-reports'
$uiDirectory = Join-Path $evidence 'uia'
$screenshotDirectory = Join-Path $evidence 'screenshots'
$logDirectory = Join-Path $evidence 'logs'
$resultPath = Join-Path $evidence 'installed-recovery-matrix.json'
$transcriptPath = Join-Path $logDirectory 'installed-recovery-matrix-transcript.txt'
New-Item -ItemType Directory -Path $evidence,$fixtures,$probeDirectory,$databaseDirectory,$workbookDirectory,$reportEvidenceDirectory,$generatedReports,$uiDirectory,$screenshotDirectory,$logDirectory -Force | Out-Null

Import-Module (Join-Path $PSScriptRoot 'UiAutomation.psm1') -Force

$startedUtc = [DateTime]::UtcNow
$evidenceId = [Guid]::NewGuid().ToString('D')
$scenarios = [System.Collections.Generic.List[object]]::new()
$activeScenario = $null
$applicationProcess = $null
$mainWindow = $null
$excel = $null
$openWorkbook = $null
$transcriptStarted = $false
$fatalFailure = $null

function Assert-RecoveryCondition {
    param([bool] $Condition, [string] $Message)
    if (-not $Condition) { throw $Message }
}

function Get-RelativePath {
    param([string] $Path)
    Get-AcceptanceRelativePath -BasePath $evidence -Path $Path -UseForwardSlash
}

function Get-FileRecord {
    param([string] $Path)
    $item = Get-Item -LiteralPath $Path
    [pscustomobject]@{
        path = Get-RelativePath $item.FullName
        sha256 = Get-AcceptanceFileSha256 -Path $item.FullName
        bytes = $item.Length
    }
}

function Copy-SharedFile {
    param([string] $Source, [string] $Destination)
    $parent = Split-Path -Parent $Destination
    New-Item -ItemType Directory -Path $parent -Force | Out-Null
    $input = [System.IO.FileStream]::new($Source,[System.IO.FileMode]::Open,[System.IO.FileAccess]::Read,([System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete))
    try {
        $output = [System.IO.FileStream]::new($Destination,[System.IO.FileMode]::CreateNew,[System.IO.FileAccess]::Write,[System.IO.FileShare]::None)
        try { $input.CopyTo($output); $output.Flush() } finally { $output.Dispose() }
    }
    finally { $input.Dispose() }
}

function Copy-DatabaseEvidence {
    param([string] $ScenarioId, [string] $PhaseName)
    $safeName = ($PhaseName -replace '[^A-Za-z0-9_.-]','-').ToLowerInvariant()
    $destination = Join-Path (Join-Path $databaseDirectory $ScenarioId) $safeName
    New-Item -ItemType Directory -Path $destination -Force | Out-Null
    $records = [System.Collections.Generic.List[object]]::new()
    foreach ($suffix in @('','-wal','-shm')) {
        $source = $database + $suffix
        if (Test-Path $source -PathType Leaf) {
            $target = Join-Path $destination ((Split-Path $database -Leaf) + $suffix)
            Copy-SharedFile -Source $source -Destination $target
            $records.Add((Get-FileRecord $target))
        }
    }
    Assert-RecoveryCondition ($records.Count -ge 1) "No database evidence could be retained for $ScenarioId/$PhaseName."
    @($records)
}

function Start-InstalledApplication {
    $existing = @(Get-Process -Name 'ExcelDiffTracker' -ErrorAction SilentlyContinue | Where-Object {
        try { $_.Path -eq $application } catch { $false }
    })
    if ($existing.Count -gt 0) { throw 'Excel Diff Tracker is already running. Exit it before starting this disposable-VM matrix.' }
    $launched = Start-Process -FilePath $application -PassThru
    Start-Sleep -Milliseconds 800
    if (-not $launched.HasExited) { $script:applicationProcess = $launched }
    else {
        $script:applicationProcess = Get-Process -Name 'ExcelDiffTracker' -ErrorAction SilentlyContinue | Where-Object {
            try { $_.Path -eq $application } catch { $false }
        } | Select-Object -First 1
    }
    Assert-RecoveryCondition ($null -ne $script:applicationProcess) 'The installed application did not start.'
    $script:mainWindow = Find-UiaWindow -Title 'Excel Diff Tracker' -ProcessId $script:applicationProcess.Id -TimeoutSeconds 20
}

function Restart-InstalledApplication {
    $script:applicationProcess = Start-Process -FilePath $application -PassThru
    $script:mainWindow = Find-UiaWindow -Title 'Excel Diff Tracker' -ProcessId $script:applicationProcess.Id -TimeoutSeconds 20
}

function Stop-InstalledApplication {
    Assert-RecoveryCondition ($null -ne $script:applicationProcess -and -not $script:applicationProcess.HasExited) 'The installed application is not running.'
    $script:applicationProcess.Kill()
    Assert-RecoveryCondition ($script:applicationProcess.WaitForExit(10000)) 'The installed application did not stop within 10 seconds.'
}

function Set-FolderBrowserPath {
    param([System.Windows.Automation.AutomationElement] $Dialog, [string] $Path)
    $edit = Find-UiaElement -Root $Dialog -AutomationId '1001' -Optional
    if (-not $edit) { $edit = Find-UiaElement -Root $Dialog -AutomationId '1152' -Optional }
    if ($edit) {
        Set-UiaValue -Element $edit -Value $Path
    }
    else {
        Send-UiaKeys -Window $Dialog -Keys '%d'
        [System.Windows.Forms.SendKeys]::SendWait($Path)
        [System.Windows.Forms.SendKeys]::SendWait('{ENTER}')
        Start-Sleep -Milliseconds 500
    }
    $select = Find-UiaElement -Root $Dialog -AutomationId '1' -Optional
    if (-not $select) { $select = Find-UiaElement -Root $Dialog -Name 'Select Folder' -Optional }
    if (-not $select) { $select = Find-UiaElement -Root $Dialog -Name 'OK' }
    Invoke-UiaElement -Element $select
}

function Set-DefaultReportDirectory {
    param([string] $Path)
    Set-UiaForeground -Window $mainWindow
    Invoke-UiaElement -Element (Find-UiaElement -Root $mainWindow -AutomationId 'SettingsNavigationButton')
    Start-Sleep -Milliseconds 300
    $script:mainWindow = Find-UiaWindow -Title 'Excel Diff Tracker' -ProcessId $applicationProcess.Id -TimeoutSeconds 10
    Invoke-UiaElement -Element (Find-UiaElement -Root $mainWindow -AutomationId 'ChooseDefaultReportFolderButton')
    $dialog = Find-UiaWindow -Title 'Choose where Excel Diff Tracker should save Markdown reports' -TimeoutSeconds 15
    Set-FolderBrowserPath -Dialog $dialog -Path $Path
    Start-Sleep -Milliseconds 500
}

function Set-WorkbookReportDirectory {
    param([string] $WorkbookPath, [string] $Path)
    Set-UiaForeground -Window $mainWindow
    Invoke-UiaElement -Element (Find-UiaElement -Root $mainWindow -AutomationId 'WorkbooksNavigationButton')
    Start-Sleep -Milliseconds 400
    $script:mainWindow = Find-UiaWindow -Title 'Excel Diff Tracker' -ProcessId $applicationProcess.Id -TimeoutSeconds 10
    $pathElement = Find-UiaElement -Root $mainWindow -Name $WorkbookPath
    $walker = [System.Windows.Automation.TreeWalker]::ControlViewWalker
    $ancestor = $pathElement
    $button = $null
    for ($index = 0; $index -lt 8 -and $ancestor; $index++) {
        $button = Find-UiaElement -Root $ancestor -Name ('Report folder' + [char]0x2026) -Optional
        if ($button) { break }
        $ancestor = $walker.GetParent($ancestor)
    }
    Assert-RecoveryCondition ($null -ne $button) "Could not locate the report-folder action for $WorkbookPath."
    Invoke-UiaElement -Element $button
    $title = "Choose the Markdown report folder for $(Split-Path $WorkbookPath -Leaf)"
    $dialog = Find-UiaWindow -Title $title -TimeoutSeconds 15
    Set-FolderBrowserPath -Dialog $dialog -Path $Path
    Start-Sleep -Milliseconds 500
}

function Choose-WorkbookFromDialog {
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
    Invoke-UiaElement -Element (Find-UiaElement -Root $mainWindow -AutomationId 'DashboardNavigationButton')
    Start-Sleep -Milliseconds 250
    $script:mainWindow = Find-UiaWindow -Title 'Excel Diff Tracker' -ProcessId $applicationProcess.Id -TimeoutSeconds 10
    Invoke-UiaElement -Element (Find-UiaElement -Root $mainWindow -AutomationId 'DashboardAddWorkbookButton')
    Choose-WorkbookFromDialog -Path $Path
}

function New-WorkbookFixture {
    param([string] $Path, [string] $Value, [int] $Rows = 1)
    $app = New-Object -ComObject Excel.Application
    $book = $null
    try {
        $app.Visible = $false
        $app.DisplayAlerts = $false
        $book = $app.Workbooks.Add()
        $sheet = $book.Worksheets.Item(1)
        $sheet.Name = 'Recovery'
        if ($Rows -le 1) {
            $sheet.Range('A1').Value2 = $Value
        }
        else {
            $values = [object[,]]::new($Rows,20)
            for ($row = 0; $row -lt $Rows; $row++) {
                for ($column = 0; $column -lt 20; $column++) { $values[$row,$column] = "$Value-$row-$column" }
            }
            $sheet.Range("A1:T$Rows").Value2 = $values
        }
        $book.SaveAs($Path,51)
        $book.Close($false)
        $book = $null
    }
    finally {
        if ($book) { $book.Close($false) }
        $app.Quit()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)
    }
}

function Set-WorkbookValue {
    param([string] $Path, [string] $Value, [string] $Address = 'A1')
    $app = New-Object -ComObject Excel.Application
    $book = $null
    try {
        $app.Visible = $false
        $app.DisplayAlerts = $false
        $book = $app.Workbooks.Open($Path)
        $book.Worksheets.Item(1).Range($Address).Value2 = $Value
        $book.Save()
        $book.Close($false)
        $book = $null
    }
    finally {
        if ($book) { $book.Close($false) }
        $app.Quit()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)
    }
}

function Save-WorkbookAs {
    param([string] $Source, [string] $Destination, [string] $Value)
    $app = New-Object -ComObject Excel.Application
    $book = $null
    try {
        $app.Visible = $true
        $app.DisplayAlerts = $false
        $book = $app.Workbooks.Open($Source)
        $book.Worksheets.Item(1).Range('A1').Value2 = $Value
        $book.SaveAs($Destination,51)
        $book.Close($false)
        $book = $null
    }
    finally {
        if ($book) { $book.Close($false) }
        $app.Quit()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)
    }
}

function New-PasswordProtectedFixture {
    param([string] $Path)
    if ($PasswordProtectedFixture) {
        Copy-Item -LiteralPath (Resolve-Path $PasswordProtectedFixture).Path -Destination $Path
        return 'supplied-fixture'
    }
    $app = New-Object -ComObject Excel.Application
    $book = $null
    try {
        $app.Visible = $false
        $app.DisplayAlerts = $false
        $book = $app.Workbooks.Add()
        $book.Worksheets.Item(1).Range('A1').Value2 = 'encrypted-rejection'
        $book.SaveAs($Path,51,'EDT-Recovery-Matrix-Only')
        $book.Close($false)
        $book = $null
        return 'excel-saveas-password'
    }
    catch {
        throw "Excel could not safely create the password-protected fixture. Supply one with -PasswordProtectedFixture. $($_.Exception.Message)"
    }
    finally {
        if ($book) { $book.Close($false) }
        $app.Quit()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)
    }
}

function Replace-WorkbookAtomically {
    param([string] $Replacement, [string] $Target, [string] $Backup)
    if (Test-Path $Backup) { Remove-Item -LiteralPath $Backup -Force }
    [System.IO.File]::Replace($Replacement,$Target,$Backup,$true)
}

function Open-ExclusiveWorkbookLock {
    param([string] $Path, [int] $TimeoutSeconds = 3)
    $deadline = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
    do {
        try {
            return [System.IO.FileStream]::new($Path,[System.IO.FileMode]::Open,[System.IO.FileAccess]::ReadWrite,[System.IO.FileShare]::None)
        }
        catch [System.IO.IOException] {
            Start-Sleep -Milliseconds 25
        }
    } while ([DateTime]::UtcNow -lt $deadline)
    throw "Could not acquire an exclusive test lock on $Path within $TimeoutSeconds seconds."
}

function Invoke-ProbeWait {
    param(
        [string] $WorkbookPath,
        [string[]] $Arguments,
        [string] $OutputPath,
        [int] $TimeoutSeconds = 90
    )
    $deadline = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
    $last = ''
    do {
        $allArguments = @('--database',$database,'--workbook',$WorkbookPath) + $Arguments
        $last = & $probe @allArguments 2>&1 | Out-String
        $exit = $LASTEXITCODE
        if ($exit -eq 0) {
            Write-AcceptanceUtf8File -Path $OutputPath -Content $last
            return ($last | ConvertFrom-Json)
        }
        Start-Sleep -Milliseconds 300
    } while ([DateTime]::UtcNow -lt $deadline)
    Write-AcceptanceUtf8File -Path $OutputPath -Content $last
    throw "AcceptanceProbe did not reach the required state within $TimeoutSeconds seconds. See $OutputPath"
}

function Assert-RecoveryLiteralDelta {
    param([object] $Phase, [string] $Address, [string] $BeforeValue, [string] $AfterValue)
    $observed = $Phase.result
    Assert-RecoveryCondition ($observed.latestVersion.cellChangeCount -eq 1 -and $observed.latestVersion.sheetChangeCount -eq 0) "$($Phase.name) must contain exactly one changed cell and no sheet changes."
    Assert-RecoveryCondition ($observed.latestVersion.sequence -eq $observed.currentSequence -and $observed.latestVersion.sha256 -ceq $observed.currentHash) "$($Phase.name) latest version differs from its captured state."
    Assert-RecoveryCondition ($observed.cellChange.sheetName -ceq 'Recovery' -and $observed.cellChange.address -ceq $Address -and $observed.cellChange.kinds -ceq 'LiteralChanged') "$($Phase.name) changed the wrong cell or change kind."
    $before = $observed.cellChange.beforeJson | ConvertFrom-Json
    $after = $observed.cellChange.afterJson | ConvertFrom-Json
    Assert-RecoveryCondition ($before.literalValue -ceq $BeforeValue -and $after.literalValue -ceq $AfterValue) "$($Phase.name) literal contents differ from the fixture mutation."
    foreach ($state in @($before,$after)) {
        Assert-RecoveryCondition ($state.sheetName -ceq 'Recovery' -and $state.address -ceq $Address -and $state.cellType -ceq 'SharedString') "$($Phase.name) cell identity or type differs."
        foreach ($field in @('formulaText','formulaType','formulaReference','formulaSharedIndex','formulaXml','cachedResult')) {
            Assert-RecoveryCondition ($null -eq $state.$field) "$($Phase.name) unexpectedly contains formula data: $field."
        }
    }
    if ($observed.latestVersion.reportStatus -eq 'Ready') {
        Assert-RecoveryCondition ($null -ne $Phase.report) "$($Phase.name) has no retained ready report."
        $markdown = [System.IO.File]::ReadAllText((Join-Path $evidence $Phase.report.path))
        $requiredLines = @(
            "- Version: $($observed.currentSequence)",
            "- Current SHA-256: ``$($observed.currentHash)``",
            '| Sheet changes | 0 |','| Changed cells | 1 |','| Literal value changes | 1 |',
            '| Formula text changes | 0 |','| Calculated result changes | 0 |','| Cell type changes | 0 |',
            "### Recovery!$Address",'**Changes:** Literal changed',
            '| Type | <pre>SharedString</pre> | <pre>SharedString</pre> |',
            "| Literal value | <pre>$BeforeValue</pre> | <pre>$AfterValue</pre> |")
        foreach ($line in $requiredLines) {
            Assert-RecoveryCondition ([regex]::Matches($markdown,('(?m)^' + [regex]::Escape($line) + '\r?$')).Count -eq 1) "$($Phase.name) report omits or duplicates the expected line: $line"
        }
        Assert-RecoveryCondition ([regex]::Matches($markdown,'(?m)^### ').Count -eq 1) "$($Phase.name) report contains an unexpected cell detail."
    }
}

function Add-ProbePhase {
    param(
        [string] $Name,
        [string] $WorkbookPath,
        [string[]] $Arguments,
        [int] $TimeoutSeconds = 90,
        [string] $ExpectedSourceSha256 = '',
        [string] $ExpectedBeforeValue = '',
        [string] $ExpectedAfterValue = '',
        [string] $ExpectedAddress = 'A1',
        [switch] $SourceInaccessible
    )
    $scenarioId = $activeScenario.id
    $safeName = ($Name -replace '[^A-Za-z0-9_.-]','-').ToLowerInvariant()
    $probePath = Join-Path (Join-Path $probeDirectory $scenarioId) ($safeName + '.json')
    if (-not [string]::IsNullOrWhiteSpace($ExpectedAfterValue)) {
        $Arguments += @('--address',$ExpectedAddress,'--expected-kind','LiteralChanged','--expected-before-value',$ExpectedBeforeValue,'--expected-value',$ExpectedAfterValue,'--expected-cell-change-count','1','--expected-sheet-change-count','0')
    }
    $probeResult = Invoke-ProbeWait -WorkbookPath $WorkbookPath -Arguments $Arguments -OutputPath $probePath -TimeoutSeconds $TimeoutSeconds
    $sourceRecord = [pscustomobject]@{ exists = Test-Path $WorkbookPath -PathType Leaf; inaccessible = [bool]$SourceInaccessible; path = $null; sha256 = $null; bytes = $null }
    if ($sourceRecord.exists -and -not $SourceInaccessible) {
        $copyPath = Join-Path (Join-Path $workbookDirectory $scenarioId) ($safeName + [System.IO.Path]::GetExtension($WorkbookPath))
        Copy-SharedFile -Source $WorkbookPath -Destination $copyPath
        $fileRecord = Get-FileRecord $copyPath
        $sourceRecord.path = $fileRecord.path
        $sourceRecord.sha256 = $fileRecord.sha256
        $sourceRecord.bytes = $fileRecord.bytes
        if (-not [string]::IsNullOrWhiteSpace($ExpectedSourceSha256)) {
            Assert-RecoveryCondition ($fileRecord.sha256 -eq $ExpectedSourceSha256.ToUpperInvariant()) "$scenarioId/$Name source hash differs from the expected bytes."
        }
    }
    $reportRecord = $null
    if ($null -ne $probeResult.latestVersion -and -not [string]::IsNullOrWhiteSpace([string]$probeResult.latestVersion.reportPath) -and (Test-Path $probeResult.latestVersion.reportPath -PathType Leaf)) {
        $reportPath = Join-Path (Join-Path $reportEvidenceDirectory $scenarioId) ($safeName + '.md')
        Copy-SharedFile -Source $probeResult.latestVersion.reportPath -Destination $reportPath
        $reportRecord = Get-FileRecord $reportPath
    }
    $phase = [pscustomobject]@{
        name = $Name
        observedUtc = [DateTime]::UtcNow.ToString('O')
        workbookPath = $WorkbookPath
        expectedSourceSha256 = if ([string]::IsNullOrWhiteSpace($ExpectedSourceSha256)) { $null } else { $ExpectedSourceSha256.ToUpperInvariant() }
        source = $sourceRecord
        probe = Get-FileRecord $probePath
        database = @(Copy-DatabaseEvidence -ScenarioId $scenarioId -PhaseName $Name)
        report = $reportRecord
        result = $probeResult
    }
    $activeScenario.phases.Add($phase)
    if (-not [string]::IsNullOrWhiteSpace($ExpectedAfterValue)) {
        Assert-RecoveryLiteralDelta -Phase $phase -Address $ExpectedAddress -BeforeValue $ExpectedBeforeValue -AfterValue $ExpectedAfterValue
    }
    $probeResult
}

function Add-ManualPhase {
    param([string] $Name, [object] $Detail)
    $activeScenario.phases.Add([pscustomobject]@{
        name = $Name
        observedUtc = [DateTime]::UtcNow.ToString('O')
        detail = $Detail
        expectedSourceSha256 = $null
        source = $null
        probe = $null
        database = @(Copy-DatabaseEvidence -ScenarioId $activeScenario.id -PhaseName $Name)
        report = $null
        result = $null
    })
}

function Register-ScenarioWorkbook {
    param([string] $ScenarioId, [string] $Value = 'baseline', [int] $Rows = 1, [string] $BaselineName = 'baseline', [string] $ExpectedReportRoot = $generatedReports)
    $path = Join-Path $fixtures ($ScenarioId + '.xlsx')
    New-WorkbookFixture -Path $path -Value $Value -Rows $Rows
    $hash = Get-AcceptanceFileSha256 -Path $path
    Add-TrackedWorkbook -Path $path
    $baseline = Add-ProbePhase -Name $BaselineName -WorkbookPath $path -Arguments @('--expected-sequence','0','--expected-version-count','0','--expected-error-count','0','--require-active','--require-no-last-error','--require-unique-version-hashes','--require-source-hash-match') -TimeoutSeconds 60 -ExpectedSourceSha256 $hash
    Assert-RecoveryCondition ($baseline.reportDirectory.Equals([System.IO.Path]::GetFullPath($ExpectedReportRoot),[StringComparison]::OrdinalIgnoreCase)) "$ScenarioId was registered outside the exact expected fresh report directory: $($baseline.reportDirectory)"
    $path
}

function Capture-ScenarioUi {
    param([string] $ScenarioId)
    $script:mainWindow = Find-UiaWindow -Title 'Excel Diff Tracker' -ProcessId $applicationProcess.Id -TimeoutSeconds 10
    Set-UiaForeground -Window $mainWindow
    $png = Join-Path $screenshotDirectory ($ScenarioId + '.png')
    $uiaPath = Join-Path $uiDirectory ($ScenarioId + '.json')
    Save-DesktopScreenshot -Path $png
    Export-UiaTree -Root $mainWindow -Path $uiaPath
    $activeScenario.uiEvidence = [pscustomobject]@{ screenshot = Get-FileRecord $png; tree = Get-FileRecord $uiaPath }
}

function Invoke-RecoveryScenario {
    param([string] $Id, [scriptblock] $Action)
    $scenario = [pscustomobject]@{
        id = $Id
        startedUtc = [DateTime]::UtcNow.ToString('O')
        finishedUtc = $null
        passed = $false
        failure = $null
        phases = [System.Collections.Generic.List[object]]::new()
        uiEvidence = $null
    }
    $script:activeScenario = $scenario
    try {
        & $Action
        Capture-ScenarioUi -ScenarioId $Id
        $scenario.passed = $true
    }
    catch {
        $scenario.failure = $_.Exception.ToString()
        if ($null -eq $script:fatalFailure) { $script:fatalFailure = "${Id}: $($_.Exception.Message)" }
    }
    finally {
        $scenario.finishedUtc = [DateTime]::UtcNow.ToString('O')
        $scenarios.Add($scenario)
        $script:activeScenario = $null
    }
}

try {
    Start-Transcript -Path $transcriptPath -Force | Out-Null
    $transcriptStarted = $true
    Assert-RecoveryCondition (Test-Path $database -PathType Leaf) "Installed database is missing: $database"
    $installerHash = Get-AcceptanceFileSha256 -Path $installer
    $applicationHash = Get-AcceptanceFileSha256 -Path $application
    $null = & (Join-Path $PSScriptRoot 'Test-SingleFilePayload.ps1') -ExecutablePath $application
    $null = & (Join-Path $PSScriptRoot 'Test-SingleFilePayload.ps1') -ExecutablePath $probe
    $probeHash = Get-AcceptanceFileSha256 -Path $probe
    Assert-RecoveryCondition ($installerHash -eq $ExpectedInstallerSha256.ToUpperInvariant()) 'Installer SHA-256 differs from the frozen candidate.'
    Assert-RecoveryCondition ($applicationHash -eq $ExpectedApplicationSha256.ToUpperInvariant()) 'Installed executable SHA-256 differs from the frozen candidate.'
    Start-InstalledApplication
    Set-DefaultReportDirectory -Path $generatedReports

    Invoke-RecoveryScenario 'compatible-exclusive-lock-recovery' {
        $path = Register-ScenarioWorkbook -ScenarioId $activeScenario.id
        $compatible = [System.IO.FileStream]::new($path,[System.IO.FileMode]::Open,[System.IO.FileAccess]::Read,([System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete))
        try {
            $compatibleStarted = [DateTime]::UtcNow
            Set-WorkbookValue -Path $path -Value 'compatible-lock-save'
            $hash1 = Get-AcceptanceFileSha256 -Path $path
            $p1 = Add-ProbePhase -Name 'compatible-lock-captured' -WorkbookPath $path -Arguments @('--expected-sequence','1','--expected-version-count','1','--expected-error-count','0','--require-active','--require-no-last-error','--require-unique-version-hashes','--require-source-hash-match','--require-ready-report') -TimeoutSeconds $ShortLockSeconds -ExpectedSourceSha256 $hash1 -ExpectedBeforeValue 'baseline' -ExpectedAfterValue 'compatible-lock-save'
            $remainingMilliseconds = [math]::Ceiling(($ShortLockSeconds - ([DateTime]::UtcNow - $compatibleStarted).TotalSeconds) * 1000)
            if ($remainingMilliseconds -gt 0) { Start-Sleep -Milliseconds $remainingMilliseconds }
            $compatibleFinished = [DateTime]::UtcNow
            $null = $p1
        }
        finally { $compatible.Dispose() }
        Add-ManualPhase -Name 'compatible-lock-window' -Detail ([pscustomobject]@{ startedUtc=$compatibleStarted.ToString('O'); releasedUtc=$compatibleFinished.ToString('O'); heldSeconds=[math]::Round(($compatibleFinished-$compatibleStarted).TotalSeconds,3) })
        Set-WorkbookValue -Path $path -Value 'exclusive-lock-save'
        $hash2 = Get-AcceptanceFileSha256 -Path $path
        $exclusive = Open-ExclusiveWorkbookLock -Path $path
        $exclusiveStarted = [DateTime]::UtcNow
        try {
            Start-Sleep -Seconds $ShortLockSeconds
            $null = Add-ProbePhase -Name 'exclusive-lock-held' -WorkbookPath $path -Arguments @('--expected-sequence','1','--expected-version-count','1','--expected-error-count','0','--expected-status','Processing','--require-no-last-error','--require-unique-version-hashes') -TimeoutSeconds 5 -ExpectedSourceSha256 $hash2 -SourceInaccessible -ExpectedBeforeValue 'baseline' -ExpectedAfterValue 'compatible-lock-save'
        }
        finally { $exclusive.Dispose(); $exclusiveFinished = [DateTime]::UtcNow }
        Add-ManualPhase -Name 'exclusive-lock-window' -Detail ([pscustomobject]@{ startedUtc=$exclusiveStarted.ToString('O'); releasedUtc=$exclusiveFinished.ToString('O'); heldSeconds=[math]::Round(($exclusiveFinished-$exclusiveStarted).TotalSeconds,3) })
        $null = Add-ProbePhase -Name 'exclusive-lock-recovered' -WorkbookPath $path -Arguments @('--expected-sequence','2','--expected-version-count','2','--expected-error-count','0','--require-active','--require-no-last-error','--require-unique-version-hashes','--require-source-hash-match','--require-ready-report') -TimeoutSeconds 30 -ExpectedSourceSha256 $hash2 -ExpectedBeforeValue 'compatible-lock-save' -ExpectedAfterValue 'exclusive-lock-save'
    }

    Invoke-RecoveryScenario 'atomic-replacement' {
        $path = Register-ScenarioWorkbook -ScenarioId $activeScenario.id
        $replacement = Join-Path $fixtures 'atomic-replacement-new.xlsx'
        $backup = Join-Path $fixtures 'atomic-replacement-old.xlsx'
        New-WorkbookFixture -Path $replacement -Value 'atomic-replacement-value'
        $hash = Get-AcceptanceFileSha256 -Path $replacement
        Replace-WorkbookAtomically -Replacement $replacement -Target $path -Backup $backup
        $null = Add-ProbePhase -Name 'replacement-captured' -WorkbookPath $path -Arguments @('--expected-sequence','1','--expected-version-count','1','--expected-error-count','0','--require-active','--require-no-last-error','--require-unique-version-hashes','--require-source-hash-match','--require-ready-report') -ExpectedSourceSha256 $hash -ExpectedBeforeValue 'baseline' -ExpectedAfterValue 'atomic-replacement-value'
    }

    Invoke-RecoveryScenario 'save-as-path-behavior' {
        $path = Register-ScenarioWorkbook -ScenarioId $activeScenario.id
        $saveAs = Join-Path $fixtures 'save-as-destination.xlsx'
        $originalHash = Get-AcceptanceFileSha256 -Path $path
        Save-WorkbookAs -Source $path -Destination $saveAs -Value 'save-as-only'
        Set-WorkbookValue -Path $saveAs -Value 'save-as-destination-second-write'
        Start-Sleep -Seconds 12
        $p = Add-ProbePhase -Name 'original-path-unchanged' -WorkbookPath $path -Arguments @('--expected-sequence','0','--expected-version-count','0','--expected-error-count','0','--require-active','--require-no-last-error','--require-unique-version-hashes','--require-source-hash-match') -TimeoutSeconds 5 -ExpectedSourceSha256 $originalHash
        $destinationCopy = Join-Path (Join-Path $workbookDirectory $activeScenario.id) 'save-as-destination.xlsx'
        Copy-SharedFile -Source $saveAs -Destination $destinationCopy
        Add-ManualPhase -Name 'save-as-destination-retained' -Detail ([pscustomobject]@{ destination = Get-FileRecord $destinationCopy; trackedOriginalHash = $p.currentHash })
    }

    Invoke-RecoveryScenario 'autosave-write-burst-dedup-order' {
        $path = Register-ScenarioWorkbook -ScenarioId $activeScenario.id
        $states = [System.Collections.Generic.List[object]]::new()
        for ($index = 1; $index -le 5; $index++) {
            $statePath = Join-Path $fixtures ("burst-state-$index.xlsx")
            New-WorkbookFixture -Path $statePath -Value ("burst-$index")
            $states.Add([pscustomobject]@{ order = $index; path = $statePath; sha256 = Get-AcceptanceFileSha256 -Path $statePath })
        }
        foreach ($state in $states) {
            $incoming = Join-Path $fixtures ("burst-incoming-$($state.order).xlsx")
            Copy-Item -LiteralPath $state.path -Destination $incoming
            Replace-WorkbookAtomically -Replacement $incoming -Target $path -Backup (Join-Path $fixtures ("burst-backup-$($state.order).xlsx"))
            Start-Sleep -Milliseconds 10
        }
        $finalHash = $states[$states.Count - 1].sha256
        Add-ManualPhase -Name 'ordered-write-manifest' -Detail ([pscustomobject]@{ transport = 'AutoSave-equivalent ordered atomic write burst'; writes = @($states | ForEach-Object { $record = Get-FileRecord $_.path; [pscustomobject]@{ order=$_.order; path=$record.path; sha256=$record.sha256; bytes=$record.bytes } }) })
        $null = Add-ProbePhase -Name 'burst-deduplicated' -WorkbookPath $path -Arguments @('--expected-sequence','1','--expected-version-count','1','--expected-error-count','0','--require-active','--require-no-last-error','--require-unique-version-hashes','--require-source-hash-match','--require-ready-report') -TimeoutSeconds 30 -ExpectedSourceSha256 $finalHash -ExpectedBeforeValue 'baseline' -ExpectedAfterValue 'burst-5'
    }

    Invoke-RecoveryScenario 'two-workbook-isolation' {
        $pathA = Register-ScenarioWorkbook -ScenarioId ($activeScenario.id + '-a') -BaselineName 'baseline-a'
        $pathB = Register-ScenarioWorkbook -ScenarioId ($activeScenario.id + '-b') -BaselineName 'baseline-b'
        $replacementA = Join-Path $fixtures 'isolation-a-new.xlsx'
        $replacementB = Join-Path $fixtures 'isolation-b-new.xlsx'
        New-WorkbookFixture -Path $replacementA -Value 'locked-a'
        New-WorkbookFixture -Path $replacementB -Value 'free-b'
        $hashA = Get-AcceptanceFileSha256 -Path $replacementA
        $hashB = Get-AcceptanceFileSha256 -Path $replacementB
        Replace-WorkbookAtomically -Replacement $replacementA -Target $pathA -Backup (Join-Path $fixtures 'isolation-a-old.xlsx')
        $lock = Open-ExclusiveWorkbookLock -Path $pathA
        try {
            Replace-WorkbookAtomically -Replacement $replacementB -Target $pathB -Backup (Join-Path $fixtures 'isolation-b-old.xlsx')
            $null = Add-ProbePhase -Name 'unlocked-b-captured' -WorkbookPath $pathB -Arguments @('--expected-sequence','1','--expected-version-count','1','--expected-error-count','0','--require-active','--require-no-last-error','--require-unique-version-hashes','--require-source-hash-match','--require-ready-report') -TimeoutSeconds 20 -ExpectedSourceSha256 $hashB -ExpectedBeforeValue 'baseline' -ExpectedAfterValue 'free-b'
            $null = Add-ProbePhase -Name 'locked-a-not-advanced' -WorkbookPath $pathA -Arguments @('--expected-sequence','0','--expected-version-count','0','--expected-error-count','0','--expected-status','Processing','--require-no-last-error','--require-unique-version-hashes') -TimeoutSeconds 5 -ExpectedSourceSha256 $hashA -SourceInaccessible
        }
        finally { $lock.Dispose() }
        $null = Add-ProbePhase -Name 'locked-a-recovered' -WorkbookPath $pathA -Arguments @('--expected-sequence','1','--expected-version-count','1','--expected-error-count','0','--require-active','--require-no-last-error','--require-unique-version-hashes','--require-source-hash-match','--require-ready-report') -TimeoutSeconds 30 -ExpectedSourceSha256 $hashA -ExpectedBeforeValue 'baseline' -ExpectedAfterValue 'locked-a'
    }

    Invoke-RecoveryScenario 'missing-file-restoration' {
        $path = Register-ScenarioWorkbook -ScenarioId $activeScenario.id
        $missingCopy = Join-Path $fixtures 'missing-file-held.xlsx'
        Move-Item -LiteralPath $path -Destination $missingCopy
        $missing = Add-ProbePhase -Name 'missing-detected' -WorkbookPath $path -Arguments @('--expected-sequence','0','--expected-version-count','0','--expected-error-count','1','--expected-status','Missing','--require-unique-version-hashes') -TimeoutSeconds ($FailureExposureSeconds + 15)
        $missingPhase = @($activeScenario.phases | Where-Object { $_.name -eq 'missing-detected' })
        Assert-RecoveryCondition ($missingPhase.Count -eq 1 -and -not $missingPhase[0].source.exists) 'The missing phase unexpectedly found the source workbook.'
        Set-WorkbookValue -Path $missingCopy -Value 'restored-after-missing'
        $hash = Get-AcceptanceFileSha256 -Path $missingCopy
        Move-Item -LiteralPath $missingCopy -Destination $path
        $null = Add-ProbePhase -Name 'restoration-captured' -WorkbookPath $path -Arguments @('--expected-sequence','1','--expected-version-count','1','--expected-error-count','1','--require-active','--require-no-last-error','--require-unique-version-hashes','--require-source-hash-match','--require-ready-report') -TimeoutSeconds 30 -ExpectedSourceSha256 $hash -ExpectedBeforeValue 'baseline' -ExpectedAfterValue 'restored-after-missing'
    }

    Invoke-RecoveryScenario 'stopped-app-restart-reconciliation' {
        $path = Register-ScenarioWorkbook -ScenarioId $activeScenario.id
        $replacement = Join-Path $fixtures 'stopped-app-new.xlsx'
        New-WorkbookFixture -Path $replacement -Value 'changed-while-stopped'
        $hash = Get-AcceptanceFileSha256 -Path $replacement
        Stop-InstalledApplication
        $databaseBefore = Get-AcceptanceFileSha256 -Path $database
        Replace-WorkbookAtomically -Replacement $replacement -Target $path -Backup (Join-Path $fixtures 'stopped-app-old.xlsx')
        Start-Sleep -Seconds 2
        $databaseAfter = Get-AcceptanceFileSha256 -Path $database
        Assert-RecoveryCondition ($databaseBefore -eq $databaseAfter) 'The database changed while the installed application was stopped.'
        $stoppedSource = Get-FileRecord $path
        Add-ManualPhase -Name 'change-while-stopped' -Detail ([pscustomobject]@{ source=$stoppedSource; sourceSha256=$hash; databaseSha256Before=$databaseBefore; databaseSha256After=$databaseAfter })
        Restart-InstalledApplication
        $null = Add-ProbePhase -Name 'restart-reconciled' -WorkbookPath $path -Arguments @('--expected-sequence','1','--expected-version-count','1','--expected-error-count','0','--require-active','--require-no-last-error','--require-unique-version-hashes','--require-source-hash-match','--require-ready-report') -TimeoutSeconds 30 -ExpectedSourceSha256 $hash -ExpectedBeforeValue 'baseline' -ExpectedAfterValue 'changed-while-stopped'
    }

    Invoke-RecoveryScenario 'unwritable-report-recovery' {
        $blocked = Join-Path $evidence 'blocked-report-target'
        New-Item -ItemType Directory -Path $blocked -Force | Out-Null
        Set-DefaultReportDirectory -Path $blocked
        try { $path = Register-ScenarioWorkbook -ScenarioId $activeScenario.id -ExpectedReportRoot $blocked }
        finally { Set-DefaultReportDirectory -Path $generatedReports }
        $blockedHolding = Join-Path $evidence 'blocked-report-target-holding'
        Move-Item -LiteralPath $blocked -Destination $blockedHolding
        [System.IO.File]::WriteAllText($blocked,'This file deliberately makes the configured report target unwritable as a directory.')
        Set-WorkbookValue -Path $path -Value 'pending-report'
        $hash = Get-AcceptanceFileSha256 -Path $path
        $pending = Add-ProbePhase -Name 'report-pending' -WorkbookPath $path -Arguments @('--expected-sequence','1','--expected-version-count','1','--expected-error-count','1','--expected-status','Warning','--require-unique-version-hashes','--require-source-hash-match') -TimeoutSeconds 30 -ExpectedSourceSha256 $hash -ExpectedBeforeValue 'baseline' -ExpectedAfterValue 'pending-report'
        Assert-RecoveryCondition ($pending.latestVersion.reportStatus -eq 'Pending') 'The unwritable target did not leave the committed report Pending.'
        $obstructionCopy = Join-Path $logDirectory 'report-target-obstruction.txt'
        Copy-SharedFile -Source $blocked -Destination $obstructionCopy
        Add-ManualPhase -Name 'report-target-obstruction' -Detail ([pscustomobject]@{ target=$blocked; obstruction=Get-FileRecord $obstructionCopy; reportStatus=$pending.latestVersion.reportStatus })
        Remove-Item -LiteralPath $blocked -Force
        Move-Item -LiteralPath $blockedHolding -Destination $blocked
        $ready = Add-ProbePhase -Name 'pending-report-recovered' -WorkbookPath $path -Arguments @('--expected-sequence','1','--expected-version-count','1','--expected-error-count','1','--require-active','--require-no-last-error','--require-unique-version-hashes','--require-source-hash-match','--require-ready-report') -TimeoutSeconds 30 -ExpectedSourceSha256 $hash -ExpectedBeforeValue 'baseline' -ExpectedAfterValue 'pending-report'
        Assert-RecoveryCondition ($ready.latestVersion.reportStatus -eq 'Ready') 'The pending report was not completed after target recovery.'
    }

    Invoke-RecoveryScenario 'interrupted-slow-large-capture-recovery' {
        $path = Register-ScenarioWorkbook -ScenarioId $activeScenario.id -Value 'large-baseline' -Rows $LargeWorkbookRows
        $replacement = Join-Path $fixtures 'large-replacement.xlsx'
        Copy-Item -LiteralPath $path -Destination $replacement
        Set-WorkbookValue -Path $replacement -Value 'large-recovery-change' -Address "T$LargeWorkbookRows"
        $hash = Get-AcceptanceFileSha256 -Path $replacement
        Replace-WorkbookAtomically -Replacement $replacement -Target $path -Backup (Join-Path $fixtures 'large-old.xlsx')
        $lock = Open-ExclusiveWorkbookLock -Path $path
        try {
            $failedLarge = Add-ProbePhase -Name 'large-capture-interrupted' -WorkbookPath $path -Arguments @('--expected-sequence','0','--expected-version-count','0','--expected-error-count','1','--expected-status','Warning','--require-unique-version-hashes') -TimeoutSeconds ($FailureExposureSeconds + 15) -ExpectedSourceSha256 $hash -SourceInaccessible
            Assert-RecoveryCondition ($failedLarge.currentSequence -eq 0) 'The interrupted large capture advanced the baseline.'
        }
        finally { $lock.Dispose() }
        $null = Add-ProbePhase -Name 'large-capture-recovered' -WorkbookPath $path -Arguments @('--expected-sequence','1','--expected-version-count','1','--expected-error-count','1','--require-active','--require-no-last-error','--require-unique-version-hashes','--require-source-hash-match','--require-ready-report') -TimeoutSeconds 60 -ExpectedSourceSha256 $hash -ExpectedAddress "T$LargeWorkbookRows" -ExpectedBeforeValue "large-baseline-$($LargeWorkbookRows - 1)-19" -ExpectedAfterValue 'large-recovery-change'
    }

    Invoke-RecoveryScenario 'corrupt-package-rejection' {
        $path = Register-ScenarioWorkbook -ScenarioId $activeScenario.id
        $baselineCopy = Join-Path $fixtures 'corrupt-valid-baseline.xlsx'
        Copy-Item -LiteralPath $path -Destination $baselineCopy
        $baselineHash = Get-AcceptanceFileSha256 -Path $baselineCopy
        [System.IO.File]::WriteAllBytes($path,[Text.Encoding]::UTF8.GetBytes('not-an-openxml-package'))
        $corruptHash = Get-AcceptanceFileSha256 -Path $path
        $rejected = Add-ProbePhase -Name 'corrupt-rejected' -WorkbookPath $path -Arguments @('--expected-sequence','0','--expected-version-count','0','--expected-error-count','1','--expected-status','Warning','--require-unique-version-hashes') -TimeoutSeconds 30 -ExpectedSourceSha256 $corruptHash
        Assert-RecoveryCondition ($rejected.currentHash -eq $baselineHash) 'Corrupt package rejection advanced the tracked baseline hash.'
        Copy-Item -LiteralPath $baselineCopy -Destination $path -Force
        $null = Add-ProbePhase -Name 'valid-package-restored' -WorkbookPath $path -Arguments @('--expected-sequence','0','--expected-version-count','0','--expected-error-count','1','--require-active','--require-no-last-error','--require-unique-version-hashes','--require-source-hash-match') -TimeoutSeconds 30 -ExpectedSourceSha256 $baselineHash
    }

    Invoke-RecoveryScenario 'encrypted-package-rejection' {
        $path = Register-ScenarioWorkbook -ScenarioId $activeScenario.id
        $baselineCopy = Join-Path $fixtures 'encrypted-valid-baseline.xlsx'
        Copy-Item -LiteralPath $path -Destination $baselineCopy
        $baselineHash = Get-AcceptanceFileSha256 -Path $baselineCopy
        $encrypted = Join-Path $fixtures 'password-protected.xlsx'
        $fixtureMethod = New-PasswordProtectedFixture -Path $encrypted
        $encryptedHash = Get-AcceptanceFileSha256 -Path $encrypted
        Copy-Item -LiteralPath $encrypted -Destination $path -Force
        $rejected = Add-ProbePhase -Name 'encrypted-rejected' -WorkbookPath $path -Arguments @('--expected-sequence','0','--expected-version-count','0','--expected-error-count','1','--expected-status','Warning','--require-unique-version-hashes') -TimeoutSeconds 30 -ExpectedSourceSha256 $encryptedHash
        Assert-RecoveryCondition ($rejected.currentHash -eq $baselineHash) 'Encrypted package rejection advanced the tracked baseline hash.'
        Add-ManualPhase -Name 'encrypted-fixture-provenance' -Detail ([pscustomobject]@{ method=$fixtureMethod; fixture=Get-FileRecord $encrypted; sha256=$encryptedHash; suppliedFixture=if ($PasswordProtectedFixture) { (Resolve-Path $PasswordProtectedFixture).Path } else { $null } })
        Copy-Item -LiteralPath $baselineCopy -Destination $path -Force
        $null = Add-ProbePhase -Name 'unencrypted-package-restored' -WorkbookPath $path -Arguments @('--expected-sequence','0','--expected-version-count','0','--expected-error-count','1','--require-active','--require-no-last-error','--require-unique-version-hashes','--require-source-hash-match') -TimeoutSeconds 30 -ExpectedSourceSha256 $baselineHash
    }
}
catch {
    $fatalFailure = $_.Exception.ToString()
}
finally {
    if ($openWorkbook) { try { $openWorkbook.Close($false) } catch { } }
    if ($excel) { try { $excel.Quit() } catch { } }
    if ($transcriptStarted) { try { Stop-Transcript | Out-Null } catch { } }
    $finishedUtc = [DateTime]::UtcNow
    $allPassed = ($null -eq $fatalFailure -and $scenarios.Count -eq 11 -and @($scenarios | Where-Object { -not $_.passed }).Count -eq 0)
    $result = [pscustomobject]@{
        schemaVersion = 2
        evidenceId = $evidenceId
        outerRunEvidenceId = $OuterRunEvidenceId.ToLowerInvariant()
        startedUtc = $startedUtc.ToString('O')
        finishedUtc = $finishedUtc.ToString('O')
        durationSeconds = [math]::Round(($finishedUtc - $startedUtc).TotalSeconds,3)
        success = $allPassed
        failure = $fatalFailure
        confirmation = [pscustomobject]@{ disposableCleanVm = [bool]$ConfirmDisposableCleanVm; freshEvidenceDirectory = $true; localEvidenceDirectory = -not $evidence.StartsWith('\\') }
        environment = [pscustomobject]@{ machineName=$env:COMPUTERNAME; userName=$env:USERNAME; powershellVersion=$PSVersionTable.PSVersion.ToString(); powershellEdition=$PSVersionTable.PSEdition; osVersion=[Environment]::OSVersion.VersionString }
        candidate = [pscustomobject]@{ installerPath=$installer; installerSha256=if (Test-Path $installer) { Get-AcceptanceFileSha256 -Path $installer } else { $null }; expectedInstallerSha256=$ExpectedInstallerSha256.ToUpperInvariant(); applicationPath=$application; applicationSha256=if (Test-Path $application) { Get-AcceptanceFileSha256 -Path $application } else { $null }; expectedApplicationSha256=$ExpectedApplicationSha256.ToUpperInvariant(); probePath=$probe; probeSha256=if (Test-Path $probe) { Get-AcceptanceFileSha256 -Path $probe } else { $null } }
        thresholds = [pscustomobject]@{ shortLockSeconds=$ShortLockSeconds; failureExposureSeconds=$FailureExposureSeconds; largeWorkbookRows=$LargeWorkbookRows; largeWorkbookColumns=20; largeWorkbookCells=($LargeWorkbookRows*20); reconciliationSeconds=10; stableCopyTimeoutSeconds=60 }
        scenarios = @($scenarios)
        transcript = if (Test-Path $transcriptPath) { Get-FileRecord $transcriptPath } else { $null }
    }
    Write-AcceptanceUtf8File -Path $resultPath -Content ($result | ConvertTo-Json -Depth 30)
}

if (-not $result.success) { throw "Installed recovery matrix failed closed. See $resultPath. $($result.failure)" }
Write-Output "INSTALLED_RECOVERY_MATRIX_PASS|result=$resultPath|evidenceId=$evidenceId"
