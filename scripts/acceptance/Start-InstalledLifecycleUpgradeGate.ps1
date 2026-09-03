[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string] $PriorInstallerPath,
    [Parameter(Mandatory)] [ValidatePattern('^[A-Fa-f0-9]{64}$')] [string] $ExpectedPriorInstallerSha256,
    [Parameter(Mandatory)] [ValidatePattern('^[A-Fa-f0-9]{64}$')] [string] $ExpectedPriorApplicationSha256,
    [Parameter(Mandatory)] [ValidatePattern('^\d+\.\d+\.\d+$')] [string] $PriorVersion,
    [Parameter(Mandatory)] [string] $CandidateInstallerPath,
    [Parameter(Mandatory)] [ValidatePattern('^[A-Fa-f0-9]{64}$')] [string] $ExpectedCandidateInstallerSha256,
    [Parameter(Mandatory)] [ValidatePattern('^[A-Fa-f0-9]{64}$')] [string] $ExpectedCandidateApplicationSha256,
    [Parameter(Mandatory)] [ValidatePattern('^\d+\.\d+\.\d+$')] [string] $CandidateVersion,
    [Parameter(Mandatory)] [string] $ProbePath,
    [Parameter(Mandatory)] [ValidatePattern('^[A-Fa-f0-9]{64}$')] [string] $ExpectedProbeSha256,
    [Parameter(Mandatory)] [string] $EvidenceDirectory,
    [Parameter(Mandatory)] [string] $VmSnapshotName,
    [Parameter(Mandatory)] [string] $VmSnapshotId,
    [string] $ApplicationPath = (Join-Path $env:LOCALAPPDATA 'Programs\Excel Diff Tracker\ExcelDiffTracker.exe'),
    [string] $DatabasePath = (Join-Path $env:LOCALAPPDATA 'Excel Diff Tracker\history.db'),
    [switch] $ConfirmDisposableCleanVm
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

if (-not $ConfirmDisposableCleanVm) {
    throw 'This gate installs and upgrades the product. Run only in a disposable clean VM and pass -ConfirmDisposableCleanVm.'
}
if ($PSVersionTable.PSVersion.Major -ne 5 -or $PSVersionTable.PSVersion.Minor -ne 1 -or $PSVersionTable.PSEdition -ne 'Desktop') {
    throw "Stock Windows PowerShell 5.1 Desktop is required; found $($PSVersionTable.PSVersion) $($PSVersionTable.PSEdition)."
}
$isAdministrator = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if ($isAdministrator) { throw 'The lifecycle/upgrade gate must run as a standard non-administrator user.' }
if ([string]::IsNullOrWhiteSpace($VmSnapshotName) -or [string]::IsNullOrWhiteSpace($VmSnapshotId)) {
    throw 'A named disposable VM snapshot and non-empty snapshot ID are required.'
}
if ([version]$CandidateVersion -le [version]$PriorVersion) {
    throw "CandidateVersion must be newer than PriorVersion; got $PriorVersion -> $CandidateVersion. A same-version repair is not an upgrade."
}

$priorInstaller = (Resolve-Path $PriorInstallerPath).Path
$candidateInstaller = (Resolve-Path $CandidateInstallerPath).Path
$probe = (Resolve-Path $ProbePath).Path
$application = [System.IO.Path]::GetFullPath($ApplicationPath)
$database = [System.IO.Path]::GetFullPath($DatabasePath)
$installDirectory = Split-Path -Parent $application
$dataDirectory = Split-Path -Parent $database
$evidence = [System.IO.Path]::GetFullPath($EvidenceDirectory)
$startMenuShortcut = Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs\Excel Diff Tracker\Excel Diff Tracker.lnk'
$desktopShortcut = Join-Path ([Environment]::GetFolderPath('Desktop')) 'Excel Diff Tracker.lnk'
$runKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'
$defaultReportDirectory = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'Excel Diff Tracker Reports'

foreach ($path in @($priorInstaller,$candidateInstaller,$probe,$application,$database,$evidence)) {
    if ($path.StartsWith('\\')) { throw "All inputs and evidence must use local VM paths, not UNC paths: $path" }
}
if ([string]::Equals($priorInstaller,$candidateInstaller,[StringComparison]::OrdinalIgnoreCase)) {
    throw 'PriorInstallerPath and CandidateInstallerPath must identify different files.'
}
if (Test-Path $evidence) { throw "Use a fresh evidence directory; it already exists: $evidence" }

Import-Module (Join-Path $PSScriptRoot 'UiAutomation.psm1') -Force

$screenshots = Join-Path $evidence 'screenshots'
$uiaDirectory = Join-Path $evidence 'uia'
$probeDirectory = Join-Path $evidence 'probe'
$binaryDirectory = Join-Path $evidence 'binaries'
$databaseDirectory = Join-Path $evidence 'database'
$workbookDirectory = Join-Path $evidence 'workbooks'
$logDirectory = Join-Path $evidence 'logs'
$preResultPath = Join-Path $evidence 'pre-logoff.json'
$pendingPath = Join-Path $evidence 'pending-logoff.json'
$failurePath = Join-Path $evidence 'pre-logoff.failure.json'
$actionLogPath = Join-Path $logDirectory 'pre-logoff-actions.txt'
$workbookPath = Join-Path $workbookDirectory 'Installed Lifecycle Upgrade.xlsx'
New-Item -ItemType Directory -Path $evidence,$screenshots,$uiaDirectory,$probeDirectory,$binaryDirectory,$databaseDirectory,$workbookDirectory,$logDirectory -Force | Out-Null

$startedUtc = [DateTime]::UtcNow
$evidenceId = [Guid]::NewGuid().ToString('D')
$phaseLinkNonce = [Guid]::NewGuid().ToString('D')
$applicationProcess = $null
$excel = $null
$excelWorkbook = $null
$failure = $null
$priorBaselineProbe = $null
$priorProbe = $null
$candidatePreservedProbe = $null
$candidateSaveProbe = $null
$preDatabaseRecord = $null
$postInstallDatabaseRecord = $null
$finalDatabaseRecords = @()
$uiRecords = [System.Collections.Generic.List[object]]::new()
$checks = [ordered]@{
    disposableCleanVmConfirmed = $false
    cleanNoInstallNoDataState = $false
    frozenInputHashesExact = $false
    installerVersionIdentitiesExact = $false
    priorInstalledPerUser = $false
    priorStartMenuLaunchUsed = $false
    priorOnboardingCompletedWithRealXlsx = $false
    priorVisibleExcelKeyboardSaveExact = $false
    priorExitedThroughTray = $false
    candidateInstalledOverPriorWithoutUninstall = $false
    installedCandidateHashExact = $false
    candidateHistoryExactlyPreserved = $false
    candidateOnboardingDidNotRepeat = $false
    candidateVisibleExcelKeyboardSaveExact = $false
    startupRegistrationExact = $false
    closeLeavesCandidateInActualTray = $false
    pendingStateWritten = $false
    externalLogoffStillRequired = $false
}

function Assert-GateCondition {
    param([bool] $Condition,[string] $Message)
    if (-not $Condition) { throw $Message }
}

function Write-ActionLog {
    param([string] $Message)
    $line = '{0}|{1}' -f [DateTime]::UtcNow.ToString('O'),$Message
    [System.IO.File]::AppendAllText($actionLogPath,$line + [Environment]::NewLine,[System.Text.UTF8Encoding]::new($false))
}

function Get-Hash {
    param([string] $Path)
    (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToUpperInvariant()
}

function Get-Relative {
    param([string] $Path)
    Get-AcceptanceRelativePath -BasePath $evidence -Path $Path -UseForwardSlash
}

function Get-FileRecord {
    param([string] $Path)
    $item = Get-Item -LiteralPath $Path
    [pscustomobject]@{ path=Get-Relative $item.FullName; sha256=Get-Hash $item.FullName; bytes=[long]$item.Length }
}

function Get-ExternalFileIdentity {
    param([string] $Path)
    $item = Get-Item -LiteralPath $Path
    [pscustomobject]@{ path=$item.FullName; sha256=Get-Hash $item.FullName; bytes=[long]$item.Length }
}

function Copy-SharedFile {
    param([string] $Source,[string] $Destination)
    $inputStream = [System.IO.FileStream]::new($Source,[System.IO.FileMode]::Open,[System.IO.FileAccess]::Read,([System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete))
    try {
        $outputStream = [System.IO.FileStream]::new($Destination,[System.IO.FileMode]::CreateNew,[System.IO.FileAccess]::Write,[System.IO.FileShare]::None)
        try { $inputStream.CopyTo($outputStream); $outputStream.Flush() } finally { $outputStream.Dispose() }
    } finally { $inputStream.Dispose() }
}

function Copy-DatabaseEvidence {
    param([string] $Name)
    $destination = Join-Path $databaseDirectory $Name
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
    Assert-GateCondition ($records.Count -ge 1) "No database bytes were retained at checkpoint $Name."
    @($records)
}

function Get-LogonIdentity {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $logonSids = @($identity.Groups | ForEach-Object { $_.Value } | Where-Object { $_ -match '^S-1-5-5-\d+-\d+$' })
    Assert-GateCondition ($logonSids.Count -eq 1) "Expected exactly one token logon SID; found $($logonSids.Count)."
    $sessionId = [System.Diagnostics.Process]::GetCurrentProcess().SessionId
    Assert-GateCondition ($sessionId -gt 0) 'The gate must run in an interactive Windows desktop session.'
    $explorers = @(Get-Process -Name explorer -ErrorAction SilentlyContinue | Where-Object { $_.SessionId -eq $sessionId })
    Assert-GateCondition ($explorers.Count -eq 1) "Expected exactly one Explorer shell in interactive session $sessionId; found $($explorers.Count)."
    [pscustomobject]@{
        accountName = $identity.Name
        accountSid = $identity.User.Value
        logonSid = $logonSids[0]
        windowsSessionId = $sessionId
        explorerProcessId = $explorers[0].Id
        explorerStartedUtc = $explorers[0].StartTime.ToUniversalTime().ToString('O')
        capturedUtc = [DateTime]::UtcNow.ToString('O')
    }
}

function Get-InstallerProductVersion {
    param([string] $Path)
    ([System.Diagnostics.FileVersionInfo]::GetVersionInfo($Path)).ProductVersion
}

function Get-ProductStartupRegistration {
    if (-not (Test-Path -LiteralPath $runKey -ErrorAction Stop)) {
        return [pscustomobject]@{ exists=$false; value=$null }
    }
    $properties = Get-ItemProperty -LiteralPath $runKey -ErrorAction Stop
    Assert-GateCondition ($null -ne $properties) 'The startup registry key could not be read.'
    $property = $properties.PSObject.Properties['ExcelDiffTracker']
    if ($null -eq $property) { return [pscustomobject]@{ exists=$false; value=$null } }
    [pscustomobject]@{ exists=$true; value=$property.Value }
}

function Get-ProductUninstallEntry {
    $roots = @(
        'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKCU:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall')
    $matches = [System.Collections.Generic.List[object]]::new()
    foreach ($root in $roots) {
        if (-not (Test-Path $root)) { continue }
        foreach ($key in @(Get-ChildItem -Path $root -ErrorAction SilentlyContinue)) {
            $properties = Get-ItemProperty -LiteralPath $key.PSPath -ErrorAction SilentlyContinue
            if ($properties -and $properties.DisplayName -eq 'Excel Diff Tracker') {
                $matches.Add([pscustomobject]@{
                    registryPath=$key.PSPath
                    displayName=[string]$properties.DisplayName
                    displayVersion=[string]$properties.DisplayVersion
                    installLocation=[string]$properties.InstallLocation
                    uninstallString=[string]$properties.UninstallString
                })
            }
        }
    }
    Assert-GateCondition ($matches.Count -eq 1) "Expected exactly one Excel Diff Tracker uninstall entry; found $($matches.Count)."
    $matches[0]
}

function Test-NoProductUninstallEntry {
    $roots = @(
        'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKCU:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall')
    foreach ($root in $roots) {
        if (-not (Test-Path $root)) { continue }
        foreach ($key in @(Get-ChildItem -Path $root -ErrorAction SilentlyContinue)) {
            $properties = Get-ItemProperty -LiteralPath $key.PSPath -ErrorAction SilentlyContinue
            if ($properties -and $properties.DisplayName -eq 'Excel Diff Tracker') { return $false }
        }
    }
    $true
}

function Save-UiEvidence {
    param([string] $Name,[System.Windows.Automation.AutomationElement] $Root)
    $png = Join-Path $screenshots ($Name + '.png')
    $tree = Join-Path $uiaDirectory ($Name + '.json')
    Save-DesktopScreenshot -Path $png
    Export-UiaTree -Root $Root -Path $tree
    $record = [pscustomobject]@{ name=$Name; screenshot=Get-FileRecord $png; uiaTree=Get-FileRecord $tree; capturedUtc=[DateTime]::UtcNow.ToString('O') }
    $uiRecords.Add($record)
    $record
}

function Get-DesktopWindows {
    ([System.Windows.Automation.AutomationElement]::RootElement).FindAll(
        [System.Windows.Automation.TreeScope]::Children,
        [System.Windows.Automation.Condition]::TrueCondition)
}

function Assert-NoWelcomeWindow {
    param([int] $ProcessId)
    $matches = @(Get-DesktopWindows | Where-Object { $_.Current.ProcessId -eq $ProcessId -and $_.Current.Name -eq 'Welcome to Excel Diff Tracker' })
    Assert-GateCondition ($matches.Count -eq 0) 'First-run onboarding unexpectedly appeared.'
}

function Get-ProductTrayIcon {
    $root = [System.Windows.Automation.AutomationElement]::RootElement
    $find = {
        $items = $root.FindAll([System.Windows.Automation.TreeScope]::Descendants,[System.Windows.Automation.Condition]::TrueCondition)
        @($items | Where-Object { $_.Current.AutomationId -eq 'NotifyItemIcon' -and $_.Current.Name -like 'Excel Diff Tracker*' }) | Select-Object -First 1
    }
    $icon = & $find
    if (-not $icon) {
        $hidden = Find-UiaElement -Root $root -AutomationId 'SystemTrayIcon' -Name 'Show Hidden Icons' -Optional
        if ($hidden) { Invoke-UiaElement -Element $hidden; Start-Sleep -Milliseconds 500; $icon = & $find }
    }
    Assert-GateCondition ($null -ne $icon) 'The actual Excel Diff Tracker notification-area icon was not found.'
    $icon
}

function Test-MainWindowVisible {
    param([int] $ProcessId)
    @(Get-DesktopWindows | Where-Object { $_.Current.ProcessId -eq $ProcessId -and $_.Current.Name -eq 'Excel Diff Tracker' }).Count -eq 1
}

function Exit-ThroughTray {
    param([System.Diagnostics.Process] $Process)
    $icon = Get-ProductTrayIcon
    Invoke-UiaMouseClick -Element $icon -Button Right
    $root = [System.Windows.Automation.AutomationElement]::RootElement
    $items = $root.FindAll([System.Windows.Automation.TreeScope]::Descendants,[System.Windows.Automation.Condition]::TrueCondition)
    $exit = @($items | Where-Object { $_.Current.Name -eq 'Exit' -and $_.Current.ControlType -eq [System.Windows.Automation.ControlType]::MenuItem }) | Select-Object -First 1
    Assert-GateCondition ($null -ne $exit) 'The native tray menu did not expose Exit.'
    Invoke-UiaElement -Element $exit
    Wait-AcceptanceCondition -TimeoutSeconds 15 -FailureMessage 'The product did not exit through its tray Exit command.' -Condition { $Process.Refresh(); $Process.HasExited }
}

function Close-MainToTray {
    param([System.Windows.Automation.AutomationElement] $Window,[System.Diagnostics.Process] $Process)
    $pattern = $null
    Assert-GateCondition ($Window.TryGetCurrentPattern([System.Windows.Automation.WindowPattern]::Pattern,[ref]$pattern)) 'The main window does not expose WindowPattern.'
    ([System.Windows.Automation.WindowPattern]$pattern).Close()
    Wait-AcceptanceCondition -TimeoutSeconds 8 -FailureMessage 'Closing the main window did not hide it.' -Condition { -not (Test-MainWindowVisible -ProcessId $Process.Id) }
    $Process.Refresh()
    Assert-GateCondition (-not $Process.HasExited) 'Closing the main window exited the candidate instead of retaining the tray process.'
    $null = Get-ProductTrayIcon
}

function Select-PriorUiaElement {
    param(
        [object[]] $Elements,
        [int] $ProcessId,
        [string] $Name,
        [System.Windows.Automation.ControlType] $ControlType,
        [switch] $Optional
    )
    $matching = @($Elements | Where-Object {
        $_.Current.ProcessId -eq $ProcessId -and $_.Current.Name -ceq $Name -and
        $_.Current.ControlType -eq $ControlType -and -not $_.Current.IsOffscreen
    })
    if ($Optional -and $matching.Count -eq 0) { return $null }
    Assert-GateCondition ($matching.Count -eq 1) "Expected one visible prior-release $($ControlType.ProgrammaticName) named '$Name' in process $ProcessId; found $($matching.Count)."
    $matching[0]
}

function Find-PriorUiaElement {
    param(
        [System.Windows.Automation.AutomationElement] $Root,
        [int] $ProcessId,
        [string] $Name,
        [System.Windows.Automation.ControlType] $ControlType,
        [switch] $Optional
    )
    Assert-GateCondition ($Root.Current.ProcessId -eq $ProcessId -and $Root.Current.ControlType -eq [System.Windows.Automation.ControlType]::Window) 'The prior-release selector must be scoped to the expected process window.'
    $elements = $Root.FindAll([System.Windows.Automation.TreeScope]::Descendants,[System.Windows.Automation.Condition]::TrueCondition)
    Select-PriorUiaElement -Elements @($elements) -ProcessId $ProcessId -Name $Name -ControlType $ControlType -Optional:$Optional
}

function Choose-WorkbookFromDialog {
    param([string] $Path,[int] $ProcessId)
    $dialog = Find-UiaWindow -Title 'Choose a workbook to track' -ProcessId $ProcessId -TimeoutSeconds 15
    $fileName = Find-UiaElement -Root $dialog -AutomationId '1148' -Optional
    if (-not $fileName) { $fileName = Find-UiaElement -Root $dialog -Name 'File name:' }
    Set-UiaValue -Element $fileName -Value $Path
    $open = Find-UiaElement -Root $dialog -AutomationId '1' -Optional
    if (-not $open) { $open = Find-UiaElement -Root $dialog -Name 'Open' }
    Invoke-UiaElement -Element $open
}

function Complete-PriorOnboarding {
    param([System.Diagnostics.Process] $Process,[string] $Workbook)
    Assert-GateCondition ($PriorVersion -eq '0.1.1') 'The prior-release UI adapter requires the pinned 0.1.1 release.'
    $window = Find-UiaWindow -Title 'Welcome to Excel Diff Tracker' -ProcessId $Process.Id -TimeoutSeconds 20
    $scope = @{ Root=$window; ProcessId=$Process.Id }
    $null = Save-UiEvidence -Name 'prior-onboarding-step-1' -Root $window
    Invoke-UiaElement -Element (Find-PriorUiaElement @scope -Name 'Continue' -ControlType ([System.Windows.Automation.ControlType]::Button))
    Start-Sleep -Milliseconds 250
    $null = Find-PriorUiaElement @scope -Name 'Choose a report folder' -ControlType ([System.Windows.Automation.ControlType]::Text)
    Invoke-UiaElement -Element (Find-PriorUiaElement @scope -Name 'Continue' -ControlType ([System.Windows.Automation.ControlType]::Button))
    Start-Sleep -Milliseconds 250
    $null = Find-PriorUiaElement @scope -Name 'Add your first workbook' -ControlType ([System.Windows.Automation.ControlType]::Text)
    Invoke-UiaElement -Element (Find-PriorUiaElement @scope -Name ('Choose workbook' + [char]0x2026) -ControlType ([System.Windows.Automation.ControlType]::Button))
    Choose-WorkbookFromDialog -Path $Workbook -ProcessId $Process.Id
    Start-Sleep -Milliseconds 400
    $null = Find-PriorUiaElement @scope -Name $Workbook -ControlType ([System.Windows.Automation.ControlType]::Text)
    Invoke-UiaElement -Element (Find-PriorUiaElement @scope -Name 'Continue' -ControlType ([System.Windows.Automation.ControlType]::Button))
    Start-Sleep -Milliseconds 250
    foreach ($name in @('Start Excel Diff Tracker with Windows','Begin tracking the selected workbook')) {
        $element = Find-PriorUiaElement @scope -Name $name -ControlType ([System.Windows.Automation.ControlType]::CheckBox)
        $toggle = $null
        Assert-GateCondition ($element.TryGetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern,[ref]$toggle)) "$name does not expose TogglePattern."
        Assert-GateCondition (([System.Windows.Automation.TogglePattern]$toggle).Current.ToggleState -eq [System.Windows.Automation.ToggleState]::On) "$name was not enabled."
    }
    Invoke-UiaElement -Element (Find-PriorUiaElement @scope -Name 'Activate tracking' -ControlType ([System.Windows.Automation.ControlType]::Button))
    Wait-AcceptanceCondition -TimeoutSeconds 60 -FailureMessage 'Onboarding did not create the real-workbook baseline.' -Condition { $null -ne (Find-PriorUiaElement @scope -Name 'Tracking is active' -ControlType ([System.Windows.Automation.ControlType]::Text) -Optional) }
    $null = Save-UiEvidence -Name 'prior-onboarding-complete' -Root $window
    Invoke-UiaElement -Element (Find-PriorUiaElement @scope -Name 'Open dashboard' -ControlType ([System.Windows.Automation.ControlType]::Button))
    Find-UiaWindow -Title 'Excel Diff Tracker' -ProcessId $Process.Id -TimeoutSeconds 20
}

function Start-FromShortcut {
    Assert-GateCondition (Test-Path $startMenuShortcut -PathType Leaf) 'The installed Start-menu shortcut is missing.'
    $before = @(Get-Process -Name ExcelDiffTracker -ErrorAction SilentlyContinue)
    Assert-GateCondition ($before.Count -eq 0) 'The application was running before the Start-menu launch.'
    Start-Process -FilePath $startMenuShortcut | Out-Null
    Wait-AcceptanceCondition -TimeoutSeconds 20 -FailureMessage 'The Start-menu shortcut did not start the application.' -Condition { @(Get-Process -Name ExcelDiffTracker -ErrorAction SilentlyContinue).Count -eq 1 }
    $process = Get-Process -Name ExcelDiffTracker | Select-Object -First 1
    Assert-GateCondition ([string]::Equals($process.Path,$application,[StringComparison]::OrdinalIgnoreCase)) "The Start-menu shortcut launched an unexpected executable: $($process.Path)"
    $process
}

function Invoke-ProbeUntilPassed {
    param([string] $Name,[string[]] $Arguments,[int] $TimeoutSeconds = 30)
    $output = Join-Path $probeDirectory ($Name + '.json')
    $deadline = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
    $last = ''
    do {
        $last = & $probe @Arguments 2>&1 | Out-String
        if ($LASTEXITCODE -eq 0) {
            Write-AcceptanceUtf8File -Path $output -Content $last
            return [pscustomobject]@{ record=Get-FileRecord $output; result=($last | ConvertFrom-Json) }
        }
        Start-Sleep -Milliseconds 500
    } while ([DateTime]::UtcNow -lt $deadline)
    Write-AcceptanceUtf8File -Path $output -Content $last
    throw "AcceptanceProbe did not pass for $Name within $TimeoutSeconds seconds. See $output"
}

function New-RealWorkbook {
    param([string] $Path)
    $app = New-Object -ComObject Excel.Application
    $book = $null
    try {
        $app.Visible = $true
        $app.DisplayAlerts = $false
        $book = $app.Workbooks.Add()
        $sheet = $book.Worksheets.Item(1)
        $sheet.Name = 'Lifecycle'
        $sheet.Range('A1').Value2 = 'Installed lifecycle upgrade baseline'
        $book.SaveAs($Path,51)
        $book.Close($false)
        $book = $null
    } finally {
        if ($book) { try { $book.Close($false) } catch { } }
        $app.Quit()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)
    }
    Assert-GateCondition (Test-Path $Path -PathType Leaf) 'Excel did not create the real .xlsx fixture.'
}

function Save-VisibleExcelKeyboardValue {
    param([string] $Workbook,[string] $Address,[string] $Value,[string] $EvidenceName)
    $script:excel = New-Object -ComObject Excel.Application
    $script:excel.Visible = $true
    $script:excel.DisplayAlerts = $false
    $script:excelWorkbook = $script:excel.Workbooks.Open($Workbook)
    $excelWindow = Get-UiaWindowFromHandle -Handle ([long]$script:excel.Hwnd)
    Set-UiaForeground -Window $excelWindow
    [System.Windows.Forms.SendKeys]::SendWait('^g')
    Start-Sleep -Milliseconds 200
    [System.Windows.Forms.SendKeys]::SendWait($Address)
    [System.Windows.Forms.SendKeys]::SendWait('{ENTER}')
    Start-Sleep -Milliseconds 200
    [System.Windows.Forms.SendKeys]::SendWait($Value)
    [System.Windows.Forms.SendKeys]::SendWait('{ENTER}')
    [System.Windows.Forms.SendKeys]::SendWait('^s')
    Start-Sleep -Milliseconds 800
    $null = Save-UiEvidence -Name $EvidenceName -Root $excelWindow
    $script:excelWorkbook.Close($false)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($script:excelWorkbook)
    $script:excelWorkbook = $null
    $script:excel.Quit()
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($script:excel)
    $script:excel = $null
}

function Assert-DashboardState {
    param([System.Windows.Automation.AutomationElement] $Window,[long] $ExpectedSequence,[switch] $PriorRelease)
    if ($PriorRelease) {
        Invoke-UiaElement -Element (Find-PriorUiaElement -Root $Window -ProcessId $Window.Current.ProcessId -Name ([string][char]0x2302 + '  Dashboard') -ControlType ([System.Windows.Automation.ControlType]::Button))
    } else {
        Invoke-UiaElement -Element (Find-UiaElement -Root $Window -AutomationId 'DashboardNavigationButton')
    }
    Start-Sleep -Milliseconds 300
    $pathElement = if ($PriorRelease) {
        Find-PriorUiaElement -Root $Window -ProcessId $Window.Current.ProcessId -Name $workbookPath -ControlType ([System.Windows.Automation.ControlType]::Text)
    } else {
        Find-UiaElement -Root $Window -Name $workbookPath
    }
    $walker = [System.Windows.Automation.TreeWalker]::ControlViewWalker
    $ancestor = $pathElement
    $proved = $false
    for ($index = 0; $index -lt 8 -and $ancestor; $index++) {
        $nodes = $ancestor.FindAll([System.Windows.Automation.TreeScope]::Descendants,[System.Windows.Automation.Condition]::TrueCondition)
        $names = @($nodes | ForEach-Object { $_.Current.Name })
        if ($names -contains 'VERSIONS' -and $names -contains ([string]$ExpectedSequence)) { $proved = $true; break }
        $ancestor = $walker.GetParent($ancestor)
    }
    Assert-GateCondition $proved "The dashboard did not bind the tracked workbook to exact sequence $ExpectedSequence."
}

function Write-Json {
    param([string] $Path,[object] $Value)
    Write-AcceptanceUtf8File -Path $Path -Content ($Value | ConvertTo-Json -Depth 30)
}

try {
    Write-ActionLog 'PRE_PHASE_STARTED'
    $checks.disposableCleanVmConfirmed = $ConfirmDisposableCleanVm.IsPresent
    Assert-GateCondition (-not (Test-Path $installDirectory)) "Clean gate failed: install directory exists: $installDirectory"
    Assert-GateCondition (-not (Test-Path $dataDirectory)) "Clean gate failed: local data directory exists: $dataDirectory"
    Assert-GateCondition (-not (Test-Path $defaultReportDirectory)) "Clean gate failed: default report directory exists: $defaultReportDirectory"
    Assert-GateCondition (-not (Test-Path $startMenuShortcut)) "Clean gate failed: Start-menu shortcut exists: $startMenuShortcut"
    Assert-GateCondition (-not (Test-Path $desktopShortcut)) "Clean gate failed: desktop shortcut exists: $desktopShortcut"
    Assert-GateCondition (@(Get-Process -Name ExcelDiffTracker -ErrorAction SilentlyContinue).Count -eq 0) 'Clean gate failed: Excel Diff Tracker is running.'
    $existingStartup = Get-ProductStartupRegistration
    Assert-GateCondition (-not $existingStartup.exists) 'Clean gate failed: startup registry value exists.'
    Assert-GateCondition (Test-NoProductUninstallEntry) 'Clean gate failed: an Excel Diff Tracker uninstall entry exists.'
    $checks.cleanNoInstallNoDataState = $true

    $priorInstallerHash = Get-Hash $priorInstaller
    $candidateInstallerHash = Get-Hash $candidateInstaller
    $probeHash = Get-Hash $probe
    $null = & (Join-Path $PSScriptRoot 'Test-SingleFilePayload.ps1') -ExecutablePath $probe
    Assert-GateCondition ($priorInstallerHash -eq $ExpectedPriorInstallerSha256.ToUpperInvariant()) 'Prior public installer hash differs from the frozen value.'
    Assert-GateCondition ($candidateInstallerHash -eq $ExpectedCandidateInstallerSha256.ToUpperInvariant()) 'Candidate installer hash differs from the frozen value.'
    Assert-GateCondition ($probeHash -eq $ExpectedProbeSha256.ToUpperInvariant()) 'AcceptanceProbe hash differs from the frozen value.'
    Assert-GateCondition ($priorInstallerHash -ne $candidateInstallerHash) 'Prior and candidate installer hashes are identical; this would be a repair test.'
    Assert-GateCondition ($ExpectedPriorApplicationSha256.ToUpperInvariant() -ne $ExpectedCandidateApplicationSha256.ToUpperInvariant()) 'Prior and candidate application hashes are identical; an actual payload upgrade is required.'
    $checks.frozenInputHashesExact = $true
    $priorInstallerProductVersion = Get-InstallerProductVersion $priorInstaller
    $candidateInstallerProductVersion = Get-InstallerProductVersion $candidateInstaller
    Assert-GateCondition ($priorInstallerProductVersion -eq $PriorVersion) "Prior installer ProductVersion is '$priorInstallerProductVersion', expected '$PriorVersion'."
    Assert-GateCondition ($candidateInstallerProductVersion -eq $CandidateVersion) "Candidate installer ProductVersion is '$candidateInstallerProductVersion', expected '$CandidateVersion'."
    $checks.installerVersionIdentitiesExact = $true
    $preSession = Get-LogonIdentity

    New-RealWorkbook -Path $workbookPath
    $workbookInitial = Get-FileRecord $workbookPath
    Write-ActionLog 'INSTALL_PRIOR_STARTED'
    $priorInstall = Start-Process -FilePath $priorInstaller -ArgumentList '/VERYSILENT','/CURRENTUSER','/SUPPRESSMSGBOXES','/NORESTART','/TASKS="startup"' -Wait -PassThru
    Assert-GateCondition ($priorInstall.ExitCode -eq 0) "Prior installer exited with code $($priorInstall.ExitCode)."
    Assert-GateCondition (Test-Path $application -PathType Leaf) 'The prior installed executable is missing.'
    $priorApplicationHash = Get-Hash $application
    Assert-GateCondition ($priorApplicationHash -eq $ExpectedPriorApplicationSha256.ToUpperInvariant()) 'Installed prior executable hash differs from the frozen prior payload.'
    $priorEntry = Get-ProductUninstallEntry
    Assert-GateCondition ($priorEntry.displayVersion -eq $PriorVersion) "Installed prior DisplayVersion is '$($priorEntry.displayVersion)', expected '$PriorVersion'."
    Assert-GateCondition ([System.IO.Path]::GetFullPath($priorEntry.installLocation.TrimEnd('\')) -eq $installDirectory.TrimEnd('\')) 'Prior uninstall entry points to an unexpected install directory.'
    Copy-Item -LiteralPath $application -Destination (Join-Path $binaryDirectory 'prior-ExcelDiffTracker.exe')
    $priorApplicationRecord = Get-FileRecord (Join-Path $binaryDirectory 'prior-ExcelDiffTracker.exe')
    $checks.priorInstalledPerUser = $true
    Write-ActionLog 'INSTALL_PRIOR_COMPLETED'

    $applicationProcess = Start-FromShortcut
    $checks.priorStartMenuLaunchUsed = $true
    $priorMain = Complete-PriorOnboarding -Process $applicationProcess -Workbook $workbookPath
    $checks.priorOnboardingCompletedWithRealXlsx = $true
    $priorBaselineProbe = Invoke-ProbeUntilPassed -Name 'prior-baseline' -Arguments @(
        '--database',$database,'--workbook',$workbookPath,'--expected-sequence','0','--expected-version-count','0',
        '--expected-error-count','0','--require-active','--require-no-last-error','--require-unique-version-hashes','--require-source-hash-match') -TimeoutSeconds 60
    Assert-DashboardState -Window $priorMain -ExpectedSequence 0 -PriorRelease
    $null = Save-UiEvidence -Name 'prior-dashboard-baseline' -Root $priorMain

    $preValue = 'EDT-UPGRADE-PRE-' + $phaseLinkNonce.Substring(0,8).ToUpperInvariant()
    $postValue = 'EDT-UPGRADE-CANDIDATE-' + $phaseLinkNonce.Substring(0,8).ToUpperInvariant()
    $address = 'Z1000'
    Save-VisibleExcelKeyboardValue -Workbook $workbookPath -Address $address -Value $preValue -EvidenceName 'prior-visible-excel-ctrl-s'
    $priorProbe = Invoke-ProbeUntilPassed -Name 'prior-sequence-1' -Arguments @(
        '--database',$database,'--workbook',$workbookPath,'--expected-sequence','1','--expected-version-count','1',
        '--expected-error-count','0','--require-active','--require-no-last-error','--require-unique-version-hashes','--require-source-hash-match',
        '--require-ready-report','--expected-cell-change-count','1','--expected-sheet-change-count','0','--address',$address,
        '--expected-value',$preValue,'--expected-kind','LiteralAdded','--expect-before-missing')
    $checks.priorVisibleExcelKeyboardSaveExact = $true
    $null = Save-UiEvidence -Name 'prior-dashboard-sequence-1' -Root $priorMain
    Exit-ThroughTray -Process $applicationProcess
    $checks.priorExitedThroughTray = $true
    $applicationProcess = $null

    $preDatabaseRecord = Get-ExternalFileIdentity $database
    $preUpgradeDatabaseRecords = Copy-DatabaseEvidence -Name 'before-candidate-upgrade'
    $priorEntryBeforeUpgrade = Get-ProductUninstallEntry
    Assert-GateCondition ($priorEntryBeforeUpgrade.displayVersion -eq $PriorVersion) 'The prior version identity was not present immediately before the candidate installer ran.'
    Write-ActionLog 'IN_PLACE_CANDIDATE_INSTALL_STARTED_NO_UNINSTALL_WAS_INVOKED'
    $candidateInstallStartedUtc = [DateTime]::UtcNow
    $candidateInstall = Start-Process -FilePath $candidateInstaller -ArgumentList '/VERYSILENT','/CURRENTUSER','/SUPPRESSMSGBOXES','/NORESTART','/TASKS="startup"' -Wait -PassThru
    $candidateInstallCompletedUtc = [DateTime]::UtcNow
    Assert-GateCondition ($candidateInstall.ExitCode -eq 0) "Candidate installer exited with code $($candidateInstall.ExitCode)."
    $candidateEntry = Get-ProductUninstallEntry
    Assert-GateCondition ($candidateEntry.displayVersion -eq $CandidateVersion) "Installed candidate DisplayVersion is '$($candidateEntry.displayVersion)', expected '$CandidateVersion'."
    Assert-GateCondition ($priorEntryBeforeUpgrade.registryPath -eq $candidateEntry.registryPath) 'Candidate did not replace the same per-user product registration in place.'
    $postInstallDatabaseRecord = Get-ExternalFileIdentity $database
    Assert-GateCondition ($postInstallDatabaseRecord.sha256 -eq $preDatabaseRecord.sha256 -and $postInstallDatabaseRecord.bytes -eq $preDatabaseRecord.bytes) 'The candidate installer changed the closed history database during upgrade.'
    $checks.candidateInstalledOverPriorWithoutUninstall = $true
    $candidateApplicationHash = Get-Hash $application
    $null = & (Join-Path $PSScriptRoot 'Test-SingleFilePayload.ps1') -ExecutablePath $application
    Assert-GateCondition ($candidateApplicationHash -eq $ExpectedCandidateApplicationSha256.ToUpperInvariant()) 'Installed candidate executable hash differs from the frozen candidate payload.'
    Copy-Item -LiteralPath $application -Destination (Join-Path $binaryDirectory 'candidate-ExcelDiffTracker.exe')
    $candidateApplicationRecord = Get-FileRecord (Join-Path $binaryDirectory 'candidate-ExcelDiffTracker.exe')
    $checks.installedCandidateHashExact = $true
    Write-ActionLog 'IN_PLACE_CANDIDATE_INSTALL_COMPLETED'

    $applicationProcess = Start-Process -FilePath $application -PassThru
    $candidateMain = Find-UiaWindow -Title 'Excel Diff Tracker' -ProcessId $applicationProcess.Id -TimeoutSeconds 20
    Assert-NoWelcomeWindow -ProcessId $applicationProcess.Id
    $checks.candidateOnboardingDidNotRepeat = $true
    $candidatePreservedProbe = Invoke-ProbeUntilPassed -Name 'candidate-preserved-sequence-1' -Arguments @(
        '--database',$database,'--workbook',$workbookPath,'--expected-sequence','1','--expected-version-count','1',
        '--expected-error-count','0','--require-active','--require-no-last-error','--require-unique-version-hashes','--require-source-hash-match',
        '--require-ready-report','--expected-cell-change-count','1','--expected-sheet-change-count','0','--address',$address,
        '--expected-value',$preValue,'--expected-kind','LiteralAdded','--expect-before-missing')
    Assert-GateCondition ($candidatePreservedProbe.result.currentHash -eq $priorProbe.result.currentHash) 'Candidate launch did not preserve the exact prior tracked hash.'
    Assert-DashboardState -Window $candidateMain -ExpectedSequence 1
    $null = Save-UiEvidence -Name 'candidate-dashboard-preserved-sequence-1' -Root $candidateMain
    $checks.candidateHistoryExactlyPreserved = $true

    Save-VisibleExcelKeyboardValue -Workbook $workbookPath -Address $address -Value $postValue -EvidenceName 'candidate-visible-excel-ctrl-s'
    $candidateSaveProbe = Invoke-ProbeUntilPassed -Name 'candidate-sequence-2' -Arguments @(
        '--database',$database,'--workbook',$workbookPath,'--expected-sequence','2','--expected-version-count','2',
        '--expected-error-count','0','--require-active','--require-no-last-error','--require-unique-version-hashes','--require-source-hash-match',
        '--require-ready-report','--expected-cell-change-count','1','--expected-sheet-change-count','0','--address',$address,
        '--expected-value',$postValue,'--expected-before-value',$preValue,'--expected-kind','LiteralChanged')
    Assert-DashboardState -Window $candidateMain -ExpectedSequence 2
    $null = Save-UiEvidence -Name 'candidate-dashboard-sequence-2' -Root $candidateMain
    $checks.candidateVisibleExcelKeyboardSaveExact = $true

    $startupRegistration = Get-ProductStartupRegistration
    Assert-GateCondition $startupRegistration.exists 'The candidate startup registry value is missing.'
    $startupValue = $startupRegistration.value
    $expectedStartupValue = '"' + $application + '" --background'
    Assert-GateCondition ([string]::Equals([string]$startupValue,$expectedStartupValue,[StringComparison]::Ordinal)) "Startup registration is not exact. Expected '$expectedStartupValue', found '$startupValue'."
    $checks.startupRegistrationExact = $true
    Close-MainToTray -Window $candidateMain -Process $applicationProcess
    $trayIcon = Get-ProductTrayIcon
    $desktopRoot = [System.Windows.Automation.AutomationElement]::RootElement
    $null = Save-UiEvidence -Name 'candidate-pending-logoff-tray' -Root $desktopRoot
    $checks.closeLeavesCandidateInActualTray = $true
    $finalDatabaseRecords = Copy-DatabaseEvidence -Name 'pending-logoff'
    $workbookFinal = Get-FileRecord $workbookPath
    $prePhaseCompletedUtc = [DateTime]::UtcNow
    $checks.externalLogoffStillRequired = $true

    $checks.pendingStateWritten = $true
    Write-ActionLog 'PRE_PHASE_COMPLETE_PENDING_EXTERNAL_LOGOFF_LOGON_START_SCRIPT_DID_NOT_LOG_OFF'
    $preResult = [ordered]@{
        schemaVersion = 1
        gate = 'installed-lifecycle-upgrade'
        phase = 'pre-logoff'
        status = 'PendingExternalLogoffLogon'
        success = $true
        evidenceId = $evidenceId
        phaseLinkNonce = $phaseLinkNonce
        startedUtc = $startedUtc.ToString('O')
        phaseCompletedUtc = $prePhaseCompletedUtc.ToString('O')
        confirmation = [ordered]@{ disposableCleanVm=$ConfirmDisposableCleanVm.IsPresent; vmSnapshotName=$VmSnapshotName; vmSnapshotId=$VmSnapshotId; standardUser=(-not $isAdministrator); localFreshEvidenceDirectory=$true }
        environment = [ordered]@{ powershellVersion=$PSVersionTable.PSVersion.ToString(); powershellEdition=$PSVersionTable.PSEdition; computerName=$env:COMPUTERNAME; preLogon=$preSession }
        prior = [ordered]@{ version=$PriorVersion; installerPath=$priorInstaller; installerSha256=$priorInstallerHash; expectedInstallerSha256=$ExpectedPriorInstallerSha256.ToUpperInvariant(); installerProductVersion=$priorInstallerProductVersion; expectedApplicationSha256=$ExpectedPriorApplicationSha256.ToUpperInvariant(); installedApplicationSha256=$priorApplicationHash; retainedApplication=$priorApplicationRecord; uninstallEntryBeforeUpgrade=$priorEntryBeforeUpgrade }
        candidate = [ordered]@{ version=$CandidateVersion; installerPath=$candidateInstaller; installerSha256=$candidateInstallerHash; expectedInstallerSha256=$ExpectedCandidateInstallerSha256.ToUpperInvariant(); installerProductVersion=$candidateInstallerProductVersion; expectedApplicationSha256=$ExpectedCandidateApplicationSha256.ToUpperInvariant(); installedApplicationSha256=$candidateApplicationHash; retainedApplication=$candidateApplicationRecord; uninstallEntryAfterUpgrade=$candidateEntry; installStartedUtc=$candidateInstallStartedUtc.ToString('O'); installCompletedUtc=$candidateInstallCompletedUtc.ToString('O'); installMode='in-place-over-prior-without-uninstall' }
        probe = [ordered]@{ path=$probe; sha256=$probeHash; expectedSha256=$ExpectedProbeSha256.ToUpperInvariant() }
        paths = [ordered]@{ application=$application; installDirectory=$installDirectory; database=$database; dataDirectory=$dataDirectory; startMenuShortcut=$startMenuShortcut; startupRegistryPath='HKCU\Software\Microsoft\Windows\CurrentVersion\Run\ExcelDiffTracker'; workbook=$workbookPath }
        syntheticChange = [ordered]@{ address=$address; priorValue=$preValue; candidateValue=$postValue }
        workbook = [ordered]@{ initial=$workbookInitial; pendingLogoff=$workbookFinal }
        probes = [ordered]@{ priorBaseline=$priorBaselineProbe; priorSequence1=$priorProbe; candidatePreservedSequence1=$candidatePreservedProbe; candidateSequence2=$candidateSaveProbe }
        database = [ordered]@{ closedBeforeUpgrade=$preDatabaseRecord; immediatelyAfterInstaller=$postInstallDatabaseRecord; beforeUpgradeCopies=$preUpgradeDatabaseRecords; pendingLogoffCopies=$finalDatabaseRecords }
        startup = [ordered]@{ expected=$expectedStartupValue; actual=[string]$startupValue }
        pendingProcess = [ordered]@{ processId=$applicationProcess.Id; startedUtc=$applicationProcess.StartTime.ToUniversalTime().ToString('O'); executablePath=$applicationProcess.Path; executableSha256=Get-Hash $application; mainWindowVisible=(Test-MainWindowVisible -ProcessId $applicationProcess.Id); trayIconName=$trayIcon.Current.Name; trayIconAutomationId=$trayIcon.Current.AutomationId }
        uiEvidence = @($uiRecords)
        actionLog = Get-FileRecord $actionLogPath
        checks = $checks
        externalTransition = [ordered]@{ performedByThisScript=$false; requiredNextAction='Exit this PowerShell, perform an actual Windows user logoff, sign in to the same account, then run Complete-InstalledLifecycleUpgradeGate.ps1.' }
        failure = $null
    }
    Write-Json -Path $preResultPath -Value $preResult
    $preRecord = Get-FileRecord $preResultPath
    $pending = [ordered]@{
        schemaVersion=1
        gate='installed-lifecycle-upgrade'
        state='PendingExternalLogoffLogon'
        evidenceId=$evidenceId
        phaseLinkNonce=$phaseLinkNonce
        preLogoffResult=$preRecord
        preLogon=$preSession
        prePhaseCompletedUtc=$prePhaseCompletedUtc.ToString('O')
        priorVersion=$PriorVersion
        candidateVersion=$CandidateVersion
        priorInstallerSha256=$priorInstallerHash
        candidateInstallerSha256=$candidateInstallerHash
        priorApplicationSha256=$priorApplicationHash
        candidateApplicationSha256=$candidateApplicationHash
        probeSha256=$probeHash
        workbookPath=$workbookPath
        expectedSequence=2
        expectedStartupValue=$expectedStartupValue
        applicationPath=$application
        databasePath=$database
        createdUtc=[DateTime]::UtcNow.ToString('O')
        logoffPerformedByStartScript=$false
    }
    Write-Json -Path $pendingPath -Value $pending
    Write-Output "INSTALLED_LIFECYCLE_UPGRADE_PENDING_EXTERNAL_LOGOFF_LOGON|evidence=$evidence|evidenceId=$evidenceId|pending=$pendingPath"
}
catch {
    $failure = $_.Exception.ToString()
    try {
        Write-Json -Path $failurePath -Value ([ordered]@{ schemaVersion=1; gate='installed-lifecycle-upgrade'; phase='pre-logoff'; status='Failed'; evidenceId=$evidenceId; startedUtc=$startedUtc.ToString('O'); failedUtc=[DateTime]::UtcNow.ToString('O'); checks=$checks; failure=$failure; pendingStateWritten=(Test-Path $pendingPath); logoffPerformedByThisScript=$false })
    } catch { }
    throw
}
finally {
    if ($excelWorkbook) { try { $excelWorkbook.Close($false) } catch { } }
    if ($excel) { try { $excel.Quit() } catch { } }
}
