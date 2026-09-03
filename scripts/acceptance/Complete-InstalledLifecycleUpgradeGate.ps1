[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string] $EvidenceDirectory,
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
    [switch] $ConfirmDisposableVmAfterExternalLogoffLogon
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

if (-not $ConfirmDisposableVmAfterExternalLogoffLogon) {
    throw 'Run only after an external actual Windows user logoff/logon in the disposable VM, and pass -ConfirmDisposableVmAfterExternalLogoffLogon.'
}
if ($PSVersionTable.PSVersion.Major -ne 5 -or $PSVersionTable.PSVersion.Minor -ne 1 -or $PSVersionTable.PSEdition -ne 'Desktop') {
    throw "Stock Windows PowerShell 5.1 Desktop is required; found $($PSVersionTable.PSVersion) $($PSVersionTable.PSEdition)."
}
$isAdministrator = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if ($isAdministrator) { throw 'The lifecycle/upgrade gate must run as a standard non-administrator user.' }
if ([version]$CandidateVersion -le [version]$PriorVersion) { throw 'CandidateVersion must be newer than PriorVersion; a repair is not an upgrade.' }

$evidence = (Resolve-Path $EvidenceDirectory).Path
$priorInstaller = (Resolve-Path $PriorInstallerPath).Path
$candidateInstaller = (Resolve-Path $CandidateInstallerPath).Path
$probe = (Resolve-Path $ProbePath).Path
$pendingPath = Join-Path $evidence 'pending-logoff.json'
$preResultPath = Join-Path $evidence 'pre-logoff.json'
$finalResultPath = Join-Path $evidence 'installed-lifecycle-upgrade.json'
$checksumPath = Join-Path $evidence 'SHA256SUMS.txt'
$failurePath = Join-Path $evidence 'post-logon.failure.json'
$screenshots = Join-Path $evidence 'screenshots'
$uiaDirectory = Join-Path $evidence 'uia'
$probeDirectory = Join-Path $evidence 'probe'
$binaryDirectory = Join-Path $evidence 'binaries'
$databaseDirectory = Join-Path $evidence 'database'
$logDirectory = Join-Path $evidence 'logs'
$actionLogPath = Join-Path $logDirectory 'post-logon-actions.txt'

foreach ($path in @($evidence,$priorInstaller,$candidateInstaller,$probe)) {
    if ($path.StartsWith('\\')) { throw "All gate paths must be local to the disposable VM: $path" }
}
foreach ($path in @($pendingPath,$preResultPath)) {
    if (-not (Test-Path $path -PathType Leaf)) { throw "Required pre-logoff phase file is missing: $path" }
}
foreach ($path in @($finalResultPath,$checksumPath,$failurePath,$actionLogPath)) {
    if (Test-Path $path) { throw "Post-logon output already exists; the phase is append-only: $path" }
}

Import-Module (Join-Path $PSScriptRoot 'UiAutomation.psm1') -Force
$pending = Get-Content -LiteralPath $pendingPath -Raw | ConvertFrom-Json
$pre = Get-Content -LiteralPath $preResultPath -Raw | ConvertFrom-Json
$application = [System.IO.Path]::GetFullPath([string]$pending.applicationPath)
$database = [System.IO.Path]::GetFullPath([string]$pending.databasePath)
$installDirectory = Split-Path -Parent $application
$dataDirectory = Split-Path -Parent $database
$workbookPath = [System.IO.Path]::GetFullPath([string]$pending.workbookPath)
$startMenuShortcut = Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs\Excel Diff Tracker\Excel Diff Tracker.lnk'
$runKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'

$startedUtc = [DateTime]::UtcNow
$applicationProcess = $null
$failure = $null
$uiRecords = [System.Collections.Generic.List[object]]::new()
$checks = [ordered]@{
    pendingPhaseLinkExact = $false
    frozenIdentitiesExact = $false
    actualNewInteractiveLogonProven = $false
    candidateAutoStartedInNewLogon = $false
    autoStartedExecutableHashExact = $false
    autoStartWasQuietNoWelcomeOrMain = $false
    actualTrayIconPresent = $false
    nativeTrayDoubleClickReopened = $false
    exactHistoryDashboardAndProbe = $false
    trayExitStoppedCandidate = $false
    uninstallCompleted = $false
    installedBinariesRemoved = $false
    startMenuShortcutRemoved = $false
    startupEntryRemoved = $false
    uninstallRegistrationRemoved = $false
    localHistoryBytesRetained = $false
    exactHistoryProbePassesAfterUninstall = $false
}

function Assert-GateCondition {
    param([bool] $Condition,[string] $Message)
    if (-not $Condition) { throw $Message }
}

function Get-Hash {
    param([string] $Path)
    (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToUpperInvariant()
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

function Write-ActionLog {
    param([string] $Message)
    $line = '{0}|{1}' -f [DateTime]::UtcNow.ToString('O'),$Message
    [System.IO.File]::AppendAllText($actionLogPath,$line + [Environment]::NewLine,[System.Text.UTF8Encoding]::new($false))
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

function Resolve-EvidenceRecord {
    param([object] $Record)
    Assert-GateCondition ($null -ne $Record) 'A required phase-link evidence record is missing.'
    Assert-GateCondition (-not [System.IO.Path]::IsPathRooted([string]$Record.path)) 'A phase-link evidence path is rooted.'
    $full = [System.IO.Path]::GetFullPath((Join-Path $evidence ([string]$Record.path).Replace('/','\')))
    $prefix = $evidence.TrimEnd('\') + '\'
    Assert-GateCondition ($full.StartsWith($prefix,[StringComparison]::OrdinalIgnoreCase)) 'A phase-link evidence path escapes the evidence directory.'
    Assert-GateCondition (Test-Path $full -PathType Leaf) "A linked evidence file is missing: $($Record.path)"
    $item = Get-Item -LiteralPath $full
    Assert-GateCondition ($item.Length -eq [long]$Record.bytes) "A linked evidence byte count differs: $($Record.path)"
    Assert-GateCondition ((Get-Hash $full) -eq ([string]$Record.sha256).ToUpperInvariant()) "A linked evidence hash differs: $($Record.path)"
    $full
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
    Assert-GateCondition (-not (Test-Path $destination)) "Database checkpoint already exists: $destination"
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
    Assert-GateCondition ($sessionId -gt 0) 'The completion phase must run in an interactive Windows desktop session.'
    $explorers = @(Get-Process -Name explorer -ErrorAction SilentlyContinue | Where-Object { $_.SessionId -eq $sessionId })
    Assert-GateCondition ($explorers.Count -eq 1) "Expected exactly one Explorer shell in interactive session $sessionId; found $($explorers.Count)."
    [pscustomobject]@{
        accountName=$identity.Name
        accountSid=$identity.User.Value
        logonSid=$logonSids[0]
        windowsSessionId=$sessionId
        explorerProcessId=$explorers[0].Id
        explorerStartedUtc=$explorers[0].StartTime.ToUniversalTime().ToString('O')
        capturedUtc=[DateTime]::UtcNow.ToString('O')
    }
}

function Get-ProductUninstallEntries {
    $matches = [System.Collections.Generic.List[object]]::new()
    foreach ($root in @('HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall','HKCU:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall')) {
        if (-not (Test-Path $root)) { continue }
        foreach ($key in @(Get-ChildItem -Path $root -ErrorAction SilentlyContinue)) {
            $properties = Get-ItemProperty -LiteralPath $key.PSPath -ErrorAction SilentlyContinue
            if ($properties -and $properties.DisplayName -eq 'Excel Diff Tracker') {
                $matches.Add([pscustomobject]@{ registryPath=$key.PSPath; displayVersion=[string]$properties.DisplayVersion; installLocation=[string]$properties.InstallLocation; uninstallString=[string]$properties.UninstallString })
            }
        }
    }
    @($matches)
}

function Get-DesktopWindows {
    ([System.Windows.Automation.AutomationElement]::RootElement).FindAll([System.Windows.Automation.TreeScope]::Children,[System.Windows.Automation.Condition]::TrueCondition)
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

function Save-UiEvidence {
    param([string] $Name,[System.Windows.Automation.AutomationElement] $Root)
    $png = Join-Path $screenshots ($Name + '.png')
    $tree = Join-Path $uiaDirectory ($Name + '.json')
    Assert-GateCondition (-not (Test-Path $png) -and -not (Test-Path $tree)) "UI checkpoint already exists: $Name"
    Save-DesktopScreenshot -Path $png
    Export-UiaTree -Root $Root -Path $tree
    $record = [pscustomobject]@{ name=$Name; screenshot=Get-FileRecord $png; uiaTree=Get-FileRecord $tree; capturedUtc=[DateTime]::UtcNow.ToString('O') }
    $uiRecords.Add($record)
    $record
}

function Invoke-ProbeUntilPassed {
    param([string] $Name,[string[]] $Arguments,[int] $TimeoutSeconds = 30)
    $output = Join-Path $probeDirectory ($Name + '.json')
    Assert-GateCondition (-not (Test-Path $output)) "Probe checkpoint already exists: $output"
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

function Assert-DashboardState {
    param([System.Windows.Automation.AutomationElement] $Window,[long] $ExpectedSequence)
    Invoke-UiaElement -Element (Find-UiaElement -Root $Window -AutomationId 'DashboardNavigationButton')
    Start-Sleep -Milliseconds 300
    $pathElement = Find-UiaElement -Root $Window -Name $workbookPath
    $walker = [System.Windows.Automation.TreeWalker]::ControlViewWalker
    $ancestor = $pathElement
    $proved = $false
    for ($index = 0; $index -lt 8 -and $ancestor; $index++) {
        $nodes = $ancestor.FindAll([System.Windows.Automation.TreeScope]::Descendants,[System.Windows.Automation.Condition]::TrueCondition)
        $names = @($nodes | ForEach-Object { $_.Current.Name })
        if ($names -contains 'VERSIONS' -and $names -contains ([string]$ExpectedSequence)) { $proved = $true; break }
        $ancestor = $walker.GetParent($ancestor)
    }
    Assert-GateCondition $proved "The reopened dashboard did not bind the tracked workbook to exact sequence $ExpectedSequence."
}

function Assert-HistoryState {
    param([System.Windows.Automation.AutomationElement] $Window,[object] $ProbeResult)
    Invoke-UiaElement -Element (Find-UiaElement -Root $Window -AutomationId 'HistoryNavigationButton')
    Start-Sleep -Milliseconds 400
    $nodes = $Window.FindAll([System.Windows.Automation.TreeScope]::Descendants,[System.Windows.Automation.Condition]::TrueCondition)
    $names = @($nodes | ForEach-Object { [string]$_.Current.Name })
    Assert-GateCondition (@($names | Where-Object { $_ -like 'Version 2*' }).Count -eq 1) 'History did not expose exactly one Version 2 record.'
    Assert-GateCondition (@($names | Where-Object { $_ -like 'Version 1*' }).Count -eq 1) 'History did not expose exactly one Version 1 record.'
    Assert-GateCondition ($names -contains ([string]$ProbeResult.latestVersion.summary)) 'History did not expose the exact latest summary returned by the external probe.'
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
    Wait-AcceptanceCondition -TimeoutSeconds 15 -FailureMessage 'The app did not exit from the tray menu.' -Condition { $Process.Refresh(); $Process.HasExited }
}

function Write-Json {
    param([string] $Path,[object] $Value)
    Write-AcceptanceUtf8File -Path $Path -Content ($Value | ConvertTo-Json -Depth 30)
}

function Write-ChecksumManifest {
    $lines = [System.Collections.Generic.List[string]]::new()
    $files = @(Get-ChildItem -LiteralPath $evidence -Recurse -File | Where-Object { $_.FullName -ne $checksumPath } | Sort-Object FullName)
    foreach ($file in $files) {
        $relative = Get-Relative $file.FullName
        $lines.Add(('{0} *{1}' -f (Get-Hash $file.FullName),$relative))
    }
    Write-AcceptanceUtf8File -Path $checksumPath -Content (($lines -join "`r`n") + "`r`n")
}

try {
    Write-ActionLog 'POST_PHASE_STARTED_NO_PRODUCT_LAUNCH_PERFORMED'
    Assert-GateCondition ($pending.schemaVersion -eq 1 -and $pending.gate -eq 'installed-lifecycle-upgrade' -and $pending.state -eq 'PendingExternalLogoffLogon') 'Pending state identity is invalid.'
    Assert-GateCondition ($pre.schemaVersion -eq 1 -and $pre.gate -eq 'installed-lifecycle-upgrade' -and $pre.phase -eq 'pre-logoff' -and $pre.status -eq 'PendingExternalLogoffLogon' -and $pre.success) 'Pre-logoff result did not pass into the pending state.'
    $linkedPrePath = Resolve-EvidenceRecord $pending.preLogoffResult
    Assert-GateCondition ([string]::Equals($linkedPrePath,$preResultPath,[StringComparison]::OrdinalIgnoreCase)) 'Pending state links a different pre-logoff result path.'
    Assert-GateCondition ($pending.evidenceId -eq $pre.evidenceId -and $pending.phaseLinkNonce -eq $pre.phaseLinkNonce) 'Pending/pre evidence ID or phase-link nonce differs.'
    Assert-GateCondition (-not $pending.logoffPerformedByStartScript -and -not $pre.externalTransition.performedByThisScript) 'Pre phase falsely claims that it performed the external logoff.'
    Assert-GateCondition ($pending.expectedSequence -eq 2) 'Pending state does not require exact sequence 2.'
    $checks.pendingPhaseLinkExact = $true

    $priorInstallerHash = Get-Hash $priorInstaller
    $candidateInstallerHash = Get-Hash $candidateInstaller
    $probeHash = Get-Hash $probe
    $null = & (Join-Path $PSScriptRoot 'Test-SingleFilePayload.ps1') -ExecutablePath $probe
    Assert-GateCondition ($priorInstallerHash -eq $ExpectedPriorInstallerSha256.ToUpperInvariant() -and $priorInstallerHash -eq $pending.priorInstallerSha256 -and $priorInstallerHash -eq $pre.prior.installerSha256) 'Prior installer identity differs across phases.'
    Assert-GateCondition ($candidateInstallerHash -eq $ExpectedCandidateInstallerSha256.ToUpperInvariant() -and $candidateInstallerHash -eq $pending.candidateInstallerSha256 -and $candidateInstallerHash -eq $pre.candidate.installerSha256) 'Candidate installer identity differs across phases.'
    Assert-GateCondition ($probeHash -eq $ExpectedProbeSha256.ToUpperInvariant() -and $probeHash -eq $pending.probeSha256 -and $probeHash -eq $pre.probe.sha256) 'AcceptanceProbe identity differs across phases.'
    Assert-GateCondition ($pending.priorApplicationSha256 -eq $ExpectedPriorApplicationSha256.ToUpperInvariant() -and $pre.prior.installedApplicationSha256 -eq $ExpectedPriorApplicationSha256.ToUpperInvariant()) 'Prior application identity differs across phases.'
    Assert-GateCondition ($pending.candidateApplicationSha256 -eq $ExpectedCandidateApplicationSha256.ToUpperInvariant() -and $pre.candidate.installedApplicationSha256 -eq $ExpectedCandidateApplicationSha256.ToUpperInvariant()) 'Candidate application identity differs across phases.'
    Assert-GateCondition ($PriorVersion -eq $pending.priorVersion -and $PriorVersion -eq $pre.prior.version -and $CandidateVersion -eq $pending.candidateVersion -and $CandidateVersion -eq $pre.candidate.version) 'Prior/candidate version identities differ across phases.'
    Assert-GateCondition ([version]$CandidateVersion -gt [version]$PriorVersion -and $priorInstallerHash -ne $candidateInstallerHash -and $ExpectedPriorApplicationSha256.ToUpperInvariant() -ne $ExpectedCandidateApplicationSha256.ToUpperInvariant()) 'The linked run is a repair rather than an actual version/payload upgrade.'
    $checks.frozenIdentitiesExact = $true

    $postSession = Get-LogonIdentity
    $preCompleted = [DateTime]::Parse([string]$pending.prePhaseCompletedUtc).ToUniversalTime()
    $postExplorerStarted = [DateTime]::Parse([string]$postSession.explorerStartedUtc).ToUniversalTime()
    Assert-GateCondition ($postSession.accountSid -eq $pending.preLogon.accountSid -and $postSession.accountName -eq $pending.preLogon.accountName) 'Completion is not running as the same Windows account.'
    Assert-GateCondition ($postSession.logonSid -ne $pending.preLogon.logonSid) 'The token logon SID is unchanged. An actual user logoff/logon has not occurred.'
    Assert-GateCondition ($postSession.explorerProcessId -ne $pending.preLogon.explorerProcessId) 'The Explorer shell process identity is unchanged from the pre-logoff phase.'
    Assert-GateCondition ($postExplorerStarted -gt $preCompleted) 'The current interactive shell did not start after the pre-logoff phase completed.'
    $checks.actualNewInteractiveLogonProven = $true

    Assert-GateCondition (Test-Path $application -PathType Leaf) 'The candidate executable is absent after logon.'
    $processes = @(Get-Process -Name ExcelDiffTracker -ErrorAction SilentlyContinue | Where-Object { try { [string]::Equals($_.Path,$application,[StringComparison]::OrdinalIgnoreCase) } catch { $false } })
    Assert-GateCondition ($processes.Count -eq 1) "Expected exactly one auto-started candidate process before any launch action; found $($processes.Count)."
    $applicationProcess = $processes[0]
    $processStartedUtc = $applicationProcess.StartTime.ToUniversalTime()
    Assert-GateCondition ($processStartedUtc -gt $preCompleted -and $processStartedUtc -ge $postExplorerStarted) 'The candidate process was not started by the new interactive logon startup interval.'
    $checks.candidateAutoStartedInNewLogon = $true
    $autoStartedHash = Get-Hash $application
    $null = & (Join-Path $PSScriptRoot 'Test-SingleFilePayload.ps1') -ExecutablePath $application
    Assert-GateCondition ($autoStartedHash -eq $ExpectedCandidateApplicationSha256.ToUpperInvariant()) 'The auto-started executable hash differs from the frozen candidate.'
    $checks.autoStartedExecutableHashExact = $true

    $windows = @(Get-DesktopWindows | Where-Object { $_.Current.ProcessId -eq $applicationProcess.Id -and $_.Current.Name -in @('Welcome to Excel Diff Tracker','Excel Diff Tracker') })
    Assert-GateCondition ($windows.Count -eq 0) 'Candidate auto-start was not quiet: Welcome or the main window was visible before tray interaction.'
    $checks.autoStartWasQuietNoWelcomeOrMain = $true
    $trayIcon = Get-ProductTrayIcon
    $trayIconName = $trayIcon.Current.Name
    $trayIconAutomationId = $trayIcon.Current.AutomationId
    $checks.actualTrayIconPresent = $true
    $null = Save-UiEvidence -Name 'post-logon-quiet-autostart-tray' -Root ([System.Windows.Automation.AutomationElement]::RootElement)

    Invoke-UiaMouseClick -Element $trayIcon -Button DoubleLeft
    $main = Find-UiaWindow -Title 'Excel Diff Tracker' -ProcessId $applicationProcess.Id -TimeoutSeconds 20
    $welcome = @(Get-DesktopWindows | Where-Object { $_.Current.ProcessId -eq $applicationProcess.Id -and $_.Current.Name -eq 'Welcome to Excel Diff Tracker' })
    Assert-GateCondition ($welcome.Count -eq 0) 'Onboarding repeated after tray reopen in the new logon.'
    $checks.nativeTrayDoubleClickReopened = $true

    $address = [string]$pre.syntheticChange.address
    $priorValue = [string]$pre.syntheticChange.priorValue
    $candidateValue = [string]$pre.syntheticChange.candidateValue
    $postLogonProbe = Invoke-ProbeUntilPassed -Name 'post-logon-sequence-2' -Arguments @(
        '--database',$database,'--workbook',$workbookPath,'--expected-sequence','2','--expected-version-count','2',
        '--expected-error-count','0','--require-active','--require-no-last-error','--require-unique-version-hashes','--require-source-hash-match',
        '--require-ready-report','--expected-cell-change-count','1','--expected-sheet-change-count','0','--address',$address,
        '--expected-value',$candidateValue,'--expected-before-value',$priorValue,'--expected-kind','LiteralChanged')
    Assert-GateCondition ($postLogonProbe.result.currentHash -eq $pre.probes.candidateSequence2.result.currentHash) 'Post-logon exact tracked hash differs from the pre-logoff candidate checkpoint.'
    Assert-DashboardState -Window $main -ExpectedSequence 2
    $null = Save-UiEvidence -Name 'post-logon-dashboard-sequence-2' -Root $main
    Assert-HistoryState -Window $main -ProbeResult $postLogonProbe.result
    $null = Save-UiEvidence -Name 'post-logon-history-sequences-1-2' -Root $main
    $checks.exactHistoryDashboardAndProbe = $true

    Exit-ThroughTray -Process $applicationProcess
    $checks.trayExitStoppedCandidate = $true
    Write-ActionLog 'CANDIDATE_EXITED_THROUGH_TRAY'
    $beforeUninstallDatabaseHash = Get-Hash $database
    $beforeUninstallDatabaseBytes = (Get-Item -LiteralPath $database).Length
    $beforeUninstallDatabaseRecords = Copy-DatabaseEvidence -Name 'before-uninstall-after-tray-exit'
    $entries = @(Get-ProductUninstallEntries)
    Assert-GateCondition ($entries.Count -eq 1 -and $entries[0].displayVersion -eq $CandidateVersion) 'The exact candidate uninstall registration was not present before uninstall.'
    $uninstaller = Join-Path $installDirectory 'unins000.exe'
    Assert-GateCondition (Test-Path $uninstaller -PathType Leaf) 'The candidate uninstaller is missing.'
    $retainedUninstaller = Join-Path $binaryDirectory 'candidate-unins000.exe'
    Copy-Item -LiteralPath $uninstaller -Destination $retainedUninstaller
    $uninstallerRecord = Get-FileRecord $retainedUninstaller
    Write-ActionLog 'UNINSTALL_STARTED'
    $uninstallStartedUtc = [DateTime]::UtcNow
    $uninstall = Start-Process -FilePath $uninstaller -ArgumentList '/VERYSILENT','/SUPPRESSMSGBOXES','/NORESTART' -Wait -PassThru
    $uninstallCompletedUtc = [DateTime]::UtcNow
    Assert-GateCondition ($uninstall.ExitCode -eq 0) "Uninstaller exited with code $($uninstall.ExitCode)."
    $checks.uninstallCompleted = $true
    Assert-GateCondition (-not (Test-Path $installDirectory)) 'Installed binaries remained after uninstall.'
    $checks.installedBinariesRemoved = $true
    Assert-GateCondition (-not (Test-Path $startMenuShortcut)) 'The Start-menu shortcut remained after uninstall.'
    $checks.startMenuShortcutRemoved = $true
    $remainingStartup = Get-ProductStartupRegistration
    Assert-GateCondition (-not $remainingStartup.exists) 'The startup registry entry remained after uninstall.'
    $checks.startupEntryRemoved = $true
    Assert-GateCondition (@(Get-ProductUninstallEntries).Count -eq 0) 'The product uninstall registration remained after uninstall.'
    $checks.uninstallRegistrationRemoved = $true
    Assert-GateCondition ((Test-Path $dataDirectory -PathType Container) -and (Test-Path $database -PathType Leaf)) 'Local history data was removed by uninstall.'
    $afterUninstallDatabaseHash = Get-Hash $database
    $afterUninstallDatabaseBytes = (Get-Item -LiteralPath $database).Length
    Assert-GateCondition ($afterUninstallDatabaseHash -eq $beforeUninstallDatabaseHash -and $afterUninstallDatabaseBytes -eq $beforeUninstallDatabaseBytes) 'Uninstall changed the retained history database bytes.'
    $checks.localHistoryBytesRetained = $true
    $postUninstallProbe = Invoke-ProbeUntilPassed -Name 'post-uninstall-retained-sequence-2' -Arguments @(
        '--database',$database,'--workbook',$workbookPath,'--expected-sequence','2','--expected-version-count','2',
        '--expected-error-count','0','--require-active','--require-no-last-error','--require-unique-version-hashes','--require-source-hash-match',
        '--require-ready-report','--expected-cell-change-count','1','--expected-sheet-change-count','0','--address',$address,
        '--expected-value',$candidateValue,'--expected-before-value',$priorValue,'--expected-kind','LiteralChanged')
    Assert-GateCondition ($postUninstallProbe.result.currentHash -eq $postLogonProbe.result.currentHash) 'Post-uninstall external probe did not preserve the exact tracked hash.'
    $checks.exactHistoryProbePassesAfterUninstall = $true
    $afterUninstallDatabaseRecords = Copy-DatabaseEvidence -Name 'after-uninstall-retained-history'
    Write-ActionLog 'UNINSTALL_COMPLETE_LOCAL_HISTORY_RETAINED'

    $failedChecks = @($checks.GetEnumerator() | Where-Object { -not $_.Value })
    Assert-GateCondition ($failedChecks.Count -eq 0) "Post-logon checks are incomplete: $($failedChecks.Name -join ', ')"
    $completedUtc = [DateTime]::UtcNow
    $result = [ordered]@{
        schemaVersion=1
        gate='installed-lifecycle-upgrade'
        status='Passed'
        success=$true
        evidenceId=[string]$pending.evidenceId
        phaseLinkNonce=[string]$pending.phaseLinkNonce
        startedUtc=$startedUtc.ToString('O')
        completedUtc=$completedUtc.ToString('O')
        durationSeconds=($completedUtc-$startedUtc).TotalSeconds
        phaseLinks=[ordered]@{ preLogoffResult=Get-FileRecord $preResultPath; pendingState=Get-FileRecord $pendingPath }
        environment=[ordered]@{ powershellVersion=$PSVersionTable.PSVersion.ToString(); powershellEdition=$PSVersionTable.PSEdition; standardUser=(-not $isAdministrator); preLogon=$pending.preLogon; postLogon=$postSession }
        identities=[ordered]@{ priorVersion=$PriorVersion; candidateVersion=$CandidateVersion; priorInstallerSha256=$priorInstallerHash; candidateInstallerSha256=$candidateInstallerHash; priorApplicationSha256=$ExpectedPriorApplicationSha256.ToUpperInvariant(); candidateApplicationSha256=$autoStartedHash; probeSha256=$probeHash }
        autoStart=[ordered]@{ processId=$applicationProcess.Id; processStartedUtc=$processStartedUtc.ToString('O'); executablePath=$application; executableSha256=$autoStartedHash; welcomeWindowCount=0; mainWindowCountBeforeTrayOpen=0; trayIconName=$trayIconName; trayIconAutomationId=$trayIconAutomationId }
        exactState=[ordered]@{ workbookPath=$workbookPath; address=$address; priorValue=$priorValue; candidateValue=$candidateValue; expectedSequence=2; postLogonProbe=$postLogonProbe; postUninstallProbe=$postUninstallProbe }
        uninstall=[ordered]@{ uninstaller=$uninstallerRecord; startedUtc=$uninstallStartedUtc.ToString('O'); completedUtc=$uninstallCompletedUtc.ToString('O'); exitCode=$uninstall.ExitCode; installDirectoryRemoved=(-not (Test-Path $installDirectory)); startMenuShortcutRemoved=(-not (Test-Path $startMenuShortcut)); startupEntryRemoved=(-not $remainingStartup.exists); uninstallRegistrationCount=@(Get-ProductUninstallEntries).Count; dataDirectoryRetained=(Test-Path $dataDirectory); databaseRetained=(Test-Path $database); databaseSha256Before=$beforeUninstallDatabaseHash; databaseSha256After=$afterUninstallDatabaseHash; databaseBytesBefore=[long]$beforeUninstallDatabaseBytes; databaseBytesAfter=[long]$afterUninstallDatabaseBytes; databaseEvidenceBefore=$beforeUninstallDatabaseRecords; databaseEvidenceAfter=$afterUninstallDatabaseRecords }
        uiEvidence=@($uiRecords)
        actionLog=Get-FileRecord $actionLogPath
        checks=$checks
        checksumsPath='SHA256SUMS.txt'
        failure=$null
    }
    Write-Json -Path $finalResultPath -Value $result
    Write-ChecksumManifest
    Write-Output "INSTALLED_LIFECYCLE_UPGRADE_PASS|result=$finalResultPath|checksums=$checksumPath|evidenceId=$($pending.evidenceId)"
}
catch {
    $failure = $_.Exception.ToString()
    try {
        Write-Json -Path $failurePath -Value ([ordered]@{ schemaVersion=1; gate='installed-lifecycle-upgrade'; phase='post-logon'; status='Failed'; evidenceId=[string]$pending.evidenceId; startedUtc=$startedUtc.ToString('O'); failedUtc=[DateTime]::UtcNow.ToString('O'); checks=$checks; failure=$failure })
    } catch { }
    throw
}
