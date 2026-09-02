[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string] $ResultPath,
    [Parameter(Mandatory)] [string] $PriorInstallerPath,
    [Parameter(Mandatory)] [ValidatePattern('^[A-Fa-f0-9]{64}$')] [string] $ExpectedPriorApplicationSha256,
    [Parameter(Mandatory)] [ValidatePattern('^\d+\.\d+\.\d+$')] [string] $PriorVersion,
    [Parameter(Mandatory)] [string] $CandidateInstallerPath,
    [Parameter(Mandatory)] [ValidatePattern('^[A-Fa-f0-9]{64}$')] [string] $ExpectedCandidateApplicationSha256,
    [Parameter(Mandatory)] [ValidatePattern('^\d+\.\d+\.\d+$')] [string] $CandidateVersion,
    [Parameter(Mandatory)] [string] $ProbePath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$resultPath = (Resolve-Path $ResultPath).Path
$root = Split-Path -Parent $resultPath
$priorInstaller = (Resolve-Path $PriorInstallerPath).Path
$candidateInstaller = (Resolve-Path $CandidateInstallerPath).Path
$probe = (Resolve-Path $ProbePath).Path
$checksumPath = Join-Path $root 'SHA256SUMS.txt'
$prePath = Join-Path $root 'pre-logoff.json'
$pendingPath = Join-Path $root 'pending-logoff.json'
foreach ($path in @($checksumPath,$prePath,$pendingPath)) {
    if (-not (Test-Path $path -PathType Leaf)) { throw "Invalid installed lifecycle/upgrade evidence: required file is missing: $path" }
}
$result = Get-Content -LiteralPath $resultPath -Raw | ConvertFrom-Json
$pre = Get-Content -LiteralPath $prePath -Raw | ConvertFrom-Json
$pending = Get-Content -LiteralPath $pendingPath -Raw | ConvertFrom-Json

function Require-Condition {
    param([bool] $Condition,[string] $Message)
    if (-not $Condition) { throw "Invalid installed lifecycle/upgrade evidence: $Message" }
}

function Get-EvidenceUtc {
    param([object] $Value,[string] $Name)
    $parsed = [DateTime]::MinValue
    Require-Condition (-not [string]::IsNullOrWhiteSpace([string]$Value) -and [DateTime]::TryParse([string]$Value,[ref]$parsed)) "$Name timestamp is missing or invalid"
    $parsed.ToUniversalTime()
}

function Get-Hash {
    param([string] $Path)
    (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToUpperInvariant()
}

function Get-Relative {
    param([string] $Path)
    $full = [System.IO.Path]::GetFullPath($Path)
    $prefix = $root.TrimEnd('\') + '\'
    Require-Condition ($full.StartsWith($prefix,[StringComparison]::OrdinalIgnoreCase)) "file is outside the evidence root: $full"
    $full.Substring($prefix.Length).Replace('\','/')
}

function Resolve-EvidenceRecord {
    param([object] $Record)
    Require-Condition ($null -ne $Record) 'an evidence record is missing'
    Require-Condition (-not [string]::IsNullOrWhiteSpace([string]$Record.path)) 'an evidence record has an empty path'
    Require-Condition (-not [System.IO.Path]::IsPathRooted([string]$Record.path)) "an evidence path is rooted: $($Record.path)"
    $full = [System.IO.Path]::GetFullPath((Join-Path $root ([string]$Record.path).Replace('/','\')))
    $prefix = $root.TrimEnd('\') + '\'
    Require-Condition ($full.StartsWith($prefix,[StringComparison]::OrdinalIgnoreCase)) "an evidence path escapes the run directory: $($Record.path)"
    Require-Condition (Test-Path $full -PathType Leaf) "an evidence file is missing: $($Record.path)"
    $item = Get-Item -LiteralPath $full
    Require-Condition ($item.Length -gt 0) "an evidence file is empty: $($Record.path)"
    Require-Condition ($item.Length -eq [long]$Record.bytes) "evidence byte count differs: $($Record.path)"
    Require-Condition ((Get-Hash $full) -eq ([string]$Record.sha256).ToUpperInvariant()) "evidence hash differs: $($Record.path)"
    $full
}

function Require-AllChecks {
    param([object] $Checks,[string[]] $Expected,[string] $Phase)
    $properties = @($Checks.PSObject.Properties)
    $names = @($properties | ForEach-Object { $_.Name })
    Require-Condition ($names.Count -eq $Expected.Count -and @(Compare-Object $Expected $names).Count -eq 0) "$Phase check set differs"
    foreach ($name in $Expected) { Require-Condition ([bool]$Checks.$name) "$Phase check '$name' did not pass" }
}

function Validate-ProbeBundle {
    param(
        [object] $Bundle,
        [long] $Sequence,
        [string] $ExpectedValue,
        [string] $ExpectedBeforeValue,
        [string] $ExpectedKind,
        [switch] $ExpectBeforeMissing
    )
    $path = Resolve-EvidenceRecord $Bundle.record
    $raw = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
    Require-Condition ($raw.passed) "external probe did not pass: $($Bundle.record.path)"
    Require-Condition (($raw | ConvertTo-Json -Depth 20 -Compress) -eq ($Bundle.result | ConvertTo-Json -Depth 20 -Compress)) "stored probe summary differs from raw output: $($Bundle.record.path)"
    Require-Condition ($raw.currentSequence -eq $Sequence -and $raw.versionCount -eq $Sequence -and $raw.distinctVersionHashCount -eq $Sequence) "probe sequence/version counts differ: $($Bundle.record.path)"
    Require-Condition ($raw.errorCount -eq 0 -and $raw.workbookStatus -eq 'Active' -and [string]::IsNullOrWhiteSpace([string]$raw.lastError)) "probe status/error state differs: $($Bundle.record.path)"
    Require-Condition ($raw.latestVersion.sequence -eq $Sequence -and $raw.latestVersion.reportStatus -eq 'Ready') "probe latest version/report state differs: $($Bundle.record.path)"
    Require-Condition ($raw.latestVersion.cellChangeCount -eq 1 -and $raw.latestVersion.sheetChangeCount -eq 0) "probe exact change counts differ: $($Bundle.record.path)"
    Require-Condition ($raw.cellChange.address -eq $pre.syntheticChange.address -and $raw.cellChange.kinds.Split(',') -contains $ExpectedKind) "probe cell address/kind differs: $($Bundle.record.path)"
    $after = if ($raw.cellChange.afterJson) { $raw.cellChange.afterJson | ConvertFrom-Json } else { $null }
    Require-Condition ($null -ne $after -and [string]$after.literalValue -eq $ExpectedValue) "probe after-value differs: $($Bundle.record.path)"
    if ($ExpectBeforeMissing) {
        Require-Condition ($null -eq $raw.cellChange.beforeJson) "probe before-state should be missing: $($Bundle.record.path)"
    } else {
        $before = if ($raw.cellChange.beforeJson) { $raw.cellChange.beforeJson | ConvertFrom-Json } else { $null }
        Require-Condition ($null -ne $before -and [string]$before.literalValue -eq $ExpectedBeforeValue) "probe before-value differs: $($Bundle.record.path)"
    }
    $raw
}

function Validate-BaselineProbeBundle {
    param([object] $Bundle)
    $path = Resolve-EvidenceRecord $Bundle.record
    $raw = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
    Require-Condition ($raw.passed) 'prior baseline external probe did not pass'
    Require-Condition (($raw | ConvertTo-Json -Depth 20 -Compress) -eq ($Bundle.result | ConvertTo-Json -Depth 20 -Compress)) 'stored prior baseline differs from raw probe'
    Require-Condition ($raw.currentSequence -eq 0 -and $raw.versionCount -eq 0 -and $raw.distinctVersionHashCount -eq 0 -and $raw.errorCount -eq 0) 'prior baseline was not exact sequence/version/error zero'
    Require-Condition ($raw.workbookStatus -eq 'Active' -and [string]::IsNullOrWhiteSpace([string]$raw.lastError) -and $null -eq $raw.latestVersion) 'prior baseline status/version state differs'
    $raw
}

function Validate-UiSet {
    param(
        [object[]] $Records,
        [string[]] $RequiredNames,
        [string] $Phase,
        [DateTime] $PhaseStartedUtc,
        [DateTime] $PhaseCompletedUtc
    )
    Require-Condition ($PhaseCompletedUtc -gt $PhaseStartedUtc) "$Phase boundaries are invalid"
    $names = @($Records | ForEach-Object { [string]$_.name })
    Require-Condition ($names.Count -eq @($names | Select-Object -Unique).Count) "$Phase UI evidence contains duplicate names"
    foreach ($name in $RequiredNames) {
        $matches = @($Records | Where-Object { $_.name -eq $name })
        Require-Condition ($matches.Count -eq 1) "$Phase UI evidence '$name' is missing or duplicated"
        $null = Resolve-EvidenceRecord $matches[0].screenshot
        $null = Resolve-EvidenceRecord $matches[0].uiaTree
    }
    $previousCaptureUtc = $PhaseStartedUtc
    foreach ($record in $Records) {
        $capturedUtc = Get-EvidenceUtc $record.capturedUtc "$Phase UI evidence '$($record.name)'"
        Require-Condition ($capturedUtc -ge $PhaseStartedUtc -and $capturedUtc -le $PhaseCompletedUtc) "$Phase UI evidence '$($record.name)' falls outside its phase"
        Require-Condition ($capturedUtc -ge $previousCaptureUtc) "$Phase UI evidence timestamps are not monotonic at '$($record.name)'"
        $previousCaptureUtc = $capturedUtc
        $null = Resolve-EvidenceRecord $record.screenshot
        $null = Resolve-EvidenceRecord $record.uiaTree
    }
}

function Validate-DatabaseRecords {
    param([object[]] $Records,[string] $Name)
    Require-Condition (@($Records).Count -ge 1) "$Name contains no retained database evidence"
    foreach ($record in @($Records)) { $null = Resolve-EvidenceRecord $record }
}

function Validate-ChecksumManifest {
    $manifestLines = @(Get-Content -LiteralPath $checksumPath | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    $listed = @{}
    foreach ($line in $manifestLines) {
        Require-Condition ($line -match '^([A-Fa-f0-9]{64}) \*(.+)$') "invalid checksum line: $line"
        $hash = $matches[1].ToUpperInvariant()
        $relative = $matches[2]
        Require-Condition (-not [System.IO.Path]::IsPathRooted($relative) -and $relative -notmatch '(^|/)\.\.(/|$)') "unsafe checksum path: $relative"
        Require-Condition (-not $listed.ContainsKey($relative)) "duplicate checksum path: $relative"
        $full = [System.IO.Path]::GetFullPath((Join-Path $root $relative.Replace('/','\')))
        Require-Condition (Test-Path $full -PathType Leaf) "checksummed file is missing: $relative"
        Require-Condition ((Get-Hash $full) -eq $hash) "checksummed file hash differs: $relative"
        $listed[$relative] = $hash
    }
    $actual = @(Get-ChildItem -LiteralPath $root -Recurse -File | Where-Object { $_.FullName -ne $checksumPath } | ForEach-Object { Get-Relative $_.FullName })
    Require-Condition ($listed.Count -eq $actual.Count -and @(Compare-Object @($listed.Keys) $actual).Count -eq 0) 'checksum manifest does not cover exactly every evidence file except itself'
}

Validate-ChecksumManifest

$priorInstallerHash = Get-Hash $priorInstaller
$candidateInstallerHash = Get-Hash $candidateInstaller
$probeHash = Get-Hash $probe
$priorInstallerVersion = ([System.Diagnostics.FileVersionInfo]::GetVersionInfo($priorInstaller)).ProductVersion
$candidateInstallerVersion = ([System.Diagnostics.FileVersionInfo]::GetVersionInfo($candidateInstaller)).ProductVersion
$priorApplicationHash = $ExpectedPriorApplicationSha256.ToUpperInvariant()
$candidateApplicationHash = $ExpectedCandidateApplicationSha256.ToUpperInvariant()

Require-Condition ($result.schemaVersion -eq 1 -and $result.gate -eq 'installed-lifecycle-upgrade' -and $result.status -eq 'Passed' -and $result.success) 'final result identity/status is invalid'
Require-Condition ([string]::IsNullOrWhiteSpace([string]$result.failure)) 'final result retained a failure'
Require-Condition ($pre.schemaVersion -eq 1 -and $pre.gate -eq 'installed-lifecycle-upgrade' -and $pre.phase -eq 'pre-logoff' -and $pre.status -eq 'PendingExternalLogoffLogon' -and $pre.success) 'pre-logoff result identity/status is invalid'
Require-Condition ([string]::IsNullOrWhiteSpace([string]$pre.failure)) 'pre-logoff result retained a failure'
Require-Condition ($pending.schemaVersion -eq 1 -and $pending.gate -eq 'installed-lifecycle-upgrade' -and $pending.state -eq 'PendingExternalLogoffLogon') 'pending-state identity is invalid'
Require-Condition (-not $pending.logoffPerformedByStartScript -and -not $pre.externalTransition.performedByThisScript) 'the start phase falsely claims it performed logoff'
Require-Condition ($result.evidenceId -eq $pre.evidenceId -and $result.evidenceId -eq $pending.evidenceId) 'evidence IDs do not link all phases'
Require-Condition ($result.phaseLinkNonce -eq $pre.phaseLinkNonce -and $result.phaseLinkNonce -eq $pending.phaseLinkNonce) 'phase-link nonces do not link all phases'

$linkedPre = Resolve-EvidenceRecord $pending.preLogoffResult
Require-Condition ([string]::Equals($linkedPre,$prePath,[StringComparison]::OrdinalIgnoreCase)) 'pending state links a different pre-logoff path'
Require-Condition ((Get-Hash $prePath) -eq (Get-Hash (Resolve-EvidenceRecord $result.phaseLinks.preLogoffResult))) 'final and pending phase links disagree on pre-logoff bytes'
Require-Condition ([string]::Equals((Resolve-EvidenceRecord $result.phaseLinks.pendingState),$pendingPath,[StringComparison]::OrdinalIgnoreCase)) 'final result links a different pending-state path'

$preCheckNames = @(
    'disposableCleanVmConfirmed','cleanNoInstallNoDataState','frozenInputHashesExact','installerVersionIdentitiesExact',
    'priorInstalledPerUser','priorStartMenuLaunchUsed','priorOnboardingCompletedWithRealXlsx','priorVisibleExcelKeyboardSaveExact',
    'priorExitedThroughTray','candidateInstalledOverPriorWithoutUninstall','installedCandidateHashExact','candidateHistoryExactlyPreserved',
    'candidateOnboardingDidNotRepeat','candidateVisibleExcelKeyboardSaveExact','startupRegistrationExact','closeLeavesCandidateInActualTray',
    'pendingStateWritten','externalLogoffStillRequired')
Require-AllChecks -Checks $pre.checks -Expected $preCheckNames -Phase 'pre-logoff'
$postCheckNames = @(
    'pendingPhaseLinkExact','frozenIdentitiesExact','actualNewInteractiveLogonProven','candidateAutoStartedInNewLogon',
    'autoStartedExecutableHashExact','autoStartWasQuietNoWelcomeOrMain','actualTrayIconPresent','nativeTrayDoubleClickReopened',
    'exactHistoryDashboardAndProbe','trayExitStoppedCandidate','uninstallCompleted','installedBinariesRemoved',
    'startMenuShortcutRemoved','startupEntryRemoved','uninstallRegistrationRemoved','localHistoryBytesRetained','exactHistoryProbePassesAfterUninstall')
Require-AllChecks -Checks $result.checks -Expected $postCheckNames -Phase 'post-logon'

Require-Condition ($pre.confirmation.disposableCleanVm -and $pre.confirmation.standardUser -and $pre.confirmation.localFreshEvidenceDirectory) 'pre-logoff disposable/fresh/standard-user confirmation is incomplete'
Require-Condition (-not [string]::IsNullOrWhiteSpace([string]$pre.confirmation.vmSnapshotName) -and -not [string]::IsNullOrWhiteSpace([string]$pre.confirmation.vmSnapshotId)) 'VM snapshot identity is missing'
Require-Condition ($pre.environment.powershellVersion -match '^5\.1\.' -and $pre.environment.powershellEdition -eq 'Desktop') 'pre-logoff phase did not use Windows PowerShell 5.1 Desktop'
Require-Condition ($result.environment.powershellVersion -match '^5\.1\.' -and $result.environment.powershellEdition -eq 'Desktop' -and $result.environment.standardUser) 'post-logon phase did not use standard-user Windows PowerShell 5.1 Desktop'

Require-Condition ([version]$CandidateVersion -gt [version]$PriorVersion) 'candidate version is not newer than prior version'
Require-Condition ($priorInstallerVersion -eq $PriorVersion -and $candidateInstallerVersion -eq $CandidateVersion) 'supplied installer ProductVersion identities differ from expected versions'
Require-Condition ($priorInstallerHash -ne $candidateInstallerHash -and $priorApplicationHash -ne $candidateApplicationHash) 'the evidence supplies a same-installer or same-payload repair rather than an upgrade'
Require-Condition ($pre.prior.version -eq $PriorVersion -and $pre.prior.installerProductVersion -eq $PriorVersion -and $pre.prior.installerSha256 -eq $priorInstallerHash -and $pre.prior.expectedInstallerSha256 -eq $priorInstallerHash) 'prior installer/version identity differs'
Require-Condition ($pre.candidate.version -eq $CandidateVersion -and $pre.candidate.installerProductVersion -eq $CandidateVersion -and $pre.candidate.installerSha256 -eq $candidateInstallerHash -and $pre.candidate.expectedInstallerSha256 -eq $candidateInstallerHash) 'candidate installer/version identity differs'
Require-Condition ($pre.prior.installedApplicationSha256 -eq $priorApplicationHash -and $pre.prior.expectedApplicationSha256 -eq $priorApplicationHash) 'prior installed application identity differs'
Require-Condition ($pre.candidate.installedApplicationSha256 -eq $candidateApplicationHash -and $pre.candidate.expectedApplicationSha256 -eq $candidateApplicationHash) 'candidate installed application identity differs'
Require-Condition ($pre.probe.sha256 -eq $probeHash -and $pre.probe.expectedSha256 -eq $probeHash) 'pre-logoff probe identity differs'
Require-Condition ($result.identities.priorVersion -eq $PriorVersion -and $result.identities.candidateVersion -eq $CandidateVersion -and $result.identities.priorInstallerSha256 -eq $priorInstallerHash -and $result.identities.candidateInstallerSha256 -eq $candidateInstallerHash -and $result.identities.priorApplicationSha256 -eq $priorApplicationHash -and $result.identities.candidateApplicationSha256 -eq $candidateApplicationHash -and $result.identities.probeSha256 -eq $probeHash) 'final frozen identities differ'
Require-Condition ((Get-Hash (Resolve-EvidenceRecord $pre.prior.retainedApplication)) -eq $priorApplicationHash) 'retained prior executable hash differs'
Require-Condition ((Get-Hash (Resolve-EvidenceRecord $pre.candidate.retainedApplication)) -eq $candidateApplicationHash) 'retained candidate executable hash differs'
Require-Condition ($pre.prior.uninstallEntryBeforeUpgrade.displayVersion -eq $PriorVersion -and $pre.candidate.uninstallEntryAfterUpgrade.displayVersion -eq $CandidateVersion) 'installed prior/candidate DisplayVersion identity is incomplete'
Require-Condition ($pre.prior.uninstallEntryBeforeUpgrade.registryPath -eq $pre.candidate.uninstallEntryAfterUpgrade.registryPath) 'candidate did not upgrade the same product registration in place'
Require-Condition ($pre.candidate.installMode -eq 'in-place-over-prior-without-uninstall') 'candidate install is labeled as repair or replacement rather than in-place upgrade'
$candidateInstallStart = Get-EvidenceUtc $pre.candidate.installStartedUtc 'candidate install start'
$candidateInstallEnd = Get-EvidenceUtc $pre.candidate.installCompletedUtc 'candidate install completion'
Require-Condition ($candidateInstallEnd -gt $candidateInstallStart) 'candidate install timestamps are invalid'
Require-Condition ($pre.database.closedBeforeUpgrade.sha256 -eq $pre.database.immediatelyAfterInstaller.sha256 -and $pre.database.closedBeforeUpgrade.bytes -eq $pre.database.immediatelyAfterInstaller.bytes) 'closed database bytes changed during candidate upgrade'
Validate-DatabaseRecords @($pre.database.beforeUpgradeCopies) 'pre-upgrade database checkpoint'
Validate-DatabaseRecords @($pre.database.pendingLogoffCopies) 'pending-logoff database checkpoint'

$priorValue = [string]$pre.syntheticChange.priorValue
$candidateValue = [string]$pre.syntheticChange.candidateValue
Require-Condition (-not [string]::IsNullOrWhiteSpace($priorValue) -and -not [string]::IsNullOrWhiteSpace($candidateValue) -and $priorValue -ne $candidateValue -and $pre.syntheticChange.address -eq 'Z1000') 'synthetic change identity is invalid'
$null = Validate-BaselineProbeBundle -Bundle $pre.probes.priorBaseline
$priorProbe = Validate-ProbeBundle -Bundle $pre.probes.priorSequence1 -Sequence 1 -ExpectedValue $priorValue -ExpectedBeforeValue '' -ExpectedKind 'LiteralAdded' -ExpectBeforeMissing
$preservedProbe = Validate-ProbeBundle -Bundle $pre.probes.candidatePreservedSequence1 -Sequence 1 -ExpectedValue $priorValue -ExpectedBeforeValue '' -ExpectedKind 'LiteralAdded' -ExpectBeforeMissing
$candidateProbe = Validate-ProbeBundle -Bundle $pre.probes.candidateSequence2 -Sequence 2 -ExpectedValue $candidateValue -ExpectedBeforeValue $priorValue -ExpectedKind 'LiteralChanged'
Require-Condition ($priorProbe.currentHash -eq $preservedProbe.currentHash) 'candidate did not preserve the exact prior tracked hash during upgrade'

$expectedStartup = '"' + [string]$pre.paths.application + '" --background'
Require-Condition ($pre.startup.expected -eq $expectedStartup -and $pre.startup.actual -ceq $expectedStartup -and $pending.expectedStartupValue -ceq $expectedStartup) 'exact startup HKCU value differs'
Require-Condition ($pre.pendingProcess.executableSha256 -eq $candidateApplicationHash -and -not $pre.pendingProcess.mainWindowVisible -and $pre.pendingProcess.trayIconAutomationId -eq 'NotifyItemIcon') 'pending-logoff candidate tray/process proof differs'

$preStarted = Get-EvidenceUtc $pre.startedUtc 'pre-logoff start'
$preCompleted = Get-EvidenceUtc $pending.prePhaseCompletedUtc 'pre-logoff completion'
$preResultCompleted = Get-EvidenceUtc $pre.phaseCompletedUtc 'pre-logoff result completion'
$pendingCreated = Get-EvidenceUtc $pending.createdUtc 'pending-state creation'
$postStarted = Get-EvidenceUtc $result.startedUtc 'post-logon start'
$postCompleted = Get-EvidenceUtc $result.completedUtc 'post-logon completion'
Require-Condition ($preCompleted -gt $preStarted -and $preResultCompleted -eq $preCompleted) 'pre-logoff result and pending completion timestamps differ or are unordered'
Require-Condition ($pendingCreated -ge $preCompleted -and $pendingCreated -le $postCompleted) 'pending-state creation timestamp is outside the linked lifecycle'
Require-Condition ($postCompleted -gt $postStarted -and $postStarted -gt $preCompleted) 'post-logon phase does not follow pre-logoff completion'
Require-Condition ($candidateInstallStart -ge $preStarted -and $candidateInstallEnd -le $preCompleted) 'candidate install did not occur entirely within the pre-logoff phase'
$pendingProcessStarted = Get-EvidenceUtc $pre.pendingProcess.startedUtc 'pending candidate process start'
Require-Condition ($pendingProcessStarted -ge $candidateInstallEnd -and $pendingProcessStarted -le $preCompleted) 'pending candidate process did not start after installation within the pre-logoff phase'

$preLogonCaptured = Get-EvidenceUtc $pre.environment.preLogon.capturedUtc 'pre-logoff logon identity capture'
$pendingPreLogonCaptured = Get-EvidenceUtc $pending.preLogon.capturedUtc 'pending pre-logoff logon identity capture'
$resultPreLogonCaptured = Get-EvidenceUtc $result.environment.preLogon.capturedUtc 'result pre-logoff logon identity capture'
$postLogonCaptured = Get-EvidenceUtc $result.environment.postLogon.capturedUtc 'post-logon identity capture'
Require-Condition ($preLogonCaptured -eq $pendingPreLogonCaptured -and $preLogonCaptured -eq $resultPreLogonCaptured) 'pre-logoff logon identity capture timestamp differs across linked phases'
Require-Condition ($preLogonCaptured -ge $preStarted -and $preLogonCaptured -le $preCompleted) 'pre-logoff logon identity was not captured within the pre-logoff phase'
Require-Condition ($postLogonCaptured -ge $postStarted -and $postLogonCaptured -le $postCompleted) 'post-logon identity was not captured within the post-logon phase'

Validate-UiSet @($pre.uiEvidence) @(
    'prior-onboarding-step-1','prior-onboarding-complete','prior-dashboard-baseline','prior-visible-excel-ctrl-s',
    'prior-dashboard-sequence-1','candidate-dashboard-preserved-sequence-1','candidate-visible-excel-ctrl-s',
    'candidate-dashboard-sequence-2','candidate-pending-logoff-tray') 'pre-logoff' $preStarted $preCompleted
Validate-UiSet @($result.uiEvidence) @('post-logon-quiet-autostart-tray','post-logon-dashboard-sequence-2','post-logon-history-sequences-1-2') 'post-logon' $postStarted $postCompleted
$preLog = Resolve-EvidenceRecord $pre.actionLog
$postLog = Resolve-EvidenceRecord $result.actionLog
Require-Condition ((Get-Content -LiteralPath $preLog -Raw) -match 'IN_PLACE_CANDIDATE_INSTALL_STARTED_NO_UNINSTALL_WAS_INVOKED' -and (Get-Content -LiteralPath $preLog -Raw) -match 'START_SCRIPT_DID_NOT_LOG_OFF') 'pre-logoff action log lacks the no-uninstall/external-logoff boundary'
Require-Condition ((Get-Content -LiteralPath $postLog -Raw) -match 'POST_PHASE_STARTED_NO_PRODUCT_LAUNCH_PERFORMED' -and (Get-Content -LiteralPath $postLog -Raw) -match 'UNINSTALL_COMPLETE_LOCAL_HISTORY_RETAINED') 'post-logon action log lacks the auto-start observation/uninstall boundary'

$postExplorer = Get-EvidenceUtc $result.environment.postLogon.explorerStartedUtc 'post-logon Explorer start'
$autoStarted = Get-EvidenceUtc $result.autoStart.processStartedUtc 'candidate auto-start'
Require-Condition ($result.environment.preLogon.accountSid -eq $result.environment.postLogon.accountSid -and $result.environment.preLogon.accountName -eq $result.environment.postLogon.accountName) 'phases used different Windows accounts'
Require-Condition ($result.environment.preLogon.logonSid -ne $result.environment.postLogon.logonSid) 'same token logon session was reused across phases'
Require-Condition ($result.environment.preLogon.explorerProcessId -ne $result.environment.postLogon.explorerProcessId) 'same Explorer session was reused across phases'
Require-Condition ($postExplorer -gt $preCompleted -and $postExplorer -le $postLogonCaptured -and $autoStarted -ge $postExplorer -and $autoStarted -gt $preCompleted -and $autoStarted -le $postCompleted) 'new-logon shell/autostart timestamps do not follow the pre-logoff phase'
Require-Condition ($result.autoStart.executableSha256 -eq $candidateApplicationHash -and $result.autoStart.welcomeWindowCount -eq 0 -and $result.autoStart.mainWindowCountBeforeTrayOpen -eq 0 -and $result.autoStart.trayIconAutomationId -eq 'NotifyItemIcon') 'quiet exact-hash auto-start/tray proof differs'

$postLogonProbe = Validate-ProbeBundle -Bundle $result.exactState.postLogonProbe -Sequence 2 -ExpectedValue $candidateValue -ExpectedBeforeValue $priorValue -ExpectedKind 'LiteralChanged'
$postUninstallProbe = Validate-ProbeBundle -Bundle $result.exactState.postUninstallProbe -Sequence 2 -ExpectedValue $candidateValue -ExpectedBeforeValue $priorValue -ExpectedKind 'LiteralChanged'
Require-Condition ($result.exactState.expectedSequence -eq 2 -and $result.exactState.workbookPath -eq $pre.paths.workbook -and $result.exactState.address -eq $pre.syntheticChange.address) 'final exact-state identity differs'
Require-Condition ($candidateProbe.currentHash -eq $postLogonProbe.currentHash -and $postLogonProbe.currentHash -eq $postUninstallProbe.currentHash) 'exact history hash was not preserved across logon and uninstall'

Require-Condition ($result.uninstall.exitCode -eq 0 -and $result.uninstall.installDirectoryRemoved -and $result.uninstall.startMenuShortcutRemoved -and $result.uninstall.startupEntryRemoved -and $result.uninstall.uninstallRegistrationCount -eq 0) 'uninstall removal proof is incomplete'
Require-Condition ($result.uninstall.dataDirectoryRetained -and $result.uninstall.databaseRetained) 'uninstall did not retain local history data'
Require-Condition ($result.uninstall.databaseSha256Before -eq $result.uninstall.databaseSha256After -and $result.uninstall.databaseBytesBefore -eq $result.uninstall.databaseBytesAfter -and $result.uninstall.databaseBytesAfter -gt 0) 'uninstall changed or removed retained database bytes'
$null = Resolve-EvidenceRecord $result.uninstall.uninstaller
Validate-DatabaseRecords @($result.uninstall.databaseEvidenceBefore) 'pre-uninstall database checkpoint'
Validate-DatabaseRecords @($result.uninstall.databaseEvidenceAfter) 'post-uninstall retained database checkpoint'
$uninstallStart = Get-EvidenceUtc $result.uninstall.startedUtc 'uninstall start'
$uninstallEnd = Get-EvidenceUtc $result.uninstall.completedUtc 'uninstall completion'
Require-Condition ($uninstallStart -ge $postStarted -and $uninstallEnd -gt $uninstallStart -and $uninstallEnd -le $postCompleted) 'uninstall timestamps are invalid or outside the post-logon phase'

Require-Condition ($postCompleted -gt $postStarted -and $result.durationSeconds -gt 0) 'final result timestamps are invalid'
Require-Condition ($result.checksumsPath -eq 'SHA256SUMS.txt') 'final result points to an unexpected checksum manifest'

Write-Output "INSTALLED_LIFECYCLE_UPGRADE_VALID|result=$resultPath|evidenceId=$($result.evidenceId)"
