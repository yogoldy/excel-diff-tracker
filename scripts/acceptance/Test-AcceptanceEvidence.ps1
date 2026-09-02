[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string] $AcceptanceDirectory,
    [Parameter(Mandatory)] [string] $InstallerPath,
    [Parameter(Mandatory)] [string] $ProbePath,
    [Parameter(Mandatory)] [string] $XlsmFixture,
    [Parameter(Mandatory)] [string] $PriorInstallerPath,
    [Parameter(Mandatory)] [ValidatePattern('^[A-Fa-f0-9]{64}$')] [string] $ExpectedPriorApplicationSha256,
    [Parameter(Mandatory)] [ValidatePattern('^\d+\.\d+\.\d+$')] [string] $PriorVersion,
    [Parameter(Mandatory)] [ValidatePattern('^\d+\.\d+\.\d+$')] [string] $CandidateVersion,
    [ValidateRange(2, 10)] [int] $RequiredRunCount = 2
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'UiAutomation.psm1') -Force
$root = (Resolve-Path $AcceptanceDirectory).Path
$installer = (Resolve-Path $InstallerPath).Path
$probe = (Resolve-Path $ProbePath).Path
$xlsmFixture = (Resolve-Path $XlsmFixture).Path
$priorInstaller = (Resolve-Path $PriorInstallerPath).Path
$installerHash = Get-AcceptanceFileSha256 -Path $installer
$largeBenchmarkValidator = (Resolve-Path (Join-Path $PSScriptRoot 'Test-LargeWorkbookBenchmarkResult.ps1')).Path
$soakValidator = (Resolve-Path (Join-Path $PSScriptRoot 'Test-RealExcelSoakResult.ps1')).Path
$semanticMatrixValidator = (Resolve-Path (Join-Path $PSScriptRoot 'Test-InstalledSemanticMatrixResult.ps1')).Path
$visualMatrixValidator = (Resolve-Path (Join-Path $PSScriptRoot 'Test-VisualMatrixBundle.ps1')).Path
$recoveryMatrixValidator = (Resolve-Path (Join-Path $PSScriptRoot 'Test-InstalledRecoveryMatrixResult.ps1')).Path
$lifecycleUpgradeValidator = (Resolve-Path (Join-Path $PSScriptRoot 'Test-InstalledLifecycleUpgradeResult.ps1')).Path

if ([version]$CandidateVersion -le [version]$PriorVersion) {
    throw "CandidateVersion must be newer than PriorVersion; got $PriorVersion -> $CandidateVersion."
}

function Test-ChecksumManifest {
    param([string] $RunDirectory)
    $manifestPath = Join-Path $RunDirectory 'SHA256SUMS.txt'
    if (-not (Test-Path $manifestPath)) { throw "Evidence checksum manifest is missing: $manifestPath" }
    $seen = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($line in [System.IO.File]::ReadAllLines($manifestPath)) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        if ($line -notmatch '^([A-Fa-f0-9]{64})  (.+)$') { throw "Malformed evidence checksum: $line" }
        $relative = $Matches[2].Replace('/', [System.IO.Path]::DirectorySeparatorChar)
        if (-not $seen.Add($relative)) { throw "Duplicate evidence checksum path: $relative" }
        $path = Join-Path $RunDirectory $relative
        if (-not (Test-Path $path -PathType Leaf)) { throw "Evidence file is missing: $path" }
        $actual = Get-AcceptanceFileSha256 -Path $path
        if ($actual -ne $Matches[1].ToUpperInvariant()) { throw "Evidence checksum mismatch: $path" }
    }
    $files = @(Get-ChildItem $RunDirectory -File -Recurse | Where-Object FullName -ne $manifestPath)
    if ($seen.Count -ne $files.Count) {
        throw "Evidence manifest describes $($seen.Count) files, but $($files.Count) files exist: $RunDirectory"
    }
}

function Test-RecoveryEvidence {
    param([string] $RunDirectory)
    $path = Join-Path $RunDirectory 'recovery\recovery.json'
    if (-not (Test-Path $path -PathType Leaf)) {
        throw "Mandatory exclusive-lock recovery evidence is missing: $path"
    }
    $recovery = Get-Content $path -Raw | ConvertFrom-Json
    if ($recovery.scenarioId -ne 'held-open-exclusive-lock-over-60s') {
        throw "Unexpected recovery scenario identity: $path"
    }
    if ($recovery.status -ne 'Passed') {
        throw "Exclusive-lock recovery gate did not pass: $path"
    }
    foreach ($phase in @('baseline', 'warning', 'recovered', 'settled')) {
        $phaseProperty = $recovery.PSObject.Properties[$phase]
        $probeResult = if ($null -eq $phaseProperty) { $null } else { $phaseProperty.Value }
        if ($null -eq $probeResult -or $probeResult.passed -ne $true -or @($probeResult.failures).Count -ne 0) {
            throw "Exclusive-lock recovery probe phase did not pass ($phase): $path"
        }
    }

    $requiredChecks = @(
        'baselineAtSequenceZero',
        'lockHeldBeyond60Seconds',
        'warningRecordedExactlyOnce',
        'baselinePreservedDuringWarning',
        'actionableWarningRenderedExactlyOnce',
        'lockedBytesMatchChangedCandidate',
        'recoveredWithin20Seconds',
        'exactDeltaCaptured',
        'returnedToActive',
        'noDuplicateAfterReconciliation',
        'noFileMutationAfterRelease'
    )
    foreach ($name in $requiredChecks) {
        $property = $recovery.checks.PSObject.Properties[$name]
        if ($null -eq $property -or $property.Value -ne $true) {
            throw "Exclusive-lock recovery check is missing or failed ($name): $path"
        }
    }

    if ([double]$recovery.timing.lockedDurationSeconds -lt 60) {
        throw "Exclusive lock was not held beyond the product timeout: $path"
    }
    if ([double]$recovery.timing.recoverySeconds -lt 0 -or [double]$recovery.timing.recoverySeconds -gt 20) {
        throw "Automatic recovery did not complete within 20 seconds: $path"
    }
    if ([long]$recovery.baseline.currentSequence -ne 0 -or [long]$recovery.warning.currentSequence -ne 0) {
        throw "The baseline advanced before the exclusive lock was released: $path"
    }
    if ($recovery.baseline.currentHash -ne $recovery.warning.currentHash) {
        throw "The baseline hash changed while the workbook was exclusively locked: $path"
    }
    if ([long]$recovery.recovered.currentSequence -ne 1 -or [long]$recovery.settled.currentSequence -ne 1) {
        throw "Recovery was missed or duplicated: $path"
    }
    if ([long]$recovery.recovered.latestVersion.cellChangeCount -ne 1 -or
        [long]$recovery.recovered.latestVersion.sheetChangeCount -ne 0 -or
        $recovery.recovered.cellChange.address -ne $recovery.address -or
        $recovery.recovered.cellChange.kinds.Split(',') -notcontains 'LiteralAdded') {
        throw "Recovery did not capture the one exact expected cell delta: $path"
    }
    $recoveredAfter = $recovery.recovered.cellChange.afterJson | ConvertFrom-Json
    if ($recoveredAfter.literalValue -ne $recovery.expectedValue) {
        throw "Recovery captured the wrong literal value: $path"
    }
    if ($recovery.recovered.workbookStatus -ne 'Active' -or
        -not [string]::IsNullOrWhiteSpace($recovery.recovered.lastError) -or
        $recovery.recovered.latestVersion.reportStatus -ne 'Ready') {
        throw "Workbook did not return to Active with its current error cleared: $path"
    }
    if ([long]$recovery.warning.errorCount -ne 1 -or [long]$recovery.recovered.errorCount -ne 1 -or [long]$recovery.settled.errorCount -ne 1) {
        throw "The exclusive lock did not produce exactly one retained warning record: $path"
    }
    if ([long]$recovery.warningUi.categoryCount -ne 1 -or [long]$recovery.warningUi.messageCount -lt 1 -or [long]$recovery.warningUi.workbookPathCount -lt 1) {
        throw "The actionable recovery warning was not rendered exactly once: $path"
    }
    if ($recovery.hashes.candidate -ne $recovery.hashes.sourceAfterRelease -or
        $recovery.hashes.sourceAfterRelease -ne $recovery.hashes.sourceAfterRecovery -or
        $recovery.hashes.sourceAfterRecovery -ne $recovery.settled.currentHash) {
        throw "Recovery evidence does not prove unchanged source bytes after lock release: $path"
    }
}

function Test-GoldenExcelEvidence {
    param([string] $RunDirectory)
    $cases = @(
        @{ File = 'Acceptance.xlsx-sequence-1.json'; Sequence = 1; Value = 'test'; Before = $null; Kind = 'LiteralAdded'; Cleared = $false },
        @{ File = 'Acceptance.xlsx-sequence-2.json'; Sequence = 2; Value = 'test2'; Before = 'test'; Kind = 'LiteralChanged'; Cleared = $false },
        @{ File = 'Acceptance.xlsx-sequence-3.json'; Sequence = 3; Value = $null; Before = 'test2'; Kind = 'LiteralCleared'; Cleared = $true },
        @{ File = 'Acceptance Macro.xlsm-sequence-1.json'; Sequence = 1; Value = 'test'; Before = $null; Kind = 'LiteralAdded'; Cleared = $false },
        @{ File = 'Acceptance Macro.xlsm-sequence-2.json'; Sequence = 2; Value = 'test2'; Before = 'test'; Kind = 'LiteralChanged'; Cleared = $false },
        @{ File = 'Acceptance Macro.xlsm-sequence-3.json'; Sequence = 3; Value = $null; Before = 'test2'; Kind = 'LiteralCleared'; Cleared = $true }
    )
    foreach ($case in $cases) {
        $path = Join-Path (Join-Path $RunDirectory 'probe') $case.File
        if (-not (Test-Path $path -PathType Leaf)) { throw "Mandatory golden-path probe evidence is missing: $path" }
        $probe = Get-Content $path -Raw | ConvertFrom-Json
        if ($probe.passed -ne $true -or @($probe.failures).Count -ne 0 -or
            $probe.workbookStatus -ne 'Active' -or $probe.errorCount -ne 0 -or
            $probe.currentSequence -ne $case.Sequence -or $probe.versionCount -ne $case.Sequence -or
            $probe.distinctVersionHashCount -ne $probe.versionCount -or
            $probe.latestVersion.sequence -ne $case.Sequence -or
            $probe.latestVersion.sha256 -ne $probe.currentHash -or
            $probe.latestVersion.cellChangeCount -ne 1 -or $probe.latestVersion.sheetChangeCount -ne 0 -or
            $probe.latestVersion.reportStatus -ne 'Ready' -or
            $probe.cellChange.address -ne 'Z1000' -or
            $probe.cellChange.kinds.Split(',') -notcontains $case.Kind) {
            throw "Golden-path exact-diff evidence is inconsistent: $path"
        }
        $after = if ($null -eq $probe.cellChange.afterJson) { $null } else { $probe.cellChange.afterJson | ConvertFrom-Json }
        if ($case.Cleared) {
            if ($null -ne $after -and $null -ne $after.literalValue) { throw "Golden-path clear retained a literal value: $path" }
        } elseif ($after.literalValue -ne $case.Value) {
            throw "Golden-path new literal value is wrong: $path"
        }
        if ($null -eq $case.Before) {
            if ($null -ne $probe.cellChange.beforeJson) { throw "Golden-path add did not start from an empty cell: $path" }
        } else {
            $before = $probe.cellChange.beforeJson | ConvertFrom-Json
            if ($before.literalValue -ne $case.Before) { throw "Golden-path old literal value is wrong: $path" }
        }
        $portableReport = Join-Path (Join-Path $RunDirectory 'reports') (Split-Path $probe.latestVersion.reportPath -Leaf)
        if (-not (Test-Path $portableReport -PathType Leaf)) { throw "Golden-path portable Markdown report is missing: $portableReport" }
        $markdown = [System.IO.File]::ReadAllText($portableReport)
        if ($markdown.IndexOf('Z1000', [StringComparison]::Ordinal) -lt 0) { throw "Golden-path Markdown omits the exact address: $portableReport" }
        if (-not $case.Cleared -and $markdown.IndexOf($case.Value, [StringComparison]::Ordinal) -lt 0) { throw "Golden-path Markdown omits the new value: $portableReport" }
    }
}

function Test-LifecycleEvidence {
    param([string] $RunDirectory, [object] $Environment)
    $path = Join-Path $RunDirectory 'lifecycle\lifecycle.json'
    if (-not (Test-Path $path -PathType Leaf)) { throw "Mandatory installed-app lifecycle evidence is missing: $path" }
    $lifecycle = Get-Content $path -Raw | ConvertFrom-Json
    if ($lifecycle.schemaVersion -ne 1 -or $lifecycle.gate -ne 'installed-app-lifecycle' -or $lifecycle.status -ne 'Passed') {
        throw "Installed-app lifecycle gate did not pass: $path"
    }
    $required = @(
        'closeKeepsTrayProcessAlive', 'actualTrayIconReopensWindow', 'trayExitStopsProcess',
        'backgroundLaunchIsQuiet', 'secondLaunchActivatesWindow', 'onboardingDoesNotRepeat',
        'repairInstallPreservesHistory', 'startupRegistrationIsExact')
    foreach ($name in $required) {
        $property = $lifecycle.checks.PSObject.Properties[$name]
        if ($null -eq $property -or $property.Value -ne $true) { throw "Lifecycle check is missing or failed ($name): $path" }
    }
    $expectedStartup = '"' + $Environment.installedApplicationPath + '" --background'
    if (-not [string]::Equals($lifecycle.startupValue, $expectedStartup, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Lifecycle startup value is inconsistent: $path"
    }
    $repairPath = Join-Path $RunDirectory 'lifecycle\repair-history.json'
    if (-not (Test-Path $repairPath -PathType Leaf)) { throw "Lifecycle repair-history probe is missing: $repairPath" }
    $repair = Get-Content $repairPath -Raw | ConvertFrom-Json
    if ($repair.passed -ne $true -or $repair.workbookStatus -ne 'Active' -or $repair.currentSequence -ne 3 -or
        $repair.versionCount -ne 3 -or $repair.distinctVersionHashCount -ne 3 -or $repair.errorCount -ne 0) {
        throw "Repair install did not preserve the exact synthetic history: $repairPath"
    }
    foreach ($requiredFile in @('screenshots\tray-hidden-main.png', 'screenshots\tray-reopened-main.png', 'uia\tray-reopened-main.json')) {
        $evidencePath = Join-Path $RunDirectory $requiredFile
        if (-not (Test-Path $evidencePath -PathType Leaf) -or (Get-Item $evidencePath).Length -le 0) { throw "Lifecycle UI evidence is missing: $evidencePath" }
    }
}

$runs = @(Get-ChildItem $root -Directory -Recurse | Where-Object { Test-Path (Join-Path $_.FullName 'acceptance.json') })
if ($runs.Count -ne $RequiredRunCount) {
    throw "Expected exactly $RequiredRunCount black-box acceptance runs, found $($runs.Count)."
}
$runSummaries = foreach ($run in $runs | Sort-Object Name) {
    Test-ChecksumManifest $run.FullName
    Test-RecoveryEvidence $run.FullName
    $summary = Get-Content (Join-Path $run.FullName 'acceptance.json') -Raw | ConvertFrom-Json
    if ($summary.status -ne 'Passed') { throw "Acceptance run is not Passed: $($run.FullName)" }
    if ($summary.installerSha256.ToUpperInvariant() -ne $installerHash) {
        throw "Acceptance run used a different installer: $($run.FullName)"
    }
    if (@($summary.assertions | Where-Object { -not $_.passed }).Count -ne 0) {
        throw "Acceptance run contains a failed assertion: $($run.FullName)"
    }
    if ($summary.installedApplicationSha256 -notmatch '^[A-Fa-f0-9]{64}$') {
        throw "Acceptance run is missing the frozen installed-executable hash: $($run.FullName)"
    }
    if ([string]::IsNullOrWhiteSpace($summary.vmSnapshotName) -or [string]::IsNullOrWhiteSpace($summary.vmSnapshotId)) {
        throw "Acceptance run is missing its clean VM snapshot identity: $($run.FullName)"
    }
    $environmentPath = Join-Path $run.FullName 'environment.json'
    if (-not (Test-Path $environmentPath -PathType Leaf)) { throw "Acceptance environment manifest is missing: $environmentPath" }
    $environment = Get-Content $environmentPath -Raw | ConvertFrom-Json
    Test-GoldenExcelEvidence $run.FullName
    Test-LifecycleEvidence $run.FullName $environment
    if ($environment.isAdministrator -ne $false) { throw "Acceptance run did not use a standard non-administrator account: $environmentPath" }
    if ($environment.dotnetOnPath -ne $false) { throw "Development .NET was present on PATH during acceptance: $environmentPath" }
    if ($environment.os.OSArchitecture -notlike 'ARM*') { throw "Acceptance run did not use Windows ARM64: $environmentPath" }
    if ($environment.windowsPowerShell -notlike '5.1.*') { throw "Acceptance run did not use stock Windows PowerShell 5.1: $environmentPath" }
    if ($environment.excelPeMachine -notin @('0xAA64', '0x8664')) { throw "Acceptance run did not record a supported desktop Excel architecture: $environmentPath" }
    if ([string]::IsNullOrWhiteSpace($environment.excel.version) -or [string]::IsNullOrWhiteSpace($environment.excel.build)) { throw "Acceptance run is missing the Excel version/build: $environmentPath" }
    if ($environment.vmSnapshotName -ne $summary.vmSnapshotName -or $environment.vmSnapshotId -ne $summary.vmSnapshotId) { throw "Acceptance snapshot identity is inconsistent: $environmentPath" }
    if ($environment.installerSha256.ToUpperInvariant() -ne $installerHash -or $environment.installedApplicationSha256.ToUpperInvariant() -ne $summary.installedApplicationSha256.ToUpperInvariant()) { throw "Acceptance environment candidate hashes are inconsistent: $environmentPath" }

    $largeResults = @(Get-ChildItem $run.FullName -Filter 'large-workbook-benchmark.json' -File -Recurse)
    if ($largeResults.Count -ne 1) {
        throw "Expected exactly one mandatory large-workbook benchmark result, found $($largeResults.Count): $($run.FullName)"
    }
    $null = & $largeBenchmarkValidator `
        -ResultPath $largeResults[0].FullName `
        -InstallerPath $installer `
        -ExpectedApplicationSha256 $summary.installedApplicationSha256

    $soakResults = @(Get-ChildItem $run.FullName -Filter 'real-excel-soak.json' -File -Recurse)
    if ($soakResults.Count -ne 1) {
        throw "Expected exactly one mandatory real-Excel soak result, found $($soakResults.Count): $($run.FullName)"
    }
    $null = & $soakValidator `
        -ResultPath $soakResults[0].FullName `
        -InstallerPath $installer `
        -ExpectedApplicationSha256 $summary.installedApplicationSha256

    $semanticResults = @(Get-ChildItem $run.FullName -Filter 'installed-semantic-matrix.json' -File -Recurse)
    if ($semanticResults.Count -ne 1) {
        throw "Expected exactly one mandatory installed semantic-matrix result, found $($semanticResults.Count): $($run.FullName)"
    }
    $null = & $semanticMatrixValidator `
        -ResultPath $semanticResults[0].FullName `
        -InstallerPath $installer `
        -ExpectedApplicationSha256 $summary.installedApplicationSha256 `
        -ProbePath $probe `
        -XlsmFixture $xlsmFixture

    $recoveryMatrixResults = @(Get-ChildItem $run.FullName -Filter 'installed-recovery-matrix.json' -File -Recurse)
    if ($recoveryMatrixResults.Count -ne 1) {
        throw "Expected exactly one mandatory installed recovery-matrix result, found $($recoveryMatrixResults.Count): $($run.FullName)"
    }
    $null = & $recoveryMatrixValidator `
        -ResultPath $recoveryMatrixResults[0].FullName `
        -InstallerPath $installer `
        -ExpectedApplicationSha256 $summary.installedApplicationSha256 `
        -ProbePath $probe

    [pscustomobject]@{
        runId = $summary.runId
        runNumber = [int]$summary.runNumber
        path = $run.FullName
        sourceCommit = $summary.sourceCommit
        vmSnapshotName = $summary.vmSnapshotName
        vmSnapshotId = $summary.vmSnapshotId
        installerSha256 = $summary.installerSha256
        installedApplicationSha256 = $summary.installedApplicationSha256.ToUpperInvariant()
    }
}

if (@($runSummaries.installedApplicationSha256 | Select-Object -Unique).Count -ne 1) {
    throw 'The two clean acceptance runs used different installed executable bytes.'
}
if (@($runSummaries.runId | Select-Object -Unique).Count -ne $RequiredRunCount) {
    throw 'Black-box acceptance run IDs are not unique.'
}
$requiredRunNumbers = @(1..$RequiredRunCount)
if (@(Compare-Object $requiredRunNumbers @($runSummaries.runNumber)).Count -ne 0) {
    throw "Black-box acceptance run numbers must be exactly 1 through $RequiredRunCount."
}
if (@($runSummaries.sourceCommit | Select-Object -Unique).Count -ne 1) {
    throw 'The two clean acceptance runs name different source commits.'
}

$lifecycleUpgradeResults = @(Get-ChildItem $root -Filter 'installed-lifecycle-upgrade.json' -File -Recurse)
if ($lifecycleUpgradeResults.Count -ne 1) {
    throw "Expected exactly one mandatory installed lifecycle/upgrade result, found $($lifecycleUpgradeResults.Count)."
}
$lifecycleUpgradeValidation = & $lifecycleUpgradeValidator `
    -ResultPath $lifecycleUpgradeResults[0].FullName `
    -PriorInstallerPath $priorInstaller `
    -ExpectedPriorApplicationSha256 $ExpectedPriorApplicationSha256 `
    -PriorVersion $PriorVersion `
    -CandidateInstallerPath $installer `
    -ExpectedCandidateApplicationSha256 $runSummaries[0].installedApplicationSha256 `
    -CandidateVersion $CandidateVersion `
    -ProbePath $probe
if (@($lifecycleUpgradeValidation).Count -ne 1 -or [string]$lifecycleUpgradeValidation -notlike 'INSTALLED_LIFECYCLE_UPGRADE_VALID|*') {
    throw 'The exact installed lifecycle/upgrade validator did not approve the candidate.'
}

$visualValidation = & $visualMatrixValidator `
    -EvidenceRoot $root `
    -ExpectedInstallerSha256 $installerHash `
    -ExpectedApplicationSha256 $runSummaries[0].installedApplicationSha256 `
    -RequireHumanApproval
if ($visualValidation.status -ne 'Approved') {
    throw 'The strict visual/accessibility bundle validator did not approve the candidate.'
}
$visualMatrices = @(Get-ChildItem $root -Filter visual-matrix.json -File -Recurse)

function Get-EvidenceDigest {
    $excludedRootFiles = @('approval.json', 'EVIDENCE_SHA256SUMS.txt')
    $digestLines = Get-ChildItem $root -File -Recurse |
        Where-Object {
            $relative = Get-AcceptanceRelativePath -BasePath $root -Path $_.FullName -UseForwardSlash
            $topLevelExcluded = $relative.IndexOf('/', [StringComparison]::Ordinal) -lt 0 -and $excludedRootFiles -contains $relative
            -not $topLevelExcluded -and -not $relative.StartsWith('attestations/', [StringComparison]::OrdinalIgnoreCase)
        } |
        Sort-Object FullName |
        ForEach-Object {
            $relative = Get-AcceptanceRelativePath -BasePath $root -Path $_.FullName -UseForwardSlash
            "$((Get-AcceptanceFileSha256 -Path $_.FullName).ToLowerInvariant())  $relative"
        }
    $manifestText = (($digestLines -join "`n") + "`n")
    Write-AcceptanceUtf8File -Path (Join-Path $root 'EVIDENCE_SHA256SUMS.txt') -Content $manifestText
    $bytes = [System.Text.UTF8Encoding]::new($false).GetBytes($manifestText)
    $algorithm = [System.Security.Cryptography.SHA256]::Create()
    try { -join @($algorithm.ComputeHash($bytes) | ForEach-Object { $_.ToString('X2') }) }
    finally { $algorithm.Dispose() }
}

$evidenceDigest = Get-EvidenceDigest

$attestationDirectory = Join-Path $root 'attestations'
$requiredAttestations = [ordered]@{
    'functional-verifier.json' = 'Functional verifier'
    'visual-accessibility-reviewer.json' = 'Visual/accessibility reviewer'
    'release-custodian.json' = 'Release custodian'
}
$attestations = foreach ($entry in $requiredAttestations.GetEnumerator()) {
    $name = $entry.Key
    $path = Join-Path $attestationDirectory $name
    if (-not (Test-Path $path)) { throw "Required independent attestation is missing: $path" }
    $attestation = Get-Content $path -Raw | ConvertFrom-Json
    if ($attestation.status -ne 'Approved') { throw "Attestation is not Approved: $path" }
    if ($attestation.installerSha256.ToUpperInvariant() -ne $installerHash) { throw "Attestation names a different installer: $path" }
    if ([string]::IsNullOrWhiteSpace($attestation.reviewer)) { throw "Attestation reviewer is missing: $path" }
    if ($attestation.role -ne $entry.Value) { throw "Attestation has the wrong role: $path" }
    if ($attestation.sourceCommit -ne $runSummaries[0].sourceCommit) { throw "Attestation names a different source commit: $path" }
    if ($attestation.installedApplicationSha256.ToUpperInvariant() -ne $runSummaries[0].installedApplicationSha256) { throw "Attestation names a different installed executable: $path" }
    if ($attestation.evidenceSha256.ToUpperInvariant() -ne $evidenceDigest) { throw "Attestation is not bound to the current evidence digest: $path" }
    if ([string]::IsNullOrWhiteSpace($attestation.reviewedUtc)) { throw "Attestation review timestamp is missing: $path" }
    [pscustomobject]@{ file = $name; reviewer = $attestation.reviewer; role = $attestation.role; status = $attestation.status }
}
if (@($attestations.reviewer | Select-Object -Unique).Count -ne $requiredAttestations.Count) {
    throw 'Independent attestations must name three distinct reviewers.'
}

$approval = [ordered]@{
    status = 'Approved'
    installerSha256 = $installerHash
    installedApplicationSha256 = $runSummaries[0].installedApplicationSha256
    sourceCommit = $runSummaries[0].sourceCommit
    evidenceSha256 = $evidenceDigest
    approvedUtc = [DateTime]::UtcNow.ToString('O')
    runs = $runSummaries
    lifecycleUpgrade = [ordered]@{
        priorVersion = $PriorVersion
        candidateVersion = $CandidateVersion
        resultPath = $lifecycleUpgradeResults[0].FullName
        status = 'Passed'
    }
    visualMatrixCount = $visualMatrices.Count
    attestations = $attestations
}
$approvalPath = Join-Path $root 'approval.json'
Write-AcceptanceUtf8File -Path $approvalPath -Content ($approval | ConvertTo-Json -Depth 8)
Write-Output "ACCEPTANCE_EVIDENCE_APPROVED|$approvalPath|$installerHash"
