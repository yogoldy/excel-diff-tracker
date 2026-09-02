[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string] $AcceptanceDirectory,
    [Parameter(Mandatory)] [string] $InstallerPath,
    [ValidateRange(2, 10)] [int] $RequiredRunCount = 2
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path $AcceptanceDirectory).Path
$installer = (Resolve-Path $InstallerPath).Path
$installerHash = (Get-FileHash $installer -Algorithm SHA256).Hash.ToUpperInvariant()

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
        $actual = (Get-FileHash $path -Algorithm SHA256).Hash
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
    [pscustomobject]@{ runId = $summary.runId; path = $run.FullName; installerSha256 = $summary.installerSha256 }
}

$visualMatrices = @(Get-ChildItem $root -Filter visual-matrix.json -File -Recurse)
$requiredScales = @(100, 125, 150, 200)
foreach ($scale in $requiredScales) {
    if (-not ($visualMatrices | Where-Object { (Get-Content $_.FullName -Raw | ConvertFrom-Json).scalePercent -eq $scale })) {
        throw "Visual evidence is missing required scaling configuration: $scale%."
    }
}
if (-not ($visualMatrices | Where-Object { (Get-Content $_.FullName -Raw | ConvertFrom-Json).windowsContrastTheme })) {
    throw 'Visual evidence is missing a Windows contrast-theme configuration.'
}
foreach ($matrixPath in $visualMatrices) {
    $matrix = Get-Content $matrixPath.FullName -Raw | ConvertFrom-Json
    if ($matrix.status -notin @('GeometryPassedHumanReviewRequired', 'Approved')) {
        throw "Visual geometry did not pass: $($matrixPath.FullName)"
    }
    if (@($matrix.results | Where-Object { -not $_.passed }).Count -ne 0) {
        throw "Visual matrix contains a failed geometry assertion: $($matrixPath.FullName)"
    }
}

$attestationDirectory = Join-Path $root 'attestations'
$requiredAttestations = @('functional-verifier.json', 'visual-accessibility-reviewer.json', 'release-custodian.json')
$attestations = foreach ($name in $requiredAttestations) {
    $path = Join-Path $attestationDirectory $name
    if (-not (Test-Path $path)) { throw "Required independent attestation is missing: $path" }
    $attestation = Get-Content $path -Raw | ConvertFrom-Json
    if ($attestation.status -ne 'Approved') { throw "Attestation is not Approved: $path" }
    if ($attestation.installerSha256.ToUpperInvariant() -ne $installerHash) { throw "Attestation names a different installer: $path" }
    if ([string]::IsNullOrWhiteSpace($attestation.reviewer)) { throw "Attestation reviewer is missing: $path" }
    [pscustomobject]@{ file = $name; reviewer = $attestation.reviewer; role = $attestation.role; status = $attestation.status }
}

$approval = [ordered]@{
    status = 'Approved'
    installerSha256 = $installerHash
    approvedUtc = [DateTime]::UtcNow.ToString('O')
    runs = $runSummaries
    visualMatrixCount = $visualMatrices.Count
    attestations = $attestations
}
$approvalPath = Join-Path $root 'approval.json'
$approval | ConvertTo-Json -Depth 8 | Set-Content $approvalPath -Encoding utf8NoBOM
Write-Output "ACCEPTANCE_EVIDENCE_APPROVED|$approvalPath|$installerHash"
