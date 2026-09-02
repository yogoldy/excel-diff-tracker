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

$runs = @(Get-ChildItem $root -Directory -Recurse | Where-Object { Test-Path (Join-Path $_.FullName 'acceptance.json') })
if ($runs.Count -ne $RequiredRunCount) {
    throw "Expected exactly $RequiredRunCount black-box acceptance runs, found $($runs.Count)."
}
$runSummaries = foreach ($run in $runs | Sort-Object Name) {
    Test-ChecksumManifest $run.FullName
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
