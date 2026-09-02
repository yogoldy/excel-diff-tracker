[CmdletBinding()]
param(
    [ValidatePattern('^\d+\.\d+\.\d+$')]
    [string]$Version = '0.1.2',
    [string]$Repository = 'yogoldy/excel-diff-tracker',
    [string]$ReleaseDirectory,
    [Parameter(Mandatory = $true)]
    [string]$AcceptanceDirectory,
    [string]$ProbePath,
    [Parameter(Mandatory = $true)]
    [string]$XlsmFixture,
    [Parameter(Mandatory = $true)]
    [string]$PriorInstallerPath,
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[A-Fa-f0-9]{64}$')]
    [string]$ExpectedPriorApplicationSha256,
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^\d+\.\d+\.\d+$')]
    [string]$PriorVersion,
    [switch]$Prerelease
)

$ErrorActionPreference = 'Stop'
$repositoryRoot = Split-Path -Parent $PSScriptRoot
if (-not $ReleaseDirectory) { $ReleaseDirectory = Join-Path $repositoryRoot 'artifacts\release' }
if (-not $ProbePath) { $ProbePath = Join-Path $repositoryRoot 'artifacts\acceptance-tools\win-arm64\ExcelDiffTracker.AcceptanceProbe.exe' }
$ReleaseDirectory = (Resolve-Path $ReleaseDirectory).Path
$ProbePath = (Resolve-Path $ProbePath).Path
$XlsmFixture = (Resolve-Path $XlsmFixture).Path
$PriorInstallerPath = (Resolve-Path $PriorInstallerPath).Path
$installer = Join-Path $ReleaseDirectory 'ExcelDiffTracker-Setup-arm64.exe'
$checksumFile = Join-Path $ReleaseDirectory 'SHA256SUMS.txt'
$releaseNotes = Join-Path $repositoryRoot "docs\RELEASE_NOTES_$Version.md"
$wingetManifest = Join-Path $repositoryRoot "packaging\winget\yogoldy.ExcelDiffTracker\$Version\yogoldy.ExcelDiffTracker.installer.yaml"

foreach ($required in @($installer, $checksumFile, $releaseNotes, $wingetManifest)) {
    if (-not (Test-Path $required)) { throw "Required release file was not found: $required" }
}

$actualHash = (Get-FileHash $installer -Algorithm SHA256).Hash.ToUpperInvariant()
$acceptanceDirectoryPath = (Resolve-Path $AcceptanceDirectory).Path
$acceptanceValidator = Join-Path $repositoryRoot 'scripts\acceptance\Test-AcceptanceEvidence.ps1'
$null = & $acceptanceValidator `
    -AcceptanceDirectory $acceptanceDirectoryPath `
    -InstallerPath $installer `
    -ProbePath $ProbePath `
    -XlsmFixture $XlsmFixture `
    -PriorInstallerPath $PriorInstallerPath `
    -ExpectedPriorApplicationSha256 $ExpectedPriorApplicationSha256 `
    -PriorVersion $PriorVersion `
    -CandidateVersion $Version
$approvalPath = Join-Path $acceptanceDirectoryPath 'approval.json'
if (-not (Test-Path $approvalPath)) {
    throw 'Acceptance approval is missing. Run scripts\acceptance\Test-AcceptanceEvidence.ps1 first.'
}
$approval = Get-Content $approvalPath -Raw | ConvertFrom-Json
if ($approval.status -ne 'Approved' -or $approval.installerSha256.ToUpperInvariant() -ne $actualHash) {
    throw 'Acceptance approval does not approve this exact installer.'
}
$headCommit = (& git -C $repositoryRoot rev-parse HEAD).Trim()
if ($LASTEXITCODE -ne 0 -or $approval.sourceCommit -ne $headCommit) {
    throw 'Acceptance approval does not name the source commit being published.'
}
$checksumText = [System.IO.File]::ReadAllText($checksumFile)
$manifestText = [System.IO.File]::ReadAllText($wingetManifest)
$checksumEntries = @{}
foreach ($line in [System.IO.File]::ReadAllLines($checksumFile)) {
    if ([string]::IsNullOrWhiteSpace($line)) { continue }
    if ($line -notmatch '^([A-Fa-f0-9]{64})  (.+)$') {
        throw "Malformed checksum line: $line"
    }
    if ($checksumEntries.ContainsKey($Matches[2])) {
        throw "Duplicate checksum entry: $($Matches[2])"
    }
    $checksumEntries[$Matches[2]] = $Matches[1].ToUpperInvariant()
}
$releaseFiles = @(Get-ChildItem $ReleaseDirectory -File | Where-Object Name -ne 'SHA256SUMS.txt')
if ($checksumEntries.Count -ne $releaseFiles.Count) {
    throw 'SHA256SUMS.txt does not contain exactly one entry for every release asset.'
}
foreach ($file in $releaseFiles) {
    if (-not $checksumEntries.ContainsKey($file.Name)) {
        throw "SHA256SUMS.txt is missing release asset: $($file.Name)"
    }
    $fileHash = (Get-FileHash $file.FullName -Algorithm SHA256).Hash.ToUpperInvariant()
    if ($checksumEntries[$file.Name] -ne $fileHash) {
        throw "Release asset checksum mismatch: $($file.Name)"
    }
}
if ($checksumText -notmatch "(?im)^$($actualHash.ToLowerInvariant())\s+ExcelDiffTracker-Setup-arm64\.exe$") {
    throw 'SHA256SUMS.txt does not describe the installer being published.'
}
if ($manifestText -notmatch "(?im)^\s*InstallerSha256:\s*$actualHash\s*$") {
    throw 'The WinGet manifest hash does not match the installer being published.'
}

$dirty = & git -C $repositoryRoot status --porcelain
if ($LASTEXITCODE -ne 0 -or $dirty) { throw 'Commit all release metadata before publishing.' }
& gh auth status
if ($LASTEXITCODE -ne 0) { throw 'GitHub CLI authentication is required.' }

$tag = "v$Version"
& git -C $repositoryRoot tag -a $tag -m "Excel Diff Tracker $Version"
if ($LASTEXITCODE -ne 0) { throw "Could not create tag $tag." }
& git -C $repositoryRoot push origin $tag
if ($LASTEXITCODE -ne 0) { throw "Could not push tag $tag." }

$files = @($releaseFiles + (Get-Item $checksumFile) | Select-Object -ExpandProperty FullName)
$arguments = @('release', 'create', $tag, '--repo', $Repository, '--title', "Excel Diff Tracker $Version", '--notes-file', $releaseNotes)
if ($Prerelease) { $arguments += '--prerelease' }
$arguments += $files
& gh @arguments
if ($LASTEXITCODE -ne 0) { throw 'GitHub release creation failed.' }

Write-Host "Published $Repository release $tag with installer SHA-256 $actualHash"
