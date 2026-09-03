[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string] $IdentityPath,
    [Parameter(Mandatory)] [string] $SourceManifestPath,
    [Parameter(Mandatory)] [string] $RepositoryRoot,
    [Parameter(Mandatory)] [ValidatePattern('^\d+\.\d+\.\d+([-.][0-9A-Za-z.-]+)?$')] [string] $ExpectedVersion,
    [Parameter(Mandatory)] [ValidatePattern('^[A-Fa-f0-9]{40}$')] [string] $ExpectedSourceCommit
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$identityPath = (Resolve-Path $IdentityPath).Path
$sourceManifestPath = (Resolve-Path $SourceManifestPath).Path
$repositoryRoot = (Resolve-Path $RepositoryRoot).Path
$identity = Get-Content -LiteralPath $identityPath -Raw | ConvertFrom-Json
$expectedCommit = $ExpectedSourceCommit.ToLowerInvariant()

function Require-Condition {
    param([bool] $Condition, [string] $Message)
    if (-not $Condition) { throw "Invalid candidate build identity: $Message" }
}

Require-Condition ($identity.schemaVersion -eq 1 -and $identity.product -eq 'Excel Scenario Analysis Tool') 'schema or product identity differs'
Require-Condition ($identity.version -eq $ExpectedVersion) 'version differs from the frozen candidate version'
Require-Condition ($identity.sourceCommit -eq $expectedCommit) 'source commit differs from the frozen build commit'
Require-Condition ($identity.sourceManifest -eq 'BUILD-SOURCE-SHA256SUMS.txt') 'source manifest filename differs'
$expectedGeneratedPath = "packaging/winget/yogoldy.ExcelDiffTracker/$ExpectedVersion/yogoldy.ExcelDiffTracker.installer.yaml"
$hasGeneratedManifest = -not [string]::IsNullOrWhiteSpace([string]$identity.excludedGeneratedPath)
if ($hasGeneratedManifest) {
    Require-Condition ($identity.excludedGeneratedPath -eq $expectedGeneratedPath) 'source-manifest exclusion is not the one generated WinGet hash manifest'
}

$manifestHash = (Get-FileHash -LiteralPath $sourceManifestPath -Algorithm SHA256).Hash.ToUpperInvariant()
Require-Condition ($identity.sourceManifestSha256 -eq $manifestHash) 'source manifest hash differs'
$createdUtc = if ($identity.createdUtc -is [DateTime]) {
    ([DateTime]$identity.createdUtc).ToUniversalTime()
} else {
    [DateTime]::Parse(
        [string]$identity.createdUtc,
        [Globalization.CultureInfo]::InvariantCulture,
        [Globalization.DateTimeStyles]::RoundtripKind).ToUniversalTime()
}
Require-Condition ($createdUtc -le [DateTime]::UtcNow.AddMinutes(5)) 'build timestamp is missing or invalid'

$headCommit = (& git -C $repositoryRoot rev-parse HEAD).Trim().ToLowerInvariant()
Require-Condition ($LASTEXITCODE -eq 0 -and $headCommit -match '^[a-f0-9]{40}$') 'current release commit cannot be resolved'
& git -C $repositoryRoot cat-file -e "$expectedCommit^{commit}" 2>$null
Require-Condition ($LASTEXITCODE -eq 0) 'frozen build commit is not present in the repository'
& git -C $repositoryRoot merge-base --is-ancestor $expectedCommit $headCommit
Require-Condition ($LASTEXITCODE -eq 0) 'frozen build commit is not an ancestor of the release commit'
$changedAfterBuild = @(& git -C $repositoryRoot diff --name-only "$expectedCommit..$headCommit" --)
Require-Condition ($LASTEXITCODE -eq 0) 'build-to-release source diff cannot be read'
if ($headCommit -ne $expectedCommit) {
    Require-Condition ($changedAfterBuild.Count -eq 1 -and $changedAfterBuild[0] -eq $expectedGeneratedPath) 'release commit changed files other than the generated WinGet installer hash'
} else {
    Require-Condition ($changedAfterBuild.Count -eq 0) 'source changed without a distinct release commit'
}

$dirtySource = @(& git -C $repositoryRoot status --porcelain --untracked-files=all)
Require-Condition ($LASTEXITCODE -eq 0 -and $dirtySource.Count -eq 0) 'release checkout is dirty'
$trackedFiles = @(& git -C $repositoryRoot ls-files)
Require-Condition ($LASTEXITCODE -eq 0 -and $trackedFiles.Count -gt 0) 'tracked source files cannot be enumerated'
$includedFiles = if ($hasGeneratedManifest) {
    @($trackedFiles | Where-Object { $_ -ne $expectedGeneratedPath } | Sort-Object)
} else {
    @($trackedFiles | Sort-Object)
}
if ($hasGeneratedManifest) {
    Require-Condition ($includedFiles.Count -eq ($trackedFiles.Count - 1)) 'generated WinGet manifest is not the one and only exclusion'
} else {
    Require-Condition ($includedFiles.Count -eq $trackedFiles.Count) 'source files were excluded without a generated WinGet manifest'
}
$lines = foreach ($relative in $includedFiles) {
    $fullPath = Join-Path $repositoryRoot $relative.Replace('/', [System.IO.Path]::DirectorySeparatorChar)
    Require-Condition (Test-Path $fullPath -PathType Leaf) "tracked source file is missing: $relative"
    $hash = (Get-FileHash -LiteralPath $fullPath -Algorithm SHA256).Hash.ToLowerInvariant()
    "$hash  $relative"
}
$expectedManifest = (($lines -join "`n") + "`n")
$actualManifest = [System.IO.File]::ReadAllText($sourceManifestPath)
Require-Condition ($actualManifest -ceq $expectedManifest) 'installed source manifest does not exactly match the release checkout'
Require-Condition ([int]$identity.sourceFileCount -eq $includedFiles.Count) 'source file count differs'

Write-Output "CANDIDATE_BUILD_IDENTITY_VALID|sourceCommit=$expectedCommit|releaseCommit=$headCommit|manifestSha256=$manifestHash"
