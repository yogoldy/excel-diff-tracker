[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string] $PolicyPath,
    [Parameter(Mandatory)] [string] $RepositoryRoot,
    [Parameter(Mandatory)] [ValidatePattern('^\d+\.\d+\.\d+$')] [string] $CandidateVersion,
    [Parameter(Mandatory)] [string] $PriorInstallerPath,
    [Parameter(Mandatory)] [ValidatePattern('^[A-Fa-f0-9]{64}$')] [string] $ExpectedPriorApplicationSha256,
    [Parameter(Mandatory)] [ValidatePattern('^\d+\.\d+\.\d+$')] [string] $PriorVersion
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$policyPath = (Resolve-Path $PolicyPath).Path
$repositoryRoot = (Resolve-Path $RepositoryRoot).Path
$priorInstaller = (Resolve-Path $PriorInstallerPath).Path
$policy = Get-Content -LiteralPath $policyPath -Raw | ConvertFrom-Json

function Require-Condition {
    param([bool] $Condition, [string] $Message)
    if (-not $Condition) { throw "Invalid release-baseline policy: $Message" }
}

Require-Condition ($policy.schemaVersion -eq 1 -and $policy.repository -eq 'yogoldy/excel-diff-tracker') 'schema or repository identity differs'
$candidateProperty = $policy.candidatePolicies.PSObject.Properties[$CandidateVersion]
Require-Condition ($null -ne $candidateProperty) "candidate $CandidateVersion has no pinned prior-release policy"
$baseline = $candidateProperty.Value
$priorInstallerHash = (Get-FileHash -LiteralPath $priorInstaller -Algorithm SHA256).Hash.ToUpperInvariant()
$expectedPriorApplicationHash = $ExpectedPriorApplicationSha256.ToUpperInvariant()
Require-Condition ($PriorVersion -eq $baseline.requiredPriorVersion) 'caller-supplied prior version differs from policy'
Require-Condition ([version]$CandidateVersion -gt [version]$PriorVersion) 'candidate is not newer than the pinned prior version'
Require-Condition ($priorInstallerHash -eq $baseline.priorInstallerSha256) 'prior installer bytes are not the pinned public release asset'
Require-Condition ($expectedPriorApplicationHash -eq $baseline.priorInstalledApplicationSha256) 'prior installed-application hash differs from policy'
Require-Condition ($baseline.priorTag -eq "v$PriorVersion") 'prior tag does not match the pinned prior version'
Require-Condition ($baseline.priorReleaseUrl -eq "https://github.com/yogoldy/excel-diff-tracker/releases/tag/$($baseline.priorTag)") 'prior release URL differs from the canonical GitHub release'
Require-Condition ($baseline.priorInstallerAssetUrl -eq "https://github.com/yogoldy/excel-diff-tracker/releases/download/$($baseline.priorTag)/ExcelDiffTracker-Setup-arm64.exe") 'prior installer URL differs from the canonical GitHub asset'
Require-Condition ($baseline.priorSourceCommit -match '^[a-f0-9]{40}$') 'prior source commit is malformed'
$tagCommit = (& git -C $repositoryRoot rev-list -n 1 $baseline.priorTag).Trim().ToLowerInvariant()
Require-Condition ($LASTEXITCODE -eq 0 -and $tagCommit -eq $baseline.priorSourceCommit) 'local prior tag does not resolve to the pinned public source commit'

[pscustomobject]@{
    status = 'Passed'
    candidateVersion = $CandidateVersion
    priorVersion = $PriorVersion
    priorTag = $baseline.priorTag
    priorSourceCommit = $baseline.priorSourceCommit
    priorInstallerSha256 = $priorInstallerHash
    priorApplicationSha256 = $expectedPriorApplicationHash
    policySha256 = (Get-FileHash -LiteralPath $policyPath -Algorithm SHA256).Hash.ToUpperInvariant()
}
