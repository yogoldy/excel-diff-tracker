[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string] $EvidenceRoot,
    [Parameter(Mandatory)] [ValidatePattern('^[A-Fa-f0-9]{64}$')] [string] $ExpectedInstallerSha256,
    [Parameter(Mandatory)] [ValidatePattern('^[A-Fa-f0-9]{64}$')] [string] $ExpectedApplicationSha256,
    [Parameter(Mandatory)] [ValidateNotNullOrEmpty()] [string] $Reviewer,
    [Parameter(Mandatory)] [ValidateNotNullOrEmpty()] [string] $Notes,
    [Parameter(Mandatory)] [switch] $ApproveContrast,
    [Parameter(Mandatory)] [switch] $ApproveClippingAndOverlap,
    [Parameter(Mandatory)] [switch] $ApproveKeyboardAndFocus,
    [Parameter(Mandatory)] [switch] $ApproveThemeTransitions
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'UiAutomation.psm1') -Force

if (-not $ApproveContrast -or -not $ApproveClippingAndOverlap -or -not $ApproveKeyboardAndFocus -or -not $ApproveThemeTransitions) {
    throw 'Human review is fail-closed: contrast, clipping/overlap, keyboard/focus, and live theme-transition rendering must each be explicitly approved.'
}
if ($Reviewer.Trim().Length -lt 3) { throw 'Name the human visual/accessibility reviewer.' }
if ($Notes.Trim().Length -lt 20) { throw 'Review notes must describe what was checked (minimum 20 characters).' }

$root = (Resolve-Path $EvidenceRoot).Path
$validator = (Resolve-Path (Join-Path $PSScriptRoot 'Test-VisualMatrixBundle.ps1')).Path
$preApproval = & $validator -EvidenceRoot $root -ExpectedInstallerSha256 $ExpectedInstallerSha256 -ExpectedApplicationSha256 $ExpectedApplicationSha256
if ($preApproval.status -ne 'MachineChecksPassedHumanReviewRequired' -and $preApproval.status -ne 'Approved') { throw 'Visual bundle did not pass its machine gate.' }

$contrastPaths = @(Get-ChildItem $root -Filter 'visual-contrast.json' -File -Recurse)
$lifecyclePaths = @(Get-ChildItem $root -Filter 'visual-lifecycle.json' -File -Recurse)
if ($contrastPaths.Count -ne 1 -or $lifecyclePaths.Count -ne 1) { throw 'Exactly one contrast and one lifecycle artifact are required before human approval.' }
$contrast = Get-Content $contrastPaths[0].FullName -Raw | ConvertFrom-Json
$lifecycle = Get-Content $lifecyclePaths[0].FullName -Raw | ConvertFrom-Json
$latestCaptureUtc = ([DateTime]$contrast.capturedUtc).ToUniversalTime()
foreach ($timestamp in @($lifecycle.capturedUtc) + @($lifecycle.states | ForEach-Object { $_.capturedUtc })) {
    $captureUtc = ([DateTime]$timestamp).ToUniversalTime()
    if ($captureUtc -gt $latestCaptureUtc) { $latestCaptureUtc = $captureUtc }
}
$reviewedUtc = [DateTime]::UtcNow.ToString('O')
if (([DateTime]$reviewedUtc).ToUniversalTime() -lt $latestCaptureUtc) { throw 'Human approval cannot predate contrast or lifecycle capture.' }
$contrastArtifactHash = Get-AcceptanceFileSha256 $contrastPaths[0].FullName
$lifecycleArtifactHash = Get-AcceptanceFileSha256 $lifecyclePaths[0].FullName
$matrixPaths = @(Get-ChildItem $root -Filter 'visual-matrix.json' -File -Recurse | Sort-Object FullName)
foreach ($matrixPath in $matrixPaths) {
    $matrix = Get-Content $matrixPath.FullName -Raw | ConvertFrom-Json
    $matrix.status = 'Approved'
    $matrix.humanReview.required = $true
    $matrix.humanReview.contrast = 'Approved'
    $matrix.humanReview.clippingAndOverlap = 'Approved'
    $matrix.humanReview.keyboardAndFocus = 'Approved'
    if ($null -eq $matrix.humanReview.PSObject.Properties['themeTransitions']) { $matrix.humanReview | Add-Member -NotePropertyName themeTransitions -NotePropertyValue 'Approved' }
    else { $matrix.humanReview.themeTransitions = 'Approved' }
    foreach ($binding in @(
        @{ name = 'captureSessionId'; value = $preApproval.captureSessionId },
        @{ name = 'contrastArtifactSha256'; value = $contrastArtifactHash },
        @{ name = 'lifecycleArtifactSha256'; value = $lifecycleArtifactHash })) {
        $matrix.humanReview | Add-Member -NotePropertyName $binding.name -NotePropertyValue $binding.value -Force
    }
    $matrix.humanReview.reviewer = $Reviewer.Trim()
    $matrix.humanReview.reviewedUtc = $reviewedUtc
    $matrix.humanReview.notes = $Notes.Trim()
    Write-AcceptanceUtf8File -Path $matrixPath.FullName -Content ($matrix | ConvertTo-Json -Depth 40)
}

$result = & $validator -EvidenceRoot $root -ExpectedInstallerSha256 $ExpectedInstallerSha256 -ExpectedApplicationSha256 $ExpectedApplicationSha256 -RequireHumanApproval
if ($result.status -ne 'Approved') { throw 'Visual bundle did not remain valid after approval was recorded.' }
$result
