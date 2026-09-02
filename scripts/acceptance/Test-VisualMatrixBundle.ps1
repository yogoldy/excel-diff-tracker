[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string] $EvidenceRoot,
    [Parameter(Mandatory)] [ValidatePattern('^[A-Fa-f0-9]{64}$')] [string] $ExpectedInstallerSha256,
    [Parameter(Mandatory)] [ValidatePattern('^[A-Fa-f0-9]{64}$')] [string] $ExpectedApplicationSha256,
    [switch] $RequireHumanApproval
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'UiAutomation.psm1') -Force

$root = (Resolve-Path $EvidenceRoot).Path
$installerHash = $ExpectedInstallerSha256.ToUpperInvariant()
$applicationHash = $ExpectedApplicationSha256.ToUpperInvariant()
$matrixPaths = @(Get-ChildItem $root -Filter 'visual-matrix.json' -File -Recurse | Sort-Object FullName)
if ($matrixPaths.Count -eq 0) { throw "No visual-matrix.json evidence exists under $root." }

function Resolve-EvidenceFile {
    param([string] $MatrixDirectory, [string] $RelativePath)
    if ([string]::IsNullOrWhiteSpace($RelativePath) -or [System.IO.Path]::IsPathRooted($RelativePath)) {
        throw "Evidence path must be non-empty and relative: '$RelativePath'."
    }
    $candidate = [System.IO.Path]::GetFullPath((Join-Path $MatrixDirectory ($RelativePath.Replace('/', [System.IO.Path]::DirectorySeparatorChar))))
    $trimCharacters = [char[]]@([System.IO.Path]::DirectorySeparatorChar,[System.IO.Path]::AltDirectorySeparatorChar)
    $prefix = [System.IO.Path]::GetFullPath($MatrixDirectory).TrimEnd($trimCharacters) + [System.IO.Path]::DirectorySeparatorChar
    if (-not $candidate.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) { throw "Evidence path escapes its matrix directory: $RelativePath" }
    if (-not (Test-Path $candidate -PathType Leaf)) { throw "Visual evidence file is missing: $candidate" }
    $candidate
}

function Assert-FileHash {
    param([string] $Path, [string] $ExpectedHash)
    if ($ExpectedHash -notmatch '^[A-Fa-f0-9]{64}$') { throw "Visual evidence has a malformed SHA-256: $Path" }
    $actual = Get-AcceptanceFileSha256 -Path $Path
    if ($actual -ne $ExpectedHash.ToUpperInvariant()) { throw "Visual evidence hash mismatch: $Path" }
}

function Assert-CategoriesPresent {
    param([object[]] $Results, [string[]] $Categories)
    foreach ($category in $Categories) {
        if (-not @($Results | Where-Object { $_.stateCategory -eq $category }).Count) { throw "Visual evidence is missing required state category '$category'." }
    }
}

function Test-GeometryInsideWorkingArea {
    param([object] $Geometry)
    if ($null -eq $Geometry -or $null -eq $Geometry.actual -or $null -eq $Geometry.monitorWorkingArea) { return $false }
    $actual = $Geometry.actual; $work = $Geometry.monitorWorkingArea
    [double]$actual.width -gt 0 -and [double]$actual.height -gt 0 -and [double]$actual.x -ge [double]$work.x -and [double]$actual.y -ge [double]$work.y -and [double]$actual.right -le [double]$work.right -and [double]$actual.bottom -le [double]$work.bottom
}

$matrices = [System.Collections.Generic.List[object]]::new()
$allResults = [System.Collections.Generic.List[object]]::new()
$sessionIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$automationIdentity = @{}
$humanReviewers = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
$humanReviewTimes = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
foreach ($matrixPath in $matrixPaths) {
    $matrixDirectory = Split-Path -Parent $matrixPath.FullName
    $matrix = Get-Content $matrixPath.FullName -Raw | ConvertFrom-Json
    if ($matrix.schemaVersion -ne 2) { throw "Unsupported visual matrix schema: $($matrixPath.FullName)" }
    if ($matrix.captureSessionId -notmatch '^[0-9a-fA-F-]{36}$') { throw "Visual matrix has no valid capture-session identity: $($matrixPath.FullName)" }
    [void]$sessionIds.Add($matrix.captureSessionId)
    if ($matrix.installerSha256 -ne $installerHash -or $matrix.applicationSha256 -ne $applicationHash) { throw "Visual matrix candidate identity does not match the frozen candidate: $($matrixPath.FullName)" }
    if ([string]::IsNullOrWhiteSpace($matrix.installerPathAtCapture) -or [string]::IsNullOrWhiteSpace($matrix.applicationPath)) { throw "Visual matrix did not record the exact installer and installed executable paths used at capture time: $($matrixPath.FullName)" }
    if ($matrix.status -eq 'Failed') { throw "Visual matrix contains a retained machine-check failure: $($matrixPath.FullName)" }
    if ($RequireHumanApproval) {
        if ($matrix.status -ne 'Approved' -or $matrix.humanReview.required -ne $true -or $matrix.humanReview.contrast -ne 'Approved' -or $matrix.humanReview.clippingAndOverlap -ne 'Approved' -or $matrix.humanReview.keyboardAndFocus -ne 'Approved' -or [string]::IsNullOrWhiteSpace($matrix.humanReview.reviewer) -or [string]::IsNullOrWhiteSpace($matrix.humanReview.notes) -or [string]::IsNullOrWhiteSpace($matrix.humanReview.reviewedUtc)) {
            throw "Visual matrix human contrast/clipping/focus review is incomplete: $($matrixPath.FullName)"
        }
        if ($matrix.humanReview.reviewer.Trim().Length -lt 3 -or $matrix.humanReview.notes.Trim().Length -lt 20) { throw "Visual matrix human reviewer identity or notes were weakened after approval: $($matrixPath.FullName)" }
        $reviewedAt = [DateTime]::MinValue
        if (-not [DateTime]::TryParse($matrix.humanReview.reviewedUtc,[ref]$reviewedAt) -or $reviewedAt.ToUniversalTime() -lt ([DateTime]$matrix.capturedUtc).ToUniversalTime() -or $reviewedAt.ToUniversalTime() -gt [DateTime]::UtcNow.AddMinutes(5)) { throw "Visual matrix human approval time is invalid or predates capture: $($matrixPath.FullName)" }
        [void]$humanReviewers.Add($matrix.humanReview.reviewer.Trim())
        [void]$humanReviewTimes.Add($reviewedAt.ToUniversalTime().ToString('O'))
    }
    elseif ($matrix.status -notin @('MachineChecksPassedHumanReviewRequired','Approved')) {
        throw "Visual matrix is not ready for review: $($matrixPath.FullName)"
    }
    if ($matrix.expectedScalePercent -ne $matrix.actualEnvironment.scalePercent) { throw "Visual matrix trusted a requested scale that differs from measured DPI: $($matrixPath.FullName)" }
    $derivedScale = [int][Math]::Round(([double]$matrix.actualEnvironment.windowDpi / 96.0) * 100.0)
    if ($derivedScale -ne [int]$matrix.actualEnvironment.scalePercent) { throw "Visual matrix scale and window DPI are inconsistent: $($matrixPath.FullName)" }
    if ([bool]$matrix.windowsContrastThemeRequested -ne [bool]$matrix.actualEnvironment.highContrast) { throw "Visual matrix contrast label differs from measured Windows high-contrast state: $($matrixPath.FullName)" }
    if ($matrix.actualEnvironment.windowsAppTheme -notin @('Light','Dark') -or $matrix.actualEnvironment.windowsSystemTheme -notin @('Light','Dark')) { throw "Visual matrix is missing measured Windows theme state: $($matrixPath.FullName)" }
    $primary = @($matrix.actualEnvironment.displays | Where-Object { $_.primary })
    if ($primary.Count -ne 1 -or [int]$primary[0].physicalWidth -le 0 -or [int]$primary[0].physicalHeight -le 0) { throw "Visual matrix is missing the measured primary physical resolution: $($matrixPath.FullName)" }
    if (@($matrix.results).Count -eq 0) { throw "Visual matrix has no captured states: $($matrixPath.FullName)" }
    $firstCaptured = [DateTime]::MinValue
    $lastCaptured = [DateTime]::MinValue
    if (-not [DateTime]::TryParse($matrix.firstCapturedUtc, [ref]$firstCaptured) -or -not [DateTime]::TryParse($matrix.capturedUtc, [ref]$lastCaptured) -or $firstCaptured -gt $lastCaptured -or $lastCaptured.ToUniversalTime() -gt [DateTime]::UtcNow.AddMinutes(5)) { throw "Visual matrix timestamps are missing or inconsistent: $($matrixPath.FullName)" }
    $stateNames = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $referencedFiles = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($result in @($matrix.results)) {
        if (-not $stateNames.Add($result.state)) { throw "Duplicate state overwrote or shadowed fresh evidence: $($result.state) in $($matrixPath.FullName)" }
        if ($result.passed -ne $true -or @($result.geometryFailures).Count -ne 0 -or @($result.overlapFailures).Count -ne 0 -or @($result.accessibilityFailures).Count -ne 0 -or @($result.focusFailures).Count -ne 0 -or @($result.semanticFailures).Count -ne 0) { throw "Visual state has a failed machine assertion: $($result.state) in $($matrixPath.FullName)" }
        $tested = $result.testedWindowGeometry
        if ($null -eq $tested -or $tested.mode -ne $result.windowSizeMode -or $tested.verified -ne $true -or $tested.fullyInsideMonitorWorkingArea -ne $true -or -not (Test-GeometryInsideWorkingArea $tested)) { throw "Visual state trusts an unverified requested window-size label: $($result.state) in $($matrixPath.FullName)" }
        if ($result.windowSizeMode -eq 'Default') {
            if ($null -ne $tested.requestedWidth -or $null -ne $tested.requestedHeight) { throw "Default window evidence must be measured rather than assigned a synthetic requested size: $($result.state)" }
        }
        else {
            if ([double]$tested.requestedWidth -le 0 -or [double]$tested.requestedHeight -le 0 -or [Math]::Abs([double]$tested.actual.width - [double]$tested.requestedWidth) -gt 12 -or [Math]::Abs([double]$tested.actual.height - [double]$tested.requestedHeight) -gt 12) { throw "Measured window dimensions do not match the requested $($result.windowSizeMode) geometry: $($result.state)" }
        }
        if ($result.stateCategory -eq 'tray-menu') {
            if ($null -ne $result.rootWindowGeometry) { throw "Desktop-root tray evidence should not claim a root-window geometry: $($result.state)" }
        }
        elseif ($null -eq $result.rootWindowGeometry -or -not (Test-GeometryInsideWorkingArea $result.rootWindowGeometry) -or $result.rootWindowGeometry.fullyInsideMonitorWorkingArea -ne $true) { throw "Captured root window is missing or partly outside its monitor: $($result.state) in $($matrixPath.FullName)" }
        if ($result.stateCategory -in @('main-page','onboarding','toast','processing','warning','error','long-path')) {
            if ([Math]::Abs([double]$result.rootWindowGeometry.actual.width - [double]$tested.actual.width) -gt 12 -or [Math]::Abs([double]$result.rootWindowGeometry.actual.height - [double]$tested.actual.height) -gt 12) { throw "The product window drifted from its verified $($result.windowSizeMode) size before state capture: $($result.state)" }
        }
        if ($result.phase -eq 'Supplemental') {
            if ($result.stateSemanticVerified -ne $true -or [string]::IsNullOrWhiteSpace($result.observedStateText)) { throw "Supplemental state is only labeled, not semantically verified: $($result.state)" }
            if ($result.stateCategory -eq 'processing' -and $result.observedStateText -ne 'Processing') { throw "Processing evidence does not contain the exact product status: $($result.state)" }
            if ($result.stateCategory -eq 'warning' -and $result.observedStateText -ne 'Warning') { throw "Warning evidence does not contain the exact product status: $($result.state)" }
            if ($result.stateCategory -eq 'error') {
                $allowedErrorPrefixes = @('Missing workbook','Unsupported workbook','Corrupt, encrypted, or unsafe workbook','Permission denied','Workbook temporarily unavailable','Capture failed')
                if (-not @($allowedErrorPrefixes | Where-Object { $result.observedStateText.StartsWith($_,[StringComparison]::Ordinal) }).Count) { throw "Error evidence does not contain a recognized exact capture-error category: $($result.state)" }
            }
        }
        foreach ($interactive in @($result.interactive)) {
            if ([string]::IsNullOrWhiteSpace($interactive.automationId) -or [string]::IsNullOrWhiteSpace($interactive.name) -or [string]::IsNullOrWhiteSpace($interactive.role) -or $null -eq $interactive.state) { throw "Interactive element identity/state is incomplete in $($result.state): $($matrixPath.FullName)" }
            if ($interactive.automationId -notmatch '^\d+$') {
                $identity = "$($interactive.name)|$($interactive.role)"
                if ($automationIdentity.ContainsKey($interactive.automationId) -and $automationIdentity[$interactive.automationId] -ne $identity) { throw "AutomationId '$($interactive.automationId)' is not stable: '$($automationIdentity[$interactive.automationId])' versus '$identity'." }
                $automationIdentity[$interactive.automationId] = $identity
            }
        }
        $screenshot = Resolve-EvidenceFile $matrixDirectory $result.screenshot; Assert-FileHash $screenshot $result.screenshotSha256; [void]$referencedFiles.Add($screenshot)
        $tree = Resolve-EvidenceFile $matrixDirectory $result.uiaTree; Assert-FileHash $tree $result.uiaTreeSha256; [void]$referencedFiles.Add($tree)
        $focusPath = Resolve-EvidenceFile $matrixDirectory $result.focusEvidence; Assert-FileHash $focusPath $result.focusEvidenceSha256; [void]$referencedFiles.Add($focusPath)
        $focus = Get-Content $focusPath -Raw | ConvertFrom-Json
        if (@($focus.failures).Count -ne 0) { throw "Focus evidence contains failures: $focusPath" }
        if ($result.stateCategory -in @('onboarding','main-page','workbook-picker','folder-picker','confirmation-dialog') -and @($focus.stops).Count -eq 0) { throw "Required keyboard-focus evidence has no Tab stops: $focusPath" }
        foreach ($stop in @($focus.stops)) {
            if ($stop.hasKeyboardFocus -ne $true -or $stop.state.offscreen -eq $true -or [string]::IsNullOrWhiteSpace($stop.name) -or [string]::IsNullOrWhiteSpace($stop.role) -or $null -eq $stop.state) { throw "Invalid focus stop in $focusPath" }
            $focusPng = Resolve-EvidenceFile $matrixDirectory $stop.screenshot; Assert-FileHash $focusPng $stop.screenshotSha256; [void]$referencedFiles.Add($focusPng)
        }
        $allResults.Add([pscustomobject]@{ matrixPath = $matrixPath.FullName; matrix = $matrix; result = $result })
    }
    $capturedFiles = @(Get-ChildItem $matrixDirectory -File -Recurse | Where-Object { $_.Name -ne 'visual-matrix.json' })
    if ($capturedFiles.Count -ne $referencedFiles.Count) { throw "Visual matrix has orphaned or unreferenced capture files (actual $($capturedFiles.Count), referenced $($referencedFiles.Count)): $($matrixPath.FullName)" }
    foreach ($capturedFile in $capturedFiles) { if (-not $referencedFiles.Contains($capturedFile.FullName)) { throw "Visual matrix contains an unreferenced capture file: $($capturedFile.FullName)" } }
    $matrices.Add($matrix)
}

if ($sessionIds.Count -ne 1) { throw "Visual evidence mixes $($sessionIds.Count) capture-session IDs; use one fresh evidence session." }
if ($RequireHumanApproval -and ($humanReviewers.Count -ne 1 -or $humanReviewTimes.Count -ne 1)) { throw 'All matrices must carry the same independent reviewer identity and one approval timestamp from the bundle approval operation.' }

$requiredScales = @(100,125,150,200)
$actualScales = @($matrices | ForEach-Object { [int]$_.actualEnvironment.scalePercent } | Select-Object -Unique)
foreach ($scale in $requiredScales) { if ($actualScales -notcontains $scale) { throw "Visual evidence is missing measured $scale% DPI scaling." } }

$primaryModes = @($matrices | ForEach-Object { @($_.actualEnvironment.displays | Where-Object { $_.primary })[0] })
if (-not @($primaryModes | Where-Object { [int]$_.physicalWidth -le 1280 -and [int]$_.physicalHeight -le 720 }).Count) { throw 'Visual evidence is missing a measured 1280x720-class display.' }
if (-not @($primaryModes | Where-Object { [int]$_.physicalWidth -ge 1920 -and [int]$_.physicalHeight -ge 1080 }).Count) { throw 'Visual evidence is missing a measured 1920x1080-class display.' }
if (-not @($matrices | Where-Object { $_.actualEnvironment.highContrast }).Count) { throw 'Visual evidence is missing measured Windows contrast-theme coverage.' }
foreach ($windowsTheme in @('Light','Dark')) {
    if (-not @($matrices | Where-Object { -not $_.actualEnvironment.highContrast -and $_.actualEnvironment.windowsAppTheme -eq $windowsTheme }).Count) { throw "Visual evidence is missing measured Windows system-$($windowsTheme.ToLowerInvariant()) coverage." }
}
foreach ($sizeMode in @('Minimum','Default','Resized')) {
    if (-not @($allResults | Where-Object { $_.result.stateCategory -eq 'main-page' -and $_.result.windowSizeMode -eq $sizeMode }).Count) { throw "Visual evidence is missing $sizeMode main-window coverage." }
}
$defaultSizes = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
$resizedSizes = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
foreach ($entry in @($allResults | Where-Object { $_.result.stateCategory -eq 'main-page' -and $_.result.windowSizeMode -in @('Default','Resized') })) {
    $size = "$([int]$entry.result.testedWindowGeometry.actual.width)x$([int]$entry.result.testedWindowGeometry.actual.height)"
    if ($entry.result.windowSizeMode -eq 'Default') { [void]$defaultSizes.Add($size) } else { [void]$resizedSizes.Add($size) }
}
$distinctResized = @($resizedSizes | Where-Object { -not $defaultSizes.Contains($_) })
if ($distinctResized.Count -eq 0) { throw 'Resized coverage is not measurably distinct from Default window coverage.' }

$resultsOnly = @($allResults | ForEach-Object { $_.result })
foreach ($state in @('onboarding-step-1','onboarding-step-2','onboarding-step-3-empty','onboarding-step-3-long-path','onboarding-step-4','onboarding-step-5')) {
    if (-not @($resultsOnly | Where-Object { $_.state -eq $state }).Count) { throw "Visual evidence is missing required onboarding state '$state'." }
}
foreach ($theme in @('Light','Dark','System')) {
    foreach ($page in @('dashboard','workbooks','history','settings','about')) {
        if (-not @($resultsOnly | Where-Object { $_.stateCategory -eq 'main-page' -and $_.applicationTheme -eq $theme -and $_.state -match "-$page-" }).Count) { throw "Visual evidence is missing $theme $page coverage." }
    }
}
foreach ($scale in $requiredScales) {
    $scaleResults = @($allResults | Where-Object { [int]$_.matrix.actualEnvironment.scalePercent -eq $scale } | ForEach-Object { $_.result })
    foreach ($theme in @('Light','Dark','System')) {
        foreach ($page in @('dashboard','workbooks','history','settings','about')) {
            if (-not @($scaleResults | Where-Object { $_.stateCategory -eq 'main-page' -and $_.applicationTheme -eq $theme -and $_.state -match "-$page-" }).Count) { throw "Measured $scale% evidence is missing $theme $page coverage." }
        }
    }
}
foreach ($windowsTheme in @('Light','Dark')) {
    $systemResults = @($allResults | Where-Object { -not $_.matrix.actualEnvironment.highContrast -and $_.matrix.actualEnvironment.windowsAppTheme -eq $windowsTheme } | ForEach-Object { $_.result })
    foreach ($page in @('dashboard','workbooks','history','settings','about')) {
        if (-not @($systemResults | Where-Object { $_.stateCategory -eq 'main-page' -and $_.applicationTheme -eq 'System' -and $_.state -match "-$page-" }).Count) { throw "Measured Windows system-$($windowsTheme.ToLowerInvariant()) evidence is missing the $page page." }
    }
}
$smallDisplayResults = @($allResults | Where-Object {
    $primary = @($_.matrix.actualEnvironment.displays | Where-Object { $_.primary })[0]
    [int]$primary.physicalWidth -le 1280 -and [int]$primary.physicalHeight -le 720
} | ForEach-Object { $_.result })
$largeDisplayResults = @($allResults | Where-Object {
    $primary = @($_.matrix.actualEnvironment.displays | Where-Object { $_.primary })[0]
    [int]$primary.physicalWidth -ge 1920 -and [int]$primary.physicalHeight -ge 1080
} | ForEach-Object { $_.result })
foreach ($displayEntry in @(@{ name = '1280x720-class'; results = $smallDisplayResults },@{ name = '1920x1080-class'; results = $largeDisplayResults })) {
    foreach ($page in @('dashboard','workbooks','history','settings','about')) {
        if (-not @($displayEntry.results | Where-Object { $_.stateCategory -eq 'main-page' -and $_.state -match "-$page-" }).Count) { throw "$($displayEntry.name) evidence is missing the $page page." }
    }
}
foreach ($page in @('dashboard','workbooks','history','settings','about')) {
    if (-not @($resultsOnly | Where-Object { $_.stateCategory -eq 'main-page' -and $_.applicationTheme -eq 'System' -and $_.highContrast -eq $true -and $_.state -match "-$page-" }).Count) { throw "Windows contrast-theme evidence is missing the $page page." }
}
Assert-CategoriesPresent $resultsOnly @('workbook-picker','folder-picker','confirmation-dialog','long-path','tray-menu','toast','processing','warning','error')
if (-not @($resultsOnly | Where-Object { $_.longPathAccessible -eq $true }).Count) { throw 'No visual state proves that the full deliberately long path is exposed to accessibility.' }
if (-not @($resultsOnly | Where-Object { $_.longPathTooltipAccessible -eq $true }).Count) { throw 'No visual state proves that a deliberately truncated path exposes its full accessible HelpText/tooltip.' }
if (-not @($resultsOnly | Where-Object { $_.warningAccessible -eq $true }).Count) { throw 'No visual state proves that the warning text is exposed to accessibility.' }

[pscustomobject]@{
    status = if ($RequireHumanApproval) { 'Approved' } else { 'MachineChecksPassedHumanReviewRequired' }
    captureSessionId = @($sessionIds)[0]
    matrixCount = $matrices.Count
    stateCount = $allResults.Count
    measuredScales = $actualScales
    installerSha256 = $installerHash
    applicationSha256 = $applicationHash
}
