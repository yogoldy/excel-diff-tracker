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

function Assert-ArtifactChecksum {
    param([string] $ArtifactPath)
    $checksumPath = [System.IO.Path]::ChangeExtension($ArtifactPath,'sha256')
    if (-not (Test-Path $checksumPath -PathType Leaf)) { throw "Artifact checksum is missing: $checksumPath" }
    $line = ([System.IO.File]::ReadAllText($checksumPath)).Trim()
    $expectedName = [System.IO.Path]::GetFileName($ArtifactPath)
    if ($line -notmatch ('^(?<hash>[A-Fa-f0-9]{64})  ' + [regex]::Escape($expectedName) + '$')) { throw "Artifact checksum has an invalid format: $checksumPath" }
    Assert-FileHash $ArtifactPath $Matches['hash']
}

function Get-DeterministicPalette {
    param([string] $Content,[string] $Name)
    $block = [regex]::Match($Content,"(?s)${Name}Colors\s*=\s*new\s+Dictionary<string,\s*string>\s*\{(?<body>.*?)\};")
    if (-not $block.Success) { throw "The captured source does not contain the deterministic $Name palette." }
    $palette = [ordered]@{}
    foreach ($entry in [regex]::Matches($block.Groups['body'].Value,'\["(?<key>[^"]+)"\]\s*=\s*"(?<color>#[0-9A-Fa-f]{6})"')) {
        $key = $entry.Groups['key'].Value
        if ($palette.Contains($key)) { throw "Duplicate captured $Name palette key: $key" }
        $palette[$key] = $entry.Groups['color'].Value.ToUpperInvariant()
    }
    $required = @('AppBackgroundBrush','SidebarBrush','CardBrush','CardHoverBrush','TextBrush','MutedTextBrush','BorderBrush','AccentBrush','AccentSoftBrush','WarningBrush','ErrorBrush','PrimaryForegroundBrush')
    foreach ($key in $required) { if (-not $palette.Contains($key)) { throw "The captured $Name palette is missing $key." } }
    if ($palette.Count -ne $required.Count) { throw "The captured $Name palette has an unreviewed color resource." }
    $palette
}

function Get-WcagLuminance {
    param([string] $Color)
    $values = @([Convert]::ToInt32($Color.Substring(1,2),16),[Convert]::ToInt32($Color.Substring(3,2),16),[Convert]::ToInt32($Color.Substring(5,2),16))
    $linear = foreach ($value in $values) { $channel = $value / 255.0; if ($channel -le 0.04045) { $channel / 12.92 } else { [Math]::Pow((($channel + 0.055) / 1.055),2.4) } }
    (0.2126 * $linear[0]) + (0.7152 * $linear[1]) + (0.0722 * $linear[2])
}

function Get-WcagRatio {
    param([string] $Foreground,[string] $Background)
    $first = Get-WcagLuminance $Foreground; $second = Get-WcagLuminance $Background
    ([Math]::Max($first,$second) + 0.05) / ([Math]::Min($first,$second) + 0.05)
}

function New-ContrastContractItem {
    param([string] $Id,[string] $Category,[string] $Foreground,[string] $Background,[double] $Threshold)
    [pscustomobject]@{ id = $Id; category = $Category; foreground = $Foreground; background = $Background; threshold = $Threshold }
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
        if ($matrix.status -ne 'Approved' -or $matrix.humanReview.required -ne $true -or $matrix.humanReview.contrast -ne 'Approved' -or $matrix.humanReview.clippingAndOverlap -ne 'Approved' -or $matrix.humanReview.keyboardAndFocus -ne 'Approved' -or $matrix.humanReview.themeTransitions -ne 'Approved' -or [string]::IsNullOrWhiteSpace($matrix.humanReview.reviewer) -or [string]::IsNullOrWhiteSpace($matrix.humanReview.notes) -or [string]::IsNullOrWhiteSpace($matrix.humanReview.reviewedUtc)) {
            throw "Visual matrix human contrast/clipping/focus/theme-transition review is incomplete: $($matrixPath.FullName)"
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
$captureSessionId = @($sessionIds)[0]

$contrastArtifacts = @(Get-ChildItem $root -Filter 'visual-contrast.json' -File -Recurse)
if ($contrastArtifacts.Count -ne 1) { throw "Expected exactly one hashed visual-contrast artifact; found $($contrastArtifacts.Count)." }
$contrastPath = $contrastArtifacts[0].FullName; Assert-ArtifactChecksum $contrastPath
$contrastDirectory = Split-Path -Parent $contrastPath
$contrast = Get-Content $contrastPath -Raw | ConvertFrom-Json
if ($contrast.schemaVersion -ne 1 -or $contrast.status -ne 'Passed' -or $contrast.failedCount -ne 0) { throw 'Deterministic Light/Dark palette contrast did not pass.' }
$contrastTime = [DateTime]::MinValue
if (-not [DateTime]::TryParse($contrast.capturedUtc,[ref]$contrastTime) -or $contrastTime.ToUniversalTime() -gt [DateTime]::UtcNow.AddMinutes(5)) { throw 'Contrast evidence timestamp is invalid.' }
if ($contrast.captureSessionId -ne $captureSessionId -or $contrast.installerSha256 -ne $installerHash -or $contrast.applicationSha256 -ne $applicationHash) { throw 'Contrast evidence does not belong to the frozen candidate/session.' }
$contrastSource = Resolve-EvidenceFile $contrastDirectory $contrast.source.relativePath
Assert-FileHash $contrastSource $contrast.source.sha256
$sourceContent = [System.IO.File]::ReadAllText($contrastSource)
$parsedPalettes = [ordered]@{ Light = Get-DeterministicPalette $sourceContent 'Light'; Dark = Get-DeterministicPalette $sourceContent 'Dark' }
$contract = @(
    New-ContrastContractItem 'body-on-app' 'normal-text' 'TextBrush' 'AppBackgroundBrush' 4.5
    New-ContrastContractItem 'body-on-sidebar' 'normal-text' 'TextBrush' 'SidebarBrush' 4.5
    New-ContrastContractItem 'body-on-card' 'normal-text' 'TextBrush' 'CardBrush' 4.5
    New-ContrastContractItem 'body-on-hover' 'normal-text' 'TextBrush' 'CardHoverBrush' 4.5
    New-ContrastContractItem 'body-on-accent-soft' 'normal-text' 'TextBrush' 'AccentSoftBrush' 4.5
    New-ContrastContractItem 'muted-on-app' 'normal-text' 'MutedTextBrush' 'AppBackgroundBrush' 4.5
    New-ContrastContractItem 'muted-on-sidebar' 'normal-text' 'MutedTextBrush' 'SidebarBrush' 4.5
    New-ContrastContractItem 'muted-on-card' 'normal-text' 'MutedTextBrush' 'CardBrush' 4.5
    New-ContrastContractItem 'muted-on-hover' 'normal-text' 'MutedTextBrush' 'CardHoverBrush' 4.5
    New-ContrastContractItem 'primary-label' 'normal-text' 'PrimaryForegroundBrush' 'AccentBrush' 4.5
    New-ContrastContractItem 'warning-label' 'normal-text' 'WarningBrush' 'CardBrush' 4.5
    New-ContrastContractItem 'error-label' 'normal-text' 'ErrorBrush' 'CardBrush' 4.5
    New-ContrastContractItem 'page-title' 'large-text' 'TextBrush' 'AppBackgroundBrush' 3.0
    New-ContrastContractItem 'section-title' 'large-text' 'TextBrush' 'CardBrush' 3.0
    New-ContrastContractItem 'border-on-card' 'essential-ui' 'BorderBrush' 'CardBrush' 3.0
    New-ContrastContractItem 'border-on-app' 'essential-ui' 'BorderBrush' 'AppBackgroundBrush' 3.0
    New-ContrastContractItem 'focus-on-card' 'essential-ui' 'AccentBrush' 'CardBrush' 3.0
    New-ContrastContractItem 'focus-on-sidebar' 'essential-ui' 'AccentBrush' 'SidebarBrush' 3.0
    New-ContrastContractItem 'primary-fill-on-app' 'essential-ui' 'AccentBrush' 'AppBackgroundBrush' 3.0
    New-ContrastContractItem 'primary-fill-on-card' 'essential-ui' 'AccentBrush' 'CardBrush' 3.0
    New-ContrastContractItem 'primary-focus-on-fill' 'essential-ui' 'PrimaryForegroundBrush' 'AccentBrush' 3.0
    New-ContrastContractItem 'warning-symbol' 'essential-ui' 'WarningBrush' 'CardBrush' 3.0
    New-ContrastContractItem 'error-symbol' 'essential-ui' 'ErrorBrush' 'CardBrush' 3.0
)
if (@($contrast.checks).Count -ne ($contract.Count * 2)) { throw 'Contrast artifact does not contain the exact reviewed Light/Dark contract.' }
foreach ($theme in @('Light','Dark')) {
    foreach ($key in $parsedPalettes[$theme].Keys) {
        if ($contrast.palettes.$theme.$key -ne $parsedPalettes[$theme][$key]) { throw "Contrast palette value differs from its captured source: $theme/$key" }
    }
    foreach ($item in $contract) {
        $matches = @($contrast.checks | Where-Object { $_.theme -eq $theme -and $_.id -eq $item.id })
        if ($matches.Count -ne 1) { throw "Contrast evidence must contain exactly one $theme/$($item.id) check." }
        $check = $matches[0]; $foreground = $parsedPalettes[$theme][$item.foreground]; $background = $parsedPalettes[$theme][$item.background]
        $ratio = Get-WcagRatio $foreground $background; $rounded = [Math]::Round($ratio,4)
        if ($check.category -ne $item.category -or $check.foregroundResource -ne $item.foreground -or $check.backgroundResource -ne $item.background -or $check.foreground -ne $foreground -or $check.background -ne $background -or [double]$check.threshold -ne $item.threshold -or [Math]::Abs([double]$check.ratio - $rounded) -gt 0.00005 -or $check.passed -ne $true -or $ratio -lt $item.threshold) { throw "WCAG contrast assertion failed or was altered: $theme/$($item.id)." }
    }
}
$contrastFiles = @(Get-ChildItem $contrastDirectory -File)
if ($contrastFiles.Count -ne 3) { throw 'Contrast evidence contains missing or unreviewed files.' }

$lifecycleArtifacts = @(Get-ChildItem $root -Filter 'visual-lifecycle.json' -File -Recurse)
if ($lifecycleArtifacts.Count -ne 1) { throw "Expected exactly one hashed visual-lifecycle artifact; found $($lifecycleArtifacts.Count)." }
$lifecyclePath = $lifecycleArtifacts[0].FullName; Assert-ArtifactChecksum $lifecyclePath
$lifecycleDirectory = Split-Path -Parent $lifecyclePath
$lifecycle = Get-Content $lifecyclePath -Raw | ConvertFrom-Json
if ($lifecycle.schemaVersion -ne 1 -or $lifecycle.status -ne 'MachineObservationPassedHumanRenderReviewRequired' -or $lifecycle.transition.passed -ne $true) { throw 'Visual lifecycle machine observation did not pass.' }
if ($lifecycle.captureSessionId -ne $captureSessionId -or $lifecycle.installerSha256 -ne $installerHash -or $lifecycle.applicationSha256 -ne $applicationHash) { throw 'Lifecycle evidence does not belong to the frozen candidate/session.' }
if ([string]::IsNullOrWhiteSpace($lifecycle.restartMethodLimitation) -or [string]::IsNullOrWhiteSpace($lifecycle.transitionRenderLimitation)) { throw 'Lifecycle evidence concealed its process-termination or rendered-pixel limitation.' }
if (@($lifecycle.restarts).Count -ne 3 -or @($lifecycle.states).Count -ne 9) { throw 'Lifecycle evidence must contain exactly three restarts and nine fresh states.' }
$stateMap = @{}; $lifecycleFiles = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
[void]$lifecycleFiles.Add($lifecyclePath); [void]$lifecycleFiles.Add([System.IO.Path]::ChangeExtension($lifecyclePath,'sha256'))
foreach ($state in @($lifecycle.states)) {
    if ($stateMap.ContainsKey($state.name)) { throw "Duplicate lifecycle state: $($state.name)" }
    $stateTime = [DateTime]::MinValue; $processTime = [DateTime]::MinValue
    if (-not [DateTime]::TryParse($state.capturedUtc,[ref]$stateTime) -or -not [DateTime]::TryParse($state.processStartUtc,[ref]$processTime) -or $processTime.ToUniversalTime() -gt $stateTime.ToUniversalTime() -or $stateTime.ToUniversalTime() -gt [DateTime]::UtcNow.AddMinutes(5)) { throw "Invalid lifecycle timestamp: $($state.name)" }
    if ($state.applicationSha256 -ne $applicationHash -or [int]$state.processId -le 0 -or [string]::IsNullOrWhiteSpace($state.applicationPath)) { throw "Invalid lifecycle process identity: $($state.name)" }
    $screenshot = Resolve-EvidenceFile $lifecycleDirectory $state.screenshot; Assert-FileHash $screenshot $state.screenshotSha256; [void]$lifecycleFiles.Add($screenshot)
    $tree = Resolve-EvidenceFile $lifecycleDirectory $state.uiaTree; Assert-FileHash $tree $state.uiaTreeSha256; [void]$lifecycleFiles.Add($tree)
    $stateMap[$state.name] = $state
}
foreach ($theme in @('Light','Dark','System')) {
    $restart = @($lifecycle.restarts | Where-Object { $_.theme -eq $theme })
    if ($restart.Count -ne 1 -or $restart[0].passed -ne $true -or $restart[0].graceful -ne $false) { throw "Missing fail-closed $theme real-process restart evidence." }
    $before = $stateMap[$restart[0].beforeState]; $after = $stateMap[$restart[0].afterState]
    if ($null -eq $before -or $null -eq $after -or $before.selectedTheme -ne $theme -or $after.selectedTheme -ne $theme -or $before.processId -eq $after.processId -or $before.processStartUtc -eq $after.processStartUtc -or ([DateTime]$before.capturedUtc).ToUniversalTime() -ge ([DateTime]$after.capturedUtc).ToUniversalTime()) { throw "$theme selection was not proved across a new application process." }
}
$transitionStates = @($stateMap[$lifecycle.transition.initialState],$stateMap[$lifecycle.transition.appThemeChangedState],$stateMap[$lifecycle.transition.highContrastChangedState])
if (@($transitionStates | Where-Object { $null -eq $_ }).Count -ne 0) { throw 'Lifecycle transition references a missing captured state.' }
foreach ($state in $transitionStates) {
    if ($state.processId -ne $lifecycle.transition.processId -or $state.processStartUtc -ne $lifecycle.transition.processStartUtc -or $state.selectedTheme -ne 'System') { throw 'Windows transitions were not captured against one unchanged, System-themed product process.' }
}
if ($transitionStates[0].windowsAppTheme -notin @('Light','Dark') -or $transitionStates[1].windowsAppTheme -notin @('Light','Dark') -or $transitionStates[0].windowsAppTheme -eq $transitionStates[1].windowsAppTheme -or $transitionStates[0].highContrast -ne $false -or $transitionStates[1].highContrast -ne $false -or $transitionStates[2].highContrast -ne $true) { throw 'The recorded Windows app-theme/high-contrast transitions are not the required meaningful state changes.' }
$systemRestart = @($lifecycle.restarts | Where-Object { $_.theme -eq 'System' })[0]
if ($stateMap[$systemRestart.afterState].processId -ne $transitionStates[0].processId -or ([DateTime]$transitionStates[0].capturedUtc).ToUniversalTime() -ge ([DateTime]$transitionStates[1].capturedUtc).ToUniversalTime() -or ([DateTime]$transitionStates[1].capturedUtc).ToUniversalTime() -ge ([DateTime]$transitionStates[2].capturedUtc).ToUniversalTime()) { throw 'Lifecycle transition ordering or binding to the final System restart is invalid.' }
$actualLifecycleFiles = @(Get-ChildItem $lifecycleDirectory -File)
if ($actualLifecycleFiles.Count -ne $lifecycleFiles.Count) { throw 'Lifecycle evidence contains missing or unreviewed files.' }
foreach ($file in $actualLifecycleFiles) { if (-not $lifecycleFiles.Contains($file.FullName)) { throw "Unreferenced lifecycle evidence file: $($file.FullName)" } }

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
