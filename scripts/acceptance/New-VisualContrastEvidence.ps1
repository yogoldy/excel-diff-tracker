[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string] $EvidenceRoot,
    [Parameter(Mandatory)] [string] $ThemeManagerSourcePath,
    [Parameter(Mandatory)] [ValidatePattern('^[A-Fa-f0-9]{64}$')] [string] $ExpectedInstallerSha256,
    [Parameter(Mandatory)] [ValidatePattern('^[A-Fa-f0-9]{64}$')] [string] $ExpectedApplicationSha256,
    [Parameter(Mandatory)] [ValidatePattern('^[0-9a-fA-F-]{36}$')] [string] $CaptureSessionId
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'UiAutomation.psm1') -Force

$root = [System.IO.Path]::GetFullPath($EvidenceRoot)
$output = Join-Path $root 'contrast'
if (Test-Path $output) { throw "Fresh contrast evidence is required; the output already exists: $output" }
New-Item -ItemType Directory -Path $output | Out-Null
$source = (Resolve-Path $ThemeManagerSourcePath).Path
$sourceCopy = Join-Path $output 'ThemeManager.cs'
Copy-Item -LiteralPath $source -Destination $sourceCopy

function Get-Palette {
    param([string] $Content, [string] $Name)
    $block = [regex]::Match($Content,"(?s)${Name}Colors\s*=\s*new\s+Dictionary<string,\s*string>\s*\{(?<body>.*?)\};")
    if (-not $block.Success) { throw "Could not parse the deterministic $Name palette." }
    $palette = [ordered]@{}
    foreach ($entry in [regex]::Matches($block.Groups['body'].Value,'\["(?<key>[^"]+)"\]\s*=\s*"(?<color>#[0-9A-Fa-f]{6})"')) {
        $key = $entry.Groups['key'].Value
        if ($palette.Contains($key)) { throw "Duplicate $Name palette key: $key" }
        $palette[$key] = $entry.Groups['color'].Value.ToUpperInvariant()
    }
    $required = @('AppBackgroundBrush','SidebarBrush','CardBrush','CardHoverBrush','TextBrush','MutedTextBrush','BorderBrush','AccentBrush','AccentSoftBrush','WarningBrush','ErrorBrush','PrimaryForegroundBrush')
    foreach ($key in $required) { if (-not $palette.Contains($key)) { throw "The $Name palette is missing $key." } }
    if ($palette.Count -ne $required.Count) { throw "The $Name palette contains an unreviewed color resource; update the contrast contract." }
    $palette
}

function Get-RelativeLuminance {
    param([string] $Color)
    $channels = @([Convert]::ToInt32($Color.Substring(1,2),16),[Convert]::ToInt32($Color.Substring(3,2),16),[Convert]::ToInt32($Color.Substring(5,2),16))
    $linear = foreach ($channel in $channels) {
        $value = $channel / 255.0
        if ($value -le 0.04045) { $value / 12.92 } else { [Math]::Pow((($value + 0.055) / 1.055),2.4) }
    }
    (0.2126 * $linear[0]) + (0.7152 * $linear[1]) + (0.0722 * $linear[2])
}

function Get-ContrastRatio {
    param([string] $Foreground, [string] $Background)
    $first = Get-RelativeLuminance $Foreground; $second = Get-RelativeLuminance $Background
    ([Math]::Max($first,$second) + 0.05) / ([Math]::Min($first,$second) + 0.05)
}

function New-Check {
    param([string] $Id,[string] $Category,[string] $Foreground,[string] $Background,[double] $Threshold)
    [pscustomobject]@{ id = $Id; category = $Category; foregroundResource = $Foreground; backgroundResource = $Background; threshold = $Threshold }
}

$contract = @(
    New-Check 'body-on-app' 'normal-text' 'TextBrush' 'AppBackgroundBrush' 4.5
    New-Check 'body-on-sidebar' 'normal-text' 'TextBrush' 'SidebarBrush' 4.5
    New-Check 'body-on-card' 'normal-text' 'TextBrush' 'CardBrush' 4.5
    New-Check 'body-on-hover' 'normal-text' 'TextBrush' 'CardHoverBrush' 4.5
    New-Check 'body-on-accent-soft' 'normal-text' 'TextBrush' 'AccentSoftBrush' 4.5
    New-Check 'muted-on-app' 'normal-text' 'MutedTextBrush' 'AppBackgroundBrush' 4.5
    New-Check 'muted-on-sidebar' 'normal-text' 'MutedTextBrush' 'SidebarBrush' 4.5
    New-Check 'muted-on-card' 'normal-text' 'MutedTextBrush' 'CardBrush' 4.5
    New-Check 'muted-on-hover' 'normal-text' 'MutedTextBrush' 'CardHoverBrush' 4.5
    New-Check 'primary-label' 'normal-text' 'PrimaryForegroundBrush' 'AccentBrush' 4.5
    New-Check 'warning-label' 'normal-text' 'WarningBrush' 'CardBrush' 4.5
    New-Check 'error-label' 'normal-text' 'ErrorBrush' 'CardBrush' 4.5
    New-Check 'page-title' 'large-text' 'TextBrush' 'AppBackgroundBrush' 3.0
    New-Check 'section-title' 'large-text' 'TextBrush' 'CardBrush' 3.0
    New-Check 'border-on-card' 'essential-ui' 'BorderBrush' 'CardBrush' 3.0
    New-Check 'border-on-app' 'essential-ui' 'BorderBrush' 'AppBackgroundBrush' 3.0
    New-Check 'focus-on-card' 'essential-ui' 'AccentBrush' 'CardBrush' 3.0
    New-Check 'focus-on-sidebar' 'essential-ui' 'AccentBrush' 'SidebarBrush' 3.0
    New-Check 'primary-fill-on-app' 'essential-ui' 'AccentBrush' 'AppBackgroundBrush' 3.0
    New-Check 'primary-fill-on-card' 'essential-ui' 'AccentBrush' 'CardBrush' 3.0
    New-Check 'primary-focus-on-fill' 'essential-ui' 'PrimaryForegroundBrush' 'AccentBrush' 3.0
    New-Check 'warning-symbol' 'essential-ui' 'WarningBrush' 'CardBrush' 3.0
    New-Check 'error-symbol' 'essential-ui' 'ErrorBrush' 'CardBrush' 3.0
)

$content = [System.IO.File]::ReadAllText($sourceCopy)
$palettes = [ordered]@{ Light = Get-Palette $content 'Light'; Dark = Get-Palette $content 'Dark' }
$checks = [System.Collections.Generic.List[object]]::new()
foreach ($theme in @('Light','Dark')) {
    foreach ($item in $contract) {
        $foreground = $palettes[$theme][$item.foregroundResource]; $background = $palettes[$theme][$item.backgroundResource]
        $ratio = Get-ContrastRatio $foreground $background
        $checks.Add([pscustomobject]@{
            theme = $theme; id = $item.id; category = $item.category; foregroundResource = $item.foregroundResource; foreground = $foreground
            backgroundResource = $item.backgroundResource; background = $background; ratio = [Math]::Round($ratio,4); threshold = $item.threshold; passed = $ratio -ge $item.threshold
        })
    }
}

$failed = @($checks | Where-Object { -not $_.passed })
$artifactPath = Join-Path $output 'visual-contrast.json'
$artifact = [ordered]@{
    schemaVersion = 1; status = if ($failed.Count -eq 0) { 'Passed' } else { 'Failed' }; capturedUtc = [DateTime]::UtcNow.ToString('O')
    captureSessionId = [Guid]::Parse($CaptureSessionId).ToString('D'); installerSha256 = $ExpectedInstallerSha256.ToUpperInvariant(); applicationSha256 = $ExpectedApplicationSha256.ToUpperInvariant()
    source = [ordered]@{ relativePath = 'ThemeManager.cs'; sha256 = Get-AcceptanceFileSha256 $sourceCopy }
    scope = 'Deterministic opaque Light/Dark resource colors only; rendered pixels, opacity, compositing, fonts, images, and Windows SystemColors high-contrast colors require separate human/live evidence.'
    palettes = $palettes; checks = @($checks); failedCount = $failed.Count
}
Write-AcceptanceUtf8File -Path $artifactPath -Content ($artifact | ConvertTo-Json -Depth 20)
$artifactHash = Get-AcceptanceFileSha256 $artifactPath
Write-AcceptanceUtf8File -Path (Join-Path $output 'visual-contrast.sha256') -Content "$artifactHash  visual-contrast.json"
if ($failed.Count -ne 0) { throw "VISUAL_CONTRAST_FAILED|$artifactPath|$($failed.Count) checks below threshold" }
Write-Output "VISUAL_CONTRAST_PASSED|$artifactPath"
