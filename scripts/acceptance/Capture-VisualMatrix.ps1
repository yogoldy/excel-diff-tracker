[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string] $EvidenceRoot,
    [Parameter(Mandatory)] [ValidateSet('100','125','150','200')] [string] $ScalePercent,
    [Parameter(Mandatory)] [string] $DisplayLabel,
    [switch] $WindowsContrastTheme
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'UiAutomation.psm1') -Force

$root = [System.IO.Path]::GetFullPath($EvidenceRoot)
$configuration = "$DisplayLabel-$ScalePercent-percent$(if ($WindowsContrastTheme) { '-contrast' } else { '' })"
$output = Join-Path $root $configuration
$screenshots = Join-Path $output 'screenshots'
$trees = Join-Path $output 'uia'
New-Item -ItemType Directory -Path $screenshots, $trees -Force | Out-Null
$results = [System.Collections.Generic.List[object]]::new()
$window = Find-UiaWindow -Title 'Excel Diff Tracker' -TimeoutSeconds 20

function Capture-State {
    param([string] $Name)
    $script:window = Find-UiaWindow -Title 'Excel Diff Tracker' -TimeoutSeconds 10
    $png = Join-Path $screenshots "$Name.png"
    $json = Join-Path $trees "$Name.json"
    Save-DesktopScreenshot -Path $png
    Export-UiaTree -Root $script:window -Path $json
    $windowBounds = $script:window.Current.BoundingRectangle
    $visible = $script:window.FindAll(
        [System.Windows.Automation.TreeScope]::Descendants,
        [System.Windows.Automation.Condition]::TrueCondition)
    $geometryFailures = [System.Collections.Generic.List[string]]::new()
    foreach ($element in $visible) {
        if ($element.Current.IsOffscreen) { continue }
        $bounds = $element.Current.BoundingRectangle
        if ($bounds.IsEmpty -or [double]::IsInfinity($bounds.Width) -or [double]::IsInfinity($bounds.Height)) { continue }
        if ($bounds.Width -le 0 -or $bounds.Height -le 0) {
            $geometryFailures.Add("Non-positive bounds: $($element.Current.ControlType.ProgrammaticName) '$($element.Current.Name)'.")
            continue
        }
        $tolerance = 2.0
        if ($bounds.Left -lt $windowBounds.Left - $tolerance -or
            $bounds.Top -lt $windowBounds.Top - $tolerance -or
            $bounds.Right -gt $windowBounds.Right + $tolerance -or
            $bounds.Bottom -gt $windowBounds.Bottom + $tolerance) {
            $geometryFailures.Add("Visible element leaves window bounds: $($element.Current.ControlType.ProgrammaticName) '$($element.Current.Name)' [$bounds].")
        }
    }
    $results.Add([pscustomobject]@{
        state = $Name
        passed = $geometryFailures.Count -eq 0
        geometryFailures = $geometryFailures
        screenshot = Get-AcceptanceRelativePath -BasePath $output -Path $png -UseForwardSlash
        uiaTree = Get-AcceptanceRelativePath -BasePath $output -Path $json -UseForwardSlash
    })
}

function Navigate-To {
    param([string] $Page)
    $button = Find-UiaElement -Root $script:window -AutomationId "${Page}NavigationButton"
    Invoke-UiaElement -Element $button
    Start-Sleep -Milliseconds 350
    Capture-State $Page.ToLowerInvariant()
}

function Select-Theme {
    param([string] $Theme)
    $settingsButton = Find-UiaElement -Root $script:window -AutomationId 'SettingsNavigationButton'
    Invoke-UiaElement -Element $settingsButton
    Start-Sleep -Milliseconds 350
    $script:window = Find-UiaWindow -Title 'Excel Diff Tracker' -TimeoutSeconds 10
    $combo = Find-UiaElement -Root $script:window -AutomationId 'ThemeComboBox'
    Invoke-UiaElement -Element $combo
    Start-Sleep -Milliseconds 150
    $themeItem = Find-UiaElement -Root $script:window -Name $Theme
    Invoke-UiaElement -Element $themeItem
    Start-Sleep -Milliseconds 500
}

$themes = if ($WindowsContrastTheme) { @('System') } else { @('Light','Dark','System') }
$pages = @('Dashboard','Workbooks','History','Settings','About')
foreach ($theme in $themes) {
    Select-Theme $theme
    foreach ($page in $pages) {
        Navigate-To $page
        $last = $results[$results.Count - 1]
        $last.state = "$($theme.ToLowerInvariant())-$($last.state)"
        $sourcePng = Join-Path $screenshots "$($page.ToLowerInvariant()).png"
        $sourceTree = Join-Path $trees "$($page.ToLowerInvariant()).json"
        $targetPng = Join-Path $screenshots "$($last.state).png"
        $targetTree = Join-Path $trees "$($last.state).json"
        Move-Item $sourcePng $targetPng -Force
        Move-Item $sourceTree $targetTree -Force
        $last.screenshot = Get-AcceptanceRelativePath -BasePath $output -Path $targetPng -UseForwardSlash
        $last.uiaTree = Get-AcceptanceRelativePath -BasePath $output -Path $targetTree -UseForwardSlash
    }
}

$summary = [ordered]@{
    status = if (@($results | Where-Object { -not $_.passed }).Count -eq 0) { 'GeometryPassedHumanReviewRequired' } else { 'Failed' }
    configuration = $configuration
    scalePercent = [int]$ScalePercent
    windowsContrastTheme = [bool]$WindowsContrastTheme
    capturedUtc = [DateTime]::UtcNow.ToString('O')
    results = $results
    humanReview = [ordered]@{
        required = $true
        contrast = 'Pending'
        clippingAndOverlap = 'Pending'
        keyboardAndFocus = 'Pending'
        reviewer = $null
    }
}
$summaryPath = Join-Path $output 'visual-matrix.json'
Write-AcceptanceUtf8File -Path $summaryPath -Content ($summary | ConvertTo-Json -Depth 10)
if ($summary.status -eq 'Failed') { throw "VISUAL_MATRIX_GEOMETRY_FAILED|$summaryPath" }
Write-Output "VISUAL_MATRIX_CAPTURED_REVIEW_REQUIRED|$summaryPath"
