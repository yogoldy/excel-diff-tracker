[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string] $EvidenceRoot,
    [Parameter(Mandatory)] [string] $InstallerPath,
    [Parameter(Mandatory)] [ValidatePattern('^[A-Fa-f0-9]{64}$')] [string] $ExpectedInstallerSha256,
    [Parameter(Mandatory)] [ValidatePattern('^[A-Fa-f0-9]{64}$')] [string] $ExpectedApplicationSha256,
    [Parameter(Mandatory)] [ValidatePattern('^[0-9a-fA-F-]{36}$')] [string] $CaptureSessionId,
    [Parameter(Mandatory)] [ValidateSet('Light','Dark')] [string] $ExpectedInitialWindowsAppTheme,
    [Parameter(Mandatory)] [ValidateSet('Light','Dark')] [string] $ExpectedChangedWindowsAppTheme,
    [Parameter(Mandatory)] [bool] $ExpectedInitialHighContrast,
    [Parameter(Mandatory)] [bool] $ExpectedChangedHighContrast,
    [ValidateRange(15,600)] [int] $TransitionTimeoutSeconds = 180
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
Import-Module (Join-Path $PSScriptRoot 'UiAutomation.psm1') -Force

if ($ExpectedInitialWindowsAppTheme -eq $ExpectedChangedWindowsAppTheme) { throw 'The Windows app-theme transition must change Light to Dark or Dark to Light.' }
if ($ExpectedInitialHighContrast -ne $false -or $ExpectedChangedHighContrast -ne $true) { throw 'Capture Windows app-theme redraw with high contrast off, then enable high contrast so both transitions are visually meaningful.' }
$root = [System.IO.Path]::GetFullPath($EvidenceRoot); $output = Join-Path $root 'lifecycle'
if (Test-Path $output) { throw "Fresh lifecycle evidence is required; the output already exists: $output" }
New-Item -ItemType Directory -Path $output | Out-Null
$expectedInstallerHash = $ExpectedInstallerSha256.ToUpperInvariant(); $expectedApplicationHash = $ExpectedApplicationSha256.ToUpperInvariant()
$resolvedInstaller = (Resolve-Path $InstallerPath).Path
if ((Get-AcceptanceFileSha256 $resolvedInstaller) -ne $expectedInstallerHash) { throw 'The lifecycle installer hash does not match the frozen candidate.' }

function Get-WindowsState {
    $value = $null
    try { $value = [int](Get-ItemProperty -LiteralPath 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' -Name AppsUseLightTheme -ErrorAction Stop).AppsUseLightTheme } catch {}
    [pscustomobject]@{ appTheme = if ($value -eq 1) { 'Light' } elseif ($value -eq 0) { 'Dark' } else { 'Unknown' }; highContrast = [bool][System.Windows.Forms.SystemInformation]::HighContrast }
}

function Get-ProductProcess {
    $items = @(Get-Process -Name ExcelDiffTracker -ErrorAction SilentlyContinue | Where-Object { -not $_.HasExited })
    if ($items.Count -ne 1) { throw "Expected exactly one running product process; found $($items.Count)." }
    $item = $items[0]
    $null = & (Join-Path $PSScriptRoot 'Test-SingleFilePayload.ps1') -ExecutablePath $item.Path
    if ((Get-AcceptanceFileSha256 $item.Path) -ne $expectedApplicationHash) { throw 'The running product does not match the frozen application hash.' }
    $item
}

function Get-SelectedTheme {
    param([System.Windows.Automation.AutomationElement] $Window)
    $combo = Find-UiaElement -Root $Window -AutomationId 'ThemeComboBox'
    $pattern = $null
    if (-not $combo.TryGetCurrentPattern([System.Windows.Automation.SelectionPattern]::Pattern,[ref]$pattern)) { throw 'Theme ComboBox does not expose SelectionPattern.' }
    $selection = @(([System.Windows.Automation.SelectionPattern]$pattern).Current.GetSelection())
    if ($selection.Count -ne 1) { throw 'Theme ComboBox does not expose exactly one selection.' }
    $selection[0].Current.Name
}

function Set-SelectedTheme {
    param([System.Windows.Automation.AutomationElement] $Window,[string] $Theme)
    Invoke-UiaElement (Find-UiaElement -Root $Window -AutomationId 'SettingsNavigationButton'); Start-Sleep -Milliseconds 250
    $combo = Find-UiaElement -Root $Window -AutomationId 'ThemeComboBox'; Invoke-UiaElement $combo; Start-Sleep -Milliseconds 150
    Invoke-UiaElement (Find-UiaElement -Root $Window -Name $Theme); Start-Sleep -Milliseconds 900
    if ((Get-SelectedTheme $Window) -ne $Theme) { throw "Could not select application theme $Theme." }
}

function Save-State {
    param([string] $Name,[System.Diagnostics.Process] $Process,[System.Windows.Automation.AutomationElement] $Window)
    $png = Join-Path $output "$Name.png"; $tree = Join-Path $output "$Name.uia.json"
    Save-DesktopScreenshot $png; Export-UiaTree -Root $Window -Path $tree
    $windows = Get-WindowsState
    [pscustomobject]@{
        name = $Name; capturedUtc = [DateTime]::UtcNow.ToString('O'); processId = $Process.Id; processStartUtc = $Process.StartTime.ToUniversalTime().ToString('O')
        applicationPath = $Process.Path; applicationSha256 = Get-AcceptanceFileSha256 $Process.Path; selectedTheme = Get-SelectedTheme $Window
        windowsAppTheme = $windows.appTheme; highContrast = $windows.highContrast
        screenshot = "$Name.png"; screenshotSha256 = Get-AcceptanceFileSha256 $png; uiaTree = "$Name.uia.json"; uiaTreeSha256 = Get-AcceptanceFileSha256 $tree
    }
}

function Restart-Product {
    param([System.Diagnostics.Process] $Process)
    $path = $Process.Path
    Stop-Process -Id $Process.Id
    Wait-AcceptanceCondition -TimeoutSeconds 20 -FailureMessage 'The prior product process did not exit for the restart test.' -Condition { $null -eq (Get-Process -Id $Process.Id -ErrorAction SilentlyContinue) }
    $started = Start-Process -FilePath $path -PassThru
    $window = Find-UiaWindow -Title 'Excel Diff Tracker' -ProcessId $started.Id -TimeoutSeconds 30
    [pscustomobject]@{ process = Get-Process -Id $started.Id; window = $window }
}

$states = [System.Collections.Generic.List[object]]::new(); $restarts = [System.Collections.Generic.List[object]]::new()
$product = Get-ProductProcess; $main = Find-UiaWindow -Title 'Excel Diff Tracker' -ProcessId $product.Id -TimeoutSeconds 20
foreach ($theme in @('Light','Dark','System')) {
    Set-SelectedTheme $main $theme
    $before = Save-State "restart-$($theme.ToLowerInvariant())-before" $product $main; $states.Add($before)
    $restarted = Restart-Product $product; $product = $restarted.process; $main = $restarted.window
    Invoke-UiaElement (Find-UiaElement -Root $main -AutomationId 'SettingsNavigationButton'); Start-Sleep -Milliseconds 350
    $after = Save-State "restart-$($theme.ToLowerInvariant())-after" $product $main; $states.Add($after)
    if ($after.processId -eq $before.processId -or $after.processStartUtc -eq $before.processStartUtc -or $after.selectedTheme -ne $theme) { throw "Application theme $theme did not survive a real process restart." }
    $restarts.Add([pscustomobject]@{ theme = $theme; method = 'Stop-Process then Start-Process using the exact installed executable'; graceful = $false; beforeState = $before.name; afterState = $after.name; passed = $true })
}

$transitionPid = $product.Id; $transitionStart = $product.StartTime.ToUniversalTime().ToString('O')
$initial = Get-WindowsState
if ($initial.appTheme -ne $ExpectedInitialWindowsAppTheme -or $initial.highContrast -ne $ExpectedInitialHighContrast) { throw 'Measured initial Windows theme/high-contrast state does not match the requested lifecycle transition.' }
$initialState = Save-State 'transition-initial' $product $main; $states.Add($initialState)
Write-Host "Change Windows app theme to $ExpectedChangedWindowsAppTheme without closing Excel Diff Tracker. Waiting up to $TransitionTimeoutSeconds seconds..."
Wait-AcceptanceCondition -TimeoutSeconds $TransitionTimeoutSeconds -FailureMessage 'The requested Windows app-theme transition was not observed.' -Condition {
    $live = Get-Process -Id $transitionPid -ErrorAction SilentlyContinue; if ($null -eq $live) { throw 'The identified product process exited during the Windows app-theme transition.' }
    $observed = Get-WindowsState; $observed.appTheme -eq $ExpectedChangedWindowsAppTheme -and $observed.highContrast -eq $ExpectedInitialHighContrast
}
Start-Sleep -Milliseconds 1000
$appThemeState = Save-State 'transition-app-theme-changed' $product $main; $states.Add($appThemeState)
Write-Host "Change Windows high contrast to $ExpectedChangedHighContrast without closing Excel Diff Tracker. Waiting up to $TransitionTimeoutSeconds seconds..."
Wait-AcceptanceCondition -TimeoutSeconds $TransitionTimeoutSeconds -FailureMessage 'The requested Windows high-contrast transition was not observed.' -Condition {
    $live = Get-Process -Id $transitionPid -ErrorAction SilentlyContinue; if ($null -eq $live) { throw 'The identified product process exited during the Windows high-contrast transition.' }
    (Get-WindowsState).highContrast -eq $ExpectedChangedHighContrast
}
Start-Sleep -Milliseconds 1000
$contrastState = Save-State 'transition-high-contrast-changed' $product $main; $states.Add($contrastState)
foreach ($state in @($initialState,$appThemeState,$contrastState)) {
    if ($state.processId -ne $transitionPid -or $state.processStartUtc -ne $transitionStart -or $state.selectedTheme -ne 'System') { throw 'Windows transitions were not observed against one unchanged System-themed product process.' }
}

$artifactPath = Join-Path $output 'visual-lifecycle.json'
$artifact = [ordered]@{
    schemaVersion = 1; status = 'MachineObservationPassedHumanRenderReviewRequired'; capturedUtc = [DateTime]::UtcNow.ToString('O')
    captureSessionId = [Guid]::Parse($CaptureSessionId).ToString('D'); installerPathAtCapture = $resolvedInstaller; installerSha256 = $expectedInstallerHash; applicationSha256 = $expectedApplicationHash
    restartMethodLimitation = 'The harness proves new operating-system process identities using the exact executable, but terminates with Stop-Process rather than the tray Exit command. Graceful tray exit is covered separately.'
    transitionRenderLimitation = 'Windows state, one unchanged product PID, System selection, UIA trees, and screenshots are machine-recorded. UI Automation does not expose final WPF brush pixels; a human must compare the before/after screenshots.'
    restarts = @($restarts); transition = [ordered]@{ processId = $transitionPid; processStartUtc = $transitionStart; initialState = $initialState.name; appThemeChangedState = $appThemeState.name; highContrastChangedState = $contrastState.name; passed = $true }
    states = @($states)
}
Write-AcceptanceUtf8File -Path $artifactPath -Content ($artifact | ConvertTo-Json -Depth 20)
$hash = Get-AcceptanceFileSha256 $artifactPath
Write-AcceptanceUtf8File -Path (Join-Path $output 'visual-lifecycle.sha256') -Content "$hash  visual-lifecycle.json"
Write-Output "VISUAL_LIFECYCLE_CAPTURED_REVIEW_REQUIRED|$artifactPath"
