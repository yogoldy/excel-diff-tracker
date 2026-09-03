[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string] $EvidenceRoot,
    [Parameter(Mandatory)] [Alias('ScalePercent')] [ValidateSet('100','125','150','200')] [string] $ExpectedScalePercent,
    [Parameter(Mandatory)] [Alias('DisplayLabel')] [string] $OperatorDisplayLabel,
    [Parameter(Mandatory)] [string] $InstallerPath,
    [Parameter(Mandatory)] [ValidatePattern('^[A-Fa-f0-9]{64}$')] [string] $ExpectedInstallerSha256,
    [Parameter(Mandatory)] [ValidatePattern('^[A-Fa-f0-9]{64}$')] [string] $ExpectedApplicationSha256,
    [Parameter(Mandatory)] [ValidatePattern('^[0-9a-fA-F-]{36}$')] [string] $CaptureSessionId,
    [ValidateSet('Onboarding','Main','Supplemental')] [string] $CapturePhase = 'Main',
    [ValidateSet('Minimum','Default','Resized')] [string] $WindowSizeMode = 'Default',
    [string] $WorkbookPath,
    [string] $ExpectedLongPath,
    [string] $ExpectedWarningText,
    [string] $ExpectedStateText,
    [ValidateSet('tray-menu','toast','processing','warning','error','long-path')] [string] $StateName,
    [string] $TargetWindowTitle = 'Excel Scenario Analysis Tool',
    [switch] $UseDesktopRoot,
    [switch] $WindowsContrastTheme
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
Import-Module (Join-Path $PSScriptRoot 'UiAutomation.psm1') -Force

if (-not ('ExcelDiffTrackerVisualNative' -as [type])) {
Add-Type @"
using System;
using System.Runtime.InteropServices;

public static class ExcelDiffTrackerVisualNative
{
    [DllImport("user32.dll")]
    public static extern IntPtr SetThreadDpiAwarenessContext(IntPtr context);

    [DllImport("user32.dll")]
    public static extern uint GetDpiForWindow(IntPtr hwnd);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool SetWindowPos(IntPtr hwnd, IntPtr insertAfter, int x, int y, int cx, int cy, uint flags);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }

    [StructLayout(LayoutKind.Sequential)]
    public struct MONITORINFO { public int cbSize; public RECT rcMonitor; public RECT rcWork; public uint dwFlags; }

    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool GetWindowRect(IntPtr hwnd, out RECT rect);

    [DllImport("user32.dll")]
    public static extern IntPtr MonitorFromWindow(IntPtr hwnd, uint flags);

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    public static extern bool GetMonitorInfo(IntPtr monitor, ref MONITORINFO info);

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct DEVMODE
    {
        private const int CCHDEVICENAME = 32;
        private const int CCHFORMNAME = 32;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = CCHDEVICENAME)] public string dmDeviceName;
        public short dmSpecVersion;
        public short dmDriverVersion;
        public short dmSize;
        public short dmDriverExtra;
        public int dmFields;
        public int dmPositionX;
        public int dmPositionY;
        public int dmDisplayOrientation;
        public int dmDisplayFixedOutput;
        public short dmColor;
        public short dmDuplex;
        public short dmYResolution;
        public short dmTTOption;
        public short dmCollate;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = CCHFORMNAME)] public string dmFormName;
        public short dmLogPixels;
        public int dmBitsPerPel;
        public int dmPelsWidth;
        public int dmPelsHeight;
        public int dmDisplayFlags;
        public int dmDisplayFrequency;
        public int dmICMMethod;
        public int dmICMIntent;
        public int dmMediaType;
        public int dmDitherType;
        public int dmReserved1;
        public int dmReserved2;
        public int dmPanningWidth;
        public int dmPanningHeight;
    }

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern bool EnumDisplaySettings(string deviceName, int modeNum, ref DEVMODE devMode);
}
"@
}

$previousDpiContext = [ExcelDiffTrackerVisualNative]::SetThreadDpiAwarenessContext([IntPtr](-4))
if ($previousDpiContext -eq [IntPtr]::Zero) { throw 'Could not make the acceptance capture thread per-monitor-DPI aware.' }

$root = [System.IO.Path]::GetFullPath($EvidenceRoot)
$expectedInstallerHash = $ExpectedInstallerSha256.ToUpperInvariant()
$expectedApplicationHash = $ExpectedApplicationSha256.ToUpperInvariant()
$sessionGuid = [Guid]::Parse($CaptureSessionId).ToString('D')
$resolvedInstaller = (Resolve-Path $InstallerPath).Path
if ((Get-AcceptanceFileSha256 -Path $resolvedInstaller) -ne $expectedInstallerHash) {
    throw "Installer hash mismatch for the visual candidate: $resolvedInstaller"
}

function Get-RegistryDword {
    param([string] $Path, [string] $Name)
    try {
        $value = (Get-ItemProperty -LiteralPath $Path -Name $Name -ErrorAction Stop).$Name
        if ($null -eq $value) { return $null }
        return [int]$value
    }
    catch { return $null }
}

function Get-PhysicalDisplayModes {
    $modes = [System.Collections.Generic.List[object]]::new()
    foreach ($screen in [System.Windows.Forms.Screen]::AllScreens) {
        $mode = New-Object ExcelDiffTrackerVisualNative+DEVMODE
        $mode.dmSize = [Runtime.InteropServices.Marshal]::SizeOf($mode)
        $read = [ExcelDiffTrackerVisualNative]::EnumDisplaySettings($screen.DeviceName, -1, [ref]$mode)
        $modes.Add([pscustomobject]@{
            deviceName = $screen.DeviceName
            primary = [bool]$screen.Primary
            physicalWidth = if ($read) { [int]$mode.dmPelsWidth } else { [int]$screen.Bounds.Width }
            physicalHeight = if ($read) { [int]$mode.dmPelsHeight } else { [int]$screen.Bounds.Height }
            refreshHz = if ($read) { [int]$mode.dmDisplayFrequency } else { $null }
            logicalBounds = [ordered]@{ x = $screen.Bounds.X; y = $screen.Bounds.Y; width = $screen.Bounds.Width; height = $screen.Bounds.Height }
            workingArea = [ordered]@{ x = $screen.WorkingArea.X; y = $screen.WorkingArea.Y; width = $screen.WorkingArea.Width; height = $screen.WorkingArea.Height }
        })
    }
    @($modes)
}

function Get-ProductProcess {
    $processes = @(Get-Process -Name 'ExcelScenarioAnalysisTool' -ErrorAction SilentlyContinue | Where-Object { -not $_.HasExited })
    if ($processes.Count -ne 1) { throw "Expected exactly one running ExcelScenarioAnalysisTool process, found $($processes.Count)." }
    $process = $processes[0]
    if ([string]::IsNullOrWhiteSpace($process.Path) -or -not (Test-Path $process.Path -PathType Leaf)) {
        throw 'The installed ExcelScenarioAnalysisTool process path could not be resolved.'
    }
    $actualHash = Get-AcceptanceFileSha256 -Path $process.Path
    if ($actualHash -ne $expectedApplicationHash) { throw "Installed application hash mismatch. Expected $expectedApplicationHash, actual $actualHash." }
    $process
}

function Get-WindowDpi {
    param([System.Windows.Automation.AutomationElement] $Window)
    $handle = [IntPtr]$Window.Current.NativeWindowHandle
    if ($handle -eq [IntPtr]::Zero) { throw "Window has no native handle: $($Window.Current.Name)" }
    $dpi = [int][ExcelDiffTrackerVisualNative]::GetDpiForWindow($handle)
    if ($dpi -le 0) { throw "Windows returned an invalid DPI for $($Window.Current.Name): $dpi" }
    $dpi
}

function Get-NativeWindowGeometry {
    param([System.Windows.Automation.AutomationElement] $Window)
    $handle = [IntPtr]$Window.Current.NativeWindowHandle
    if ($handle -eq [IntPtr]::Zero) { return $null }
    $rect = New-Object ExcelDiffTrackerVisualNative+RECT
    if (-not [ExcelDiffTrackerVisualNative]::GetWindowRect($handle,[ref]$rect)) { throw "Could not read window bounds for '$($Window.Current.Name)'." }
    $monitor = [ExcelDiffTrackerVisualNative]::MonitorFromWindow($handle,2)
    if ($monitor -eq [IntPtr]::Zero) { throw "Could not resolve the monitor for '$($Window.Current.Name)'." }
    $info = New-Object ExcelDiffTrackerVisualNative+MONITORINFO
    $info.cbSize = [Runtime.InteropServices.Marshal]::SizeOf($info)
    if (-not [ExcelDiffTrackerVisualNative]::GetMonitorInfo($monitor,[ref]$info)) { throw "Could not read monitor bounds for '$($Window.Current.Name)'." }
    $inside = $rect.Left -ge $info.rcWork.Left -and $rect.Top -ge $info.rcWork.Top -and $rect.Right -le $info.rcWork.Right -and $rect.Bottom -le $info.rcWork.Bottom
    [pscustomobject]@{
        actual = [ordered]@{ x = $rect.Left; y = $rect.Top; width = $rect.Right - $rect.Left; height = $rect.Bottom - $rect.Top; right = $rect.Right; bottom = $rect.Bottom }
        monitorWorkingArea = [ordered]@{ x = $info.rcWork.Left; y = $info.rcWork.Top; width = $info.rcWork.Right - $info.rcWork.Left; height = $info.rcWork.Bottom - $info.rcWork.Top; right = $info.rcWork.Right; bottom = $info.rcWork.Bottom }
        fullyInsideMonitorWorkingArea = $inside
    }
}

function Get-ActualEnvironment {
    param([System.Windows.Automation.AutomationElement] $ReferenceWindow)
    $dpi = Get-WindowDpi -Window $ReferenceWindow
    $scale = [int][Math]::Round(($dpi / 96.0) * 100.0)
    $appsUseLight = Get-RegistryDword 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' 'AppsUseLightTheme'
    $systemUseLight = Get-RegistryDword 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' 'SystemUsesLightTheme'
    [pscustomobject]@{
        capturedUtc = [DateTime]::UtcNow.ToString('O')
        windowDpi = $dpi
        scalePercent = $scale
        highContrast = [bool][System.Windows.Forms.SystemInformation]::HighContrast
        windowsAppTheme = if ($appsUseLight -eq 1) { 'Light' } elseif ($appsUseLight -eq 0) { 'Dark' } else { 'Unknown' }
        windowsSystemTheme = if ($systemUseLight -eq 1) { 'Light' } elseif ($systemUseLight -eq 0) { 'Dark' } else { 'Unknown' }
        virtualScreen = [ordered]@{ x = [System.Windows.Forms.SystemInformation]::VirtualScreen.X; y = [System.Windows.Forms.SystemInformation]::VirtualScreen.Y; width = [System.Windows.Forms.SystemInformation]::VirtualScreen.Width; height = [System.Windows.Forms.SystemInformation]::VirtualScreen.Height }
        displays = @(Get-PhysicalDisplayModes)
    }
}

function Assert-ExpectedEnvironment {
    param([object] $Environment)
    if ($Environment.scalePercent -ne [int]$ExpectedScalePercent) { throw "Measured window scaling is $($Environment.scalePercent)% ($($Environment.windowDpi) DPI), not the requested $ExpectedScalePercent%." }
    if ($WindowsContrastTheme -and -not $Environment.highContrast) { throw 'A Windows contrast-theme capture was requested, but Windows reports high contrast is disabled.' }
    if (-not $WindowsContrastTheme -and $Environment.highContrast) { throw 'A normal-theme capture was requested, but Windows reports high contrast is enabled.' }
    if ($Environment.windowsAppTheme -eq 'Unknown' -or $Environment.windowsSystemTheme -eq 'Unknown') { throw 'Windows light/dark theme registry state could not be measured.' }
    $primary = @($Environment.displays | Where-Object { $_.primary })
    if ($primary.Count -ne 1 -or $primary[0].physicalWidth -le 0 -or $primary[0].physicalHeight -le 0) { throw 'The primary physical display resolution could not be measured.' }
}

function Convert-SafeName { param([string] $Value) (($Value.ToLowerInvariant() -replace '[^a-z0-9.-]+','-').Trim('-')) }

function Get-ElementKey {
    param([System.Windows.Automation.AutomationElement] $Element)
    try { return (($Element.GetRuntimeId() | ForEach-Object { $_.ToString() }) -join '.') }
    catch { return "$($Element.Current.ProcessId)|$($Element.Current.AutomationId)|$($Element.Current.ControlType.ProgrammaticName)|$($Element.Current.Name)" }
}

function Test-ElementInsideRoot {
    param([System.Windows.Automation.AutomationElement] $Element, [System.Windows.Automation.AutomationElement] $RootElement)
    $current = $Element
    for ($i = 0; $i -lt 40 -and $null -ne $current; $i++) {
        if ([System.Windows.Automation.AutomationElement]::Equals($current, $RootElement)) { return $true }
        try { $current = [System.Windows.Automation.TreeWalker]::RawViewWalker.GetParent($current) } catch { return $false }
    }
    $false
}

function Get-InteractiveElements {
    param([System.Windows.Automation.AutomationElement] $RootElement)
    $interactiveTypes = @('ControlType.Button','ControlType.CheckBox','ControlType.ComboBox','ControlType.Edit','ControlType.Hyperlink','ControlType.ListItem','ControlType.MenuItem','ControlType.RadioButton','ControlType.Slider','ControlType.Spinner','ControlType.TabItem','ControlType.TreeItem')
    @($RootElement.FindAll([System.Windows.Automation.TreeScope]::Descendants,[System.Windows.Automation.Condition]::TrueCondition) | Where-Object { $interactiveTypes -contains $_.Current.ControlType.ProgrammaticName -and $_.Current.IsEnabled })
}

function Test-HasScrollableAncestor {
    param([System.Windows.Automation.AutomationElement] $Element,[System.Windows.Automation.AutomationElement] $RootElement)
    $current = $Element
    for ($index = 0; $index -lt 40 -and $null -ne $current; $index++) {
        if ([System.Windows.Automation.AutomationElement]::Equals($current,$RootElement)) { return $false }
        try { $current = [System.Windows.Automation.TreeWalker]::ControlViewWalker.GetParent($current) } catch { return $false }
        if ($null -eq $current) { return $false }
        $pattern = $null
        if ($current.TryGetCurrentPattern([System.Windows.Automation.ScrollPattern]::Pattern,[ref]$pattern)) {
            $scroll = ([System.Windows.Automation.ScrollPattern]$pattern).Current
            if ($scroll.HorizontallyScrollable -or $scroll.VerticallyScrollable) { return $true }
        }
    }
    $false
}

function Get-RectangleObject {
    param([System.Windows.Rect] $Rectangle)
    [ordered]@{ x = $Rectangle.X; y = $Rectangle.Y; width = $Rectangle.Width; height = $Rectangle.Height; left = $Rectangle.Left; top = $Rectangle.Top; right = $Rectangle.Right; bottom = $Rectangle.Bottom }
}

function Get-InteractiveState {
    param([System.Windows.Automation.AutomationElement] $Element)
    $state = [ordered]@{ enabled = [bool]$Element.Current.IsEnabled; keyboardFocusable = [bool]$Element.Current.IsKeyboardFocusable; offscreen = [bool]$Element.Current.IsOffscreen; toggle = $null; selected = $null; expandCollapse = $null; value = $null }
    $pattern = $null
    if ($Element.TryGetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern,[ref]$pattern)) { $state.toggle = ([System.Windows.Automation.TogglePattern]$pattern).Current.ToggleState.ToString() }
    $pattern = $null
    if ($Element.TryGetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern,[ref]$pattern)) { $state.selected = [bool]([System.Windows.Automation.SelectionItemPattern]$pattern).Current.IsSelected }
    $pattern = $null
    if ($Element.TryGetCurrentPattern([System.Windows.Automation.ExpandCollapsePattern]::Pattern,[ref]$pattern)) { $state.expandCollapse = ([System.Windows.Automation.ExpandCollapsePattern]$pattern).Current.ExpandCollapseState.ToString() }
    $pattern = $null
    if ($Element.TryGetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern,[ref]$pattern)) { $state.value = ([System.Windows.Automation.ValuePattern]$pattern).Current.Value }
    [pscustomobject]$state
}

function Get-AccessibilityAndGeometry {
    param([System.Windows.Automation.AutomationElement] $RootElement, [int] $ProductProcessId)
    $rootBounds = $RootElement.Current.BoundingRectangle
    $all = @($RootElement.FindAll([System.Windows.Automation.TreeScope]::Descendants,[System.Windows.Automation.Condition]::TrueCondition))
    $geometryFailures = [System.Collections.Generic.List[string]]::new()
    $overlapFailures = [System.Collections.Generic.List[string]]::new()
    $accessibilityFailures = [System.Collections.Generic.List[string]]::new()
    $offscreen = [System.Collections.Generic.List[object]]::new()
    $interactive = @(Get-InteractiveElements -RootElement $RootElement)
    $tolerance = 2.0
    foreach ($element in $all) {
        $bounds = $element.Current.BoundingRectangle
        if ($element.Current.IsOffscreen) {
            if ($element.Current.IsKeyboardFocusable -or $interactive -contains $element) {
                $scrollable = Test-HasScrollableAncestor -Element $element -RootElement $RootElement
                $offscreen.Add([pscustomobject]@{ automationId = $element.Current.AutomationId; name = $element.Current.Name; role = $element.Current.ControlType.ProgrammaticName; insideScrollableContainer = $scrollable; bounds = Get-RectangleObject $bounds })
                if ($element.Current.ProcessId -eq $ProductProcessId -and -not $scrollable) { $geometryFailures.Add("Focusable/interactive product control is off-screen without a scrollable ancestor: $($element.Current.ControlType.ProgrammaticName) '$($element.Current.Name)'.") }
            }
            continue
        }
        if ($bounds.IsEmpty -or [double]::IsInfinity($bounds.Width) -or [double]::IsInfinity($bounds.Height)) { continue }
        if ($bounds.Width -le 0 -or $bounds.Height -le 0) { $geometryFailures.Add("Non-positive bounds: $($element.Current.ControlType.ProgrammaticName) '$($element.Current.Name)'."); continue }
        if (-not $UseDesktopRoot -and ($bounds.Left -lt $rootBounds.Left - $tolerance -or $bounds.Top -lt $rootBounds.Top - $tolerance -or $bounds.Right -gt $rootBounds.Right + $tolerance -or $bounds.Bottom -gt $rootBounds.Bottom + $tolerance)) {
            $geometryFailures.Add("Visible element leaves root bounds: $($element.Current.ControlType.ProgrammaticName) '$($element.Current.Name)' [$bounds].")
        }
    }
    $productInteractive = @($interactive | Where-Object { $_.Current.ProcessId -eq $ProductProcessId -and -not $_.Current.IsOffscreen })
    $ids = @{}
    foreach ($element in $productInteractive) {
        $id = $element.Current.AutomationId; $name = $element.Current.Name; $role = $element.Current.ControlType.ProgrammaticName
        if ([string]::IsNullOrWhiteSpace($id)) { $accessibilityFailures.Add("Interactive product element has no AutomationId: $role '$name'.") }
        elseif ($ids.ContainsKey($id)) { $accessibilityFailures.Add("Duplicate visible interactive AutomationId '$id': '$($ids[$id])' and '$name'.") }
        else { $ids[$id] = $name }
        if ([string]::IsNullOrWhiteSpace($name)) { $accessibilityFailures.Add("Interactive product element has no accessible name: $role AutomationId='$id'.") }
        if ([string]::IsNullOrWhiteSpace($role) -or $role -eq 'ControlType.Custom') { $accessibilityFailures.Add("Interactive product element has no meaningful role: AutomationId='$id' Name='$name'.") }
    }
    for ($leftIndex = 0; $leftIndex -lt $productInteractive.Count; $leftIndex++) {
        $left = $productInteractive[$leftIndex]; $leftBounds = $left.Current.BoundingRectangle
        if ($leftBounds.IsEmpty) { continue }
        for ($rightIndex = $leftIndex + 1; $rightIndex -lt $productInteractive.Count; $rightIndex++) {
            $right = $productInteractive[$rightIndex]; $rightBounds = $right.Current.BoundingRectangle
            if ($rightBounds.IsEmpty) { continue }
            $intersection = [System.Windows.Rect]::Intersect($leftBounds, $rightBounds)
            if ($intersection.IsEmpty -or $intersection.Width -le $tolerance -or $intersection.Height -le $tolerance) { continue }
            $leftContainsRight = $leftBounds.Left -le $rightBounds.Left + $tolerance -and $leftBounds.Top -le $rightBounds.Top + $tolerance -and $leftBounds.Right -ge $rightBounds.Right - $tolerance -and $leftBounds.Bottom -ge $rightBounds.Bottom - $tolerance
            $rightContainsLeft = $rightBounds.Left -le $leftBounds.Left + $tolerance -and $rightBounds.Top -le $leftBounds.Top + $tolerance -and $rightBounds.Right -ge $leftBounds.Right - $tolerance -and $rightBounds.Bottom -ge $leftBounds.Bottom - $tolerance
            if (-not $leftContainsRight -and -not $rightContainsLeft) { $overlapFailures.Add("Interactive controls overlap: '$($left.Current.Name)' [$leftBounds] and '$($right.Current.Name)' [$rightBounds].") }
        }
    }
    [pscustomobject]@{
        geometryFailures = @($geometryFailures); overlapFailures = @($overlapFailures); accessibilityFailures = @($accessibilityFailures); offscreenInteractive = @($offscreen)
        interactive = @($productInteractive | ForEach-Object { [pscustomobject]@{ key = Get-ElementKey $_; automationId = $_.Current.AutomationId; name = $_.Current.Name; role = $_.Current.ControlType.ProgrammaticName; state = Get-InteractiveState $_; bounds = Get-RectangleObject $_.Current.BoundingRectangle } })
    }
}

function Capture-FocusEvidence {
    param([string] $CaptureName, [System.Windows.Automation.AutomationElement] $RootElement, [int] $ProductProcessId, [string] $FocusDirectory)
    $focusFailures = [System.Collections.Generic.List[string]]::new()
    $focusStops = [System.Collections.Generic.List[object]]::new()
    $required = @(Get-InteractiveElements -RootElement $RootElement | Where-Object { $_.Current.ProcessId -eq $ProductProcessId -and $_.Current.IsKeyboardFocusable -and -not $_.Current.IsOffscreen })
    $requiredKeys = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($element in $required) { [void]$requiredKeys.Add((Get-ElementKey $element)) }
    $seen = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $firstKey = $null
    try { Set-UiaForeground -Window $RootElement } catch { $focusFailures.Add($_.Exception.Message) }
    for ($index = 0; $index -lt 80; $index++) {
        [System.Windows.Forms.SendKeys]::SendWait('{TAB}'); Start-Sleep -Milliseconds 100
        $focused = [System.Windows.Automation.AutomationElement]::FocusedElement
        if ($null -eq $focused -or $focused.Current.ProcessId -ne $ProductProcessId -or -not (Test-ElementInsideRoot -Element $focused -RootElement $RootElement)) { continue }
        $key = Get-ElementKey $focused
        if ($null -eq $firstKey) { $firstKey = $key } elseif ($key -eq $firstKey -and $seen.Count -gt 1) { break }
        if (-not $seen.Add($key)) { continue }
        $focusPng = Join-Path $FocusDirectory ("{0}-{1:D2}.png" -f $CaptureName, $seen.Count)
        Save-DesktopScreenshot -Path $focusPng
        $focusStops.Add([pscustomobject]@{ order = $seen.Count; key = $key; processId = $focused.Current.ProcessId; automationId = $focused.Current.AutomationId; name = $focused.Current.Name; role = $focused.Current.ControlType.ProgrammaticName; hasKeyboardFocus = [bool]$focused.Current.HasKeyboardFocus; state = Get-InteractiveState $focused; bounds = Get-RectangleObject $focused.Current.BoundingRectangle; screenshot = $focusPng; screenshotSha256 = Get-AcceptanceFileSha256 -Path $focusPng })
    }
    foreach ($requiredKey in $requiredKeys) { if (-not $seen.Contains($requiredKey)) { $focusFailures.Add("Visible keyboard-focusable control was not reached by Tab: $requiredKey") } }
    foreach ($stop in $focusStops) { if (-not $stop.hasKeyboardFocus -or $stop.state.offscreen) { $focusFailures.Add("Invalid keyboard focus stop: $($stop.automationId) '$($stop.name)'.") } }
    [pscustomobject]@{ failures = @($focusFailures); stops = @($focusStops) }
}

function Assert-RequiredControls {
    param([System.Windows.Automation.AutomationElement] $RootElement, [hashtable] $ExpectedControls)
    $failures = [System.Collections.Generic.List[string]]::new()
    foreach ($id in $ExpectedControls.Keys) {
        $expected = $ExpectedControls[$id]; $element = Find-UiaElement -Root $RootElement -AutomationId $id -Optional
        if ($null -eq $element) { $failures.Add("Required control is missing: AutomationId='$id'."); continue }
        if ($element.Current.Name -ne $expected.name) { $failures.Add("Control '$id' name changed. Expected '$($expected.name)', actual '$($element.Current.Name)'.") }
        if ($element.Current.ControlType.ProgrammaticName -ne $expected.role) { $failures.Add("Control '$id' role changed. Expected '$($expected.role)', actual '$($element.Current.ControlType.ProgrammaticName)'.") }
    }
    @($failures)
}

function Set-CaptureWindowSize {
    param([System.Windows.Automation.AutomationElement] $Window, [string] $Mode, [int] $Dpi)
    $handle = [IntPtr]$Window.Current.NativeWindowHandle; $screen = [System.Windows.Forms.Screen]::FromHandle($handle); $factor = $Dpi / 96.0
    $isOnboarding = $Window.Current.Name -eq 'Welcome to Excel Scenario Analysis Tool'
    $targetWidth = $null; $targetHeight = $null
    if ($Mode -eq 'Minimum') {
        $logicalWidth = if ($isOnboarding) { 660 } else { 820 }; $logicalHeight = if ($isOnboarding) { 490 } else { 560 }
        $targetWidth = [int][Math]::Ceiling($logicalWidth * $factor); $targetHeight = [int][Math]::Ceiling($logicalHeight * $factor)
    } elseif ($Mode -eq 'Resized') { $targetWidth = [int][Math]::Floor($screen.WorkingArea.Width * 0.90); $targetHeight = [int][Math]::Floor($screen.WorkingArea.Height * 0.90) }
    if ($null -ne $targetWidth) {
        $x = $screen.WorkingArea.Left + [int](($screen.WorkingArea.Width - $targetWidth) / 2); $y = $screen.WorkingArea.Top + [int](($screen.WorkingArea.Height - $targetHeight) / 2)
        if (-not [ExcelDiffTrackerVisualNative]::SetWindowPos($handle,[IntPtr]::Zero,$x,$y,$targetWidth,$targetHeight,0x0044)) { throw "Could not set the $Mode test window geometry. Win32 error: $([Runtime.InteropServices.Marshal]::GetLastWin32Error())" }
        Start-Sleep -Milliseconds 400
    }
    $measured = Get-NativeWindowGeometry -Window $Window; $tolerance = 12
    $sizeMatches = $Mode -eq 'Default' -or ([Math]::Abs($measured.actual.width - $targetWidth) -le $tolerance -and [Math]::Abs($measured.actual.height - $targetHeight) -le $tolerance)
    if (-not $measured.fullyInsideMonitorWorkingArea) { throw "The actual $Mode window is outside its monitor working area: $($measured.actual | ConvertTo-Json -Compress)" }
    if (-not $sizeMatches) { throw "Windows/WPF did not apply the requested $Mode size. Target ${targetWidth}x${targetHeight}, actual $($measured.actual.width)x$($measured.actual.height)." }
    [pscustomobject]@{ mode = $Mode; requestedWidth = $targetWidth; requestedHeight = $targetHeight; verified = $measured.fullyInsideMonitorWorkingArea -and $sizeMatches; actual = $measured.actual; monitorWorkingArea = $measured.monitorWorkingArea; fullyInsideMonitorWorkingArea = $measured.fullyInsideMonitorWorkingArea }
}

function Test-AccessibleTextPresent {
    param([System.Windows.Automation.AutomationElement] $RootElement, [string] $Text, [int] $RequiredProcessId = 0)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $false }
    foreach ($element in $RootElement.FindAll([System.Windows.Automation.TreeScope]::Descendants,[System.Windows.Automation.Condition]::TrueCondition)) {
        if (-not $element.Current.IsOffscreen -and ($RequiredProcessId -eq 0 -or $element.Current.ProcessId -eq $RequiredProcessId) -and ($element.Current.Name -eq $Text -or $element.Current.Name.IndexOf($Text,[StringComparison]::OrdinalIgnoreCase) -ge 0)) { return $true }
    }
    $false
}

function Test-AccessibleTextAndHelpPresent {
    param([System.Windows.Automation.AutomationElement] $RootElement, [string] $Text, [int] $RequiredProcessId = 0)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $false }
    foreach ($element in $RootElement.FindAll([System.Windows.Automation.TreeScope]::Descendants,[System.Windows.Automation.Condition]::TrueCondition)) {
        if (-not $element.Current.IsOffscreen -and ($RequiredProcessId -eq 0 -or $element.Current.ProcessId -eq $RequiredProcessId) -and
            ($element.Current.Name -eq $Text -or $element.Current.Name.IndexOf($Text,[StringComparison]::OrdinalIgnoreCase) -ge 0) -and
            -not [string]::IsNullOrWhiteSpace($element.Current.HelpText) -and
            $element.Current.HelpText.IndexOf($Text,[StringComparison]::OrdinalIgnoreCase) -ge 0) { return $true }
    }
    $false
}

function Get-ExactProductTextControls {
    param([System.Windows.Automation.AutomationElement] $RootElement,[string] $Text,[int] $ProductProcessId)
    @($RootElement.FindAll([System.Windows.Automation.TreeScope]::Descendants,[System.Windows.Automation.Condition]::TrueCondition) | Where-Object { $_.Current.ProcessId -eq $ProductProcessId -and -not $_.Current.IsOffscreen -and $_.Current.ControlType.ProgrammaticName -eq 'ControlType.Text' -and $_.Current.Name -eq $Text })
}

function Assert-SelectedTheme {
    param([System.Windows.Automation.AutomationElement] $ComboBox, [string] $Theme)
    $pattern = $null
    if (-not $ComboBox.TryGetCurrentPattern([System.Windows.Automation.SelectionPattern]::Pattern,[ref]$pattern)) { throw 'Theme ComboBox does not expose SelectionPattern.' }
    $selection = @(([System.Windows.Automation.SelectionPattern]$pattern).Current.GetSelection())
    if ($selection.Count -ne 1 -or $selection[0].Current.Name -ne $Theme) { throw "Theme did not apply immediately. Expected '$Theme'." }
}

$product = Get-ProductProcess
$null = & (Join-Path $PSScriptRoot 'Test-SingleFilePayload.ps1') -ExecutablePath $product.Path
$initialTitle = if ($CapturePhase -eq 'Onboarding') { 'Welcome to Excel Scenario Analysis Tool' } else { 'Excel Scenario Analysis Tool' }
$referenceWindow = Find-UiaWindow -Title $initialTitle -ProcessId $product.Id -TimeoutSeconds 20
$environment = Get-ActualEnvironment -ReferenceWindow $referenceWindow
Assert-ExpectedEnvironment -Environment $environment
$windowSizeMeasurement = Set-CaptureWindowSize -Window $referenceWindow -Mode $WindowSizeMode -Dpi $environment.windowDpi
$environment = Get-ActualEnvironment -ReferenceWindow $referenceWindow
Assert-ExpectedEnvironment -Environment $environment

$primaryDisplay = @($environment.displays | Where-Object { $_.primary })[0]
$contrastLabel = if ($environment.highContrast) { 'contrast' } else { 'normal' }
$configuration = "{0}x{1}-{2}pct-{3}-{4}-{5}" -f $primaryDisplay.physicalWidth,$primaryDisplay.physicalHeight,$environment.scalePercent,$environment.windowsAppTheme.ToLowerInvariant(),$contrastLabel,$WindowSizeMode.ToLowerInvariant()
$output = Join-Path $root $configuration; $screenshots = Join-Path $output 'screenshots'; $trees = Join-Path $output 'uia'; $focusDirectory = Join-Path $output 'focus'
New-Item -ItemType Directory -Path $screenshots,$trees,$focusDirectory -Force | Out-Null
$summaryPath = Join-Path $output 'visual-matrix.json'; $results = [System.Collections.Generic.List[object]]::new(); $firstCapturedUtc = [DateTime]::UtcNow.ToString('O')
if (Test-Path $summaryPath -PathType Leaf) {
    $existing = Get-Content $summaryPath -Raw | ConvertFrom-Json
    if ($existing.schemaVersion -ne 2 -or $existing.captureSessionId -ne $sessionGuid) { throw "Refusing to append to stale or incompatible evidence: $summaryPath" }
    if ($existing.installerSha256 -ne $expectedInstallerHash -or $existing.applicationSha256 -ne $expectedApplicationHash) { throw "Refusing to mix candidate identities in $summaryPath" }
    if ($existing.status -eq 'Failed') { throw "Refusing to append to a retained failed visual matrix: $summaryPath" }
    if ($existing.actualEnvironment.scalePercent -ne $environment.scalePercent -or $existing.actualEnvironment.highContrast -ne $environment.highContrast -or $existing.actualEnvironment.windowsAppTheme -ne $environment.windowsAppTheme -or $existing.actualEnvironment.displays[0].physicalWidth -ne $environment.displays[0].physicalWidth -or $existing.actualEnvironment.displays[0].physicalHeight -ne $environment.displays[0].physicalHeight) { throw "Measured display/theme state changed; start a new configuration capture instead of mixing evidence: $summaryPath" }
    foreach ($item in @($existing.results)) { $results.Add($item) }; $firstCapturedUtc = $existing.firstCapturedUtc
    if ($existing.status -eq 'Approved') {
        $existing.status = 'CaptureInProgressHumanReviewInvalidated'
        $existing.humanReview.contrast = 'Pending'
        $existing.humanReview.clippingAndOverlap = 'Pending'
        $existing.humanReview.keyboardAndFocus = 'Pending'
        if ($null -eq $existing.humanReview.PSObject.Properties['themeTransitions']) { $existing.humanReview | Add-Member -NotePropertyName themeTransitions -NotePropertyValue 'Pending' }
        else { $existing.humanReview.themeTransitions = 'Pending' }
        $existing.humanReview.reviewer = $null
        $existing.humanReview.reviewedUtc = $null
        $existing.humanReview.notes = $null
        Write-AcceptanceUtf8File -Path $summaryPath -Content ($existing | ConvertTo-Json -Depth 40)
    }
}

function Capture-State {
    param([string] $Name,[string] $Category,[System.Windows.Automation.AutomationElement] $RootElement,[string] $ApplicationTheme,[hashtable] $ExpectedControls = @{},[switch] $RequireLongPath,[switch] $RequireWarning,[string] $RequiredText,[string] $ObservedStateText,[switch] $StateSemanticVerified)
    if (@($results | Where-Object { $_.state -eq $Name }).Count -ne 0) { throw "State evidence already exists and cannot be overwritten: $Name" }
    $captureName = Convert-SafeName $Name; $png = Join-Path $screenshots "$captureName.png"; $json = Join-Path $trees "$captureName.json"; $focusJson = Join-Path $focusDirectory "$captureName.json"
    foreach ($path in @($png,$json,$focusJson)) { if (Test-Path $path) { throw "Fresh evidence path already exists: $path" } }
    Save-DesktopScreenshot -Path $png; Export-UiaTree -Root $RootElement -Path $json
    $audit = Get-AccessibilityAndGeometry -RootElement $RootElement -ProductProcessId $product.Id
    $rootWindowGeometry = Get-NativeWindowGeometry -Window $RootElement
    if ($null -ne $rootWindowGeometry -and -not $rootWindowGeometry.fullyInsideMonitorWorkingArea) { $audit.geometryFailures += "Root window is not fully inside its monitor working area: $($RootElement.Current.Name)." }
    $requiredControlFailures = @(Assert-RequiredControls -RootElement $RootElement -ExpectedControls $ExpectedControls)
    $focus = Capture-FocusEvidence -CaptureName $captureName -RootElement $RootElement -ProductProcessId $product.Id -FocusDirectory $focusDirectory
    foreach ($stop in @($focus.stops)) {
        $stop.screenshot = Get-AcceptanceRelativePath -BasePath $output -Path $stop.screenshot -UseForwardSlash
    }
    Write-AcceptanceUtf8File -Path $focusJson -Content ($focus | ConvertTo-Json -Depth 20)
    $semanticFailures = [System.Collections.Generic.List[string]]::new()
    if ($RequireLongPath -and -not (Test-AccessibleTextPresent -RootElement $RootElement -Text $ExpectedLongPath -RequiredProcessId $product.Id)) { $semanticFailures.Add("The complete expected long path is not present in the product UI: $ExpectedLongPath") }
    if ($RequireLongPath -and -not (Test-AccessibleTextAndHelpPresent -RootElement $RootElement -Text $ExpectedLongPath -RequiredProcessId $product.Id)) { $semanticFailures.Add("The deliberately truncated long path does not expose its complete text through product HelpText/tooltip: $ExpectedLongPath") }
    if ($RequireWarning -and -not (Test-AccessibleTextPresent -RootElement $RootElement -Text $ExpectedWarningText -RequiredProcessId $product.Id)) { $semanticFailures.Add("The expected warning text is not present in the product UI: $ExpectedWarningText") }
    if (-not [string]::IsNullOrWhiteSpace($RequiredText) -and -not (Test-AccessibleTextPresent -RootElement $RootElement -Text $RequiredText -RequiredProcessId $product.Id)) { $semanticFailures.Add("The expected state text is not present in the product UI: $RequiredText") }
    $allFailures = @($audit.geometryFailures) + @($audit.overlapFailures) + @($audit.accessibilityFailures) + $requiredControlFailures + @($focus.failures) + @($semanticFailures); $bounds = $RootElement.Current.BoundingRectangle
    $results.Add([pscustomobject]@{
        state = $Name; stateCategory = $Category; phase = $CapturePhase; applicationTheme = $ApplicationTheme; windowsAppTheme = $environment.windowsAppTheme; highContrast = $environment.highContrast; windowSizeMode = $WindowSizeMode; testedWindowGeometry = $windowSizeMeasurement; rootWindowGeometry = $rootWindowGeometry; observedStateText = $ObservedStateText; stateSemanticVerified = [bool]$StateSemanticVerified; capturedUtc = [DateTime]::UtcNow.ToString('O'); passed = $allFailures.Count -eq 0
        geometryFailures = @($audit.geometryFailures); overlapFailures = @($audit.overlapFailures); accessibilityFailures = @($audit.accessibilityFailures) + $requiredControlFailures; focusFailures = @($focus.failures); semanticFailures = @($semanticFailures); offscreenInteractive = @($audit.offscreenInteractive); interactive = @($audit.interactive); rootBounds = Get-RectangleObject $bounds
        longPathAccessible = if ($RequireLongPath) { Test-AccessibleTextPresent -RootElement $RootElement -Text $ExpectedLongPath -RequiredProcessId $product.Id } else { $null }; longPathTooltipAccessible = if ($RequireLongPath) { Test-AccessibleTextAndHelpPresent -RootElement $RootElement -Text $ExpectedLongPath -RequiredProcessId $product.Id } else { $null }; warningAccessible = if ($RequireWarning) { Test-AccessibleTextPresent -RootElement $RootElement -Text $ExpectedWarningText -RequiredProcessId $product.Id } else { $null }
        screenshot = Get-AcceptanceRelativePath -BasePath $output -Path $png -UseForwardSlash; screenshotSha256 = Get-AcceptanceFileSha256 -Path $png; uiaTree = Get-AcceptanceRelativePath -BasePath $output -Path $json -UseForwardSlash; uiaTreeSha256 = Get-AcceptanceFileSha256 -Path $json; focusEvidence = Get-AcceptanceRelativePath -BasePath $output -Path $focusJson -UseForwardSlash; focusEvidenceSha256 = Get-AcceptanceFileSha256 -Path $focusJson
    })
}

function Capture-DialogAndCancel {
    param([string] $Title,[string] $Name,[string] $Category)
    $dialog = Find-UiaWindow -Title $Title -TimeoutSeconds 15; Capture-State -Name $Name -Category $Category -RootElement $dialog -ApplicationTheme 'Windows'
    $cancel = Find-UiaElement -Root $dialog -Name 'Cancel' -Optional
    if ($cancel) { Invoke-UiaElement -Element $cancel } else { Send-UiaKeys -Window $dialog -Keys '{ESC}' }
    Start-Sleep -Milliseconds 250
}

$mainControls = @{
    DashboardNavigationButton = @{ name = 'Dashboard'; role = 'ControlType.Button' }; WorkbooksNavigationButton = @{ name = 'Workbooks'; role = 'ControlType.Button' }; HistoryNavigationButton = @{ name = 'History'; role = 'ControlType.Button' }; SettingsNavigationButton = @{ name = 'Settings'; role = 'ControlType.Button' }; AboutNavigationButton = @{ name = 'About'; role = 'ControlType.Button' }
}

if ($CapturePhase -eq 'Onboarding') {
    if ([string]::IsNullOrWhiteSpace($WorkbookPath)) { throw '-WorkbookPath is mandatory for the onboarding visual phase.' }
    $resolvedWorkbook = (Resolve-Path $WorkbookPath).Path
    if ($resolvedWorkbook.Length -lt 80) { throw 'The onboarding fixture path must be at least 80 characters so long-path rendering is exercised.' }
    $onboarding = Find-UiaWindow -Title 'Welcome to Excel Scenario Analysis Tool' -ProcessId $product.Id -TimeoutSeconds 20
    $nextControl = @{ OnboardingNextButton = @{ name = 'Continue setup'; role = 'ControlType.Button' } }
    Capture-State 'onboarding-step-1' 'onboarding' $onboarding 'Startup' $nextControl
    Invoke-UiaElement -Element (Find-UiaElement -Root $onboarding -AutomationId 'OnboardingNextButton'); Start-Sleep -Milliseconds 250
    $step2 = @{ OnboardingNextButton = $nextControl.OnboardingNextButton; OnboardingBackButton = @{ name = 'Back'; role = 'ControlType.Button' }; ChooseReportFolderButton = @{ name = 'Choose report folder'; role = 'ControlType.Button' } }
    Capture-State 'onboarding-step-2' 'onboarding' $onboarding 'Startup' $step2
    Invoke-UiaElement -Element (Find-UiaElement -Root $onboarding -AutomationId 'ChooseReportFolderButton'); Capture-DialogAndCancel 'Choose where Markdown reports should be saved' 'onboarding-report-folder-picker' 'folder-picker'
    Invoke-UiaElement -Element (Find-UiaElement -Root $onboarding -AutomationId 'OnboardingNextButton'); Start-Sleep -Milliseconds 250
    $step3 = @{ OnboardingNextButton = $nextControl.OnboardingNextButton; OnboardingBackButton = $step2.OnboardingBackButton; ChooseWorkbookButton = @{ name = 'Choose workbook'; role = 'ControlType.Button' } }
    Capture-State 'onboarding-step-3-empty' 'onboarding' $onboarding 'Startup' $step3
    Invoke-UiaElement -Element (Find-UiaElement -Root $onboarding -AutomationId 'ChooseWorkbookButton')
    $fileDialog = Find-UiaWindow -Title 'Choose a workbook to track' -TimeoutSeconds 15; Capture-State 'onboarding-workbook-picker' 'workbook-picker' $fileDialog 'Windows'
    $fileName = Find-UiaElement -Root $fileDialog -AutomationId '1148' -Optional; if (-not $fileName) { $fileName = Find-UiaElement -Root $fileDialog -Name 'File name:' }; Set-UiaValue -Element $fileName -Value $resolvedWorkbook
    $open = Find-UiaElement -Root $fileDialog -AutomationId '1' -Optional; if (-not $open) { $open = Find-UiaElement -Root $fileDialog -Name 'Open' }; Invoke-UiaElement -Element $open; Start-Sleep -Milliseconds 400
    $script:ExpectedLongPath = $resolvedWorkbook; Capture-State 'onboarding-step-3-long-path' 'long-path' $onboarding 'Startup' $step3 -RequireLongPath
    Invoke-UiaElement -Element (Find-UiaElement -Root $onboarding -AutomationId 'OnboardingNextButton'); Start-Sleep -Milliseconds 250
    $step4 = @{ OnboardingNextButton = $nextControl.OnboardingNextButton; OnboardingBackButton = $step2.OnboardingBackButton; OnboardingStartWithWindowsCheckBox = @{ name = 'Start Excel Scenario Analysis Tool with Windows'; role = 'ControlType.CheckBox' }; OnboardingBeginTrackingCheckBox = @{ name = 'Begin tracking the selected workbook'; role = 'ControlType.CheckBox' } }
    Capture-State 'onboarding-step-4' 'onboarding' $onboarding 'Startup' $step4
    Invoke-UiaElement -Element (Find-UiaElement -Root $onboarding -AutomationId 'OnboardingNextButton')
    Wait-AcceptanceCondition -TimeoutSeconds 60 -FailureMessage 'Onboarding did not reach its completion state.' -Condition { $null -ne (Find-UiaElement -Root $onboarding -AutomationId 'OnboardingCompleteHeading' -Optional) }
    Capture-State 'onboarding-step-5' 'onboarding' $onboarding 'Startup' $nextControl
}
elseif ($CapturePhase -eq 'Main') {
    if ([string]::IsNullOrWhiteSpace($ExpectedLongPath) -or $ExpectedLongPath.Length -lt 80) { throw '-ExpectedLongPath (at least 80 characters) is mandatory for the main visual phase.' }
    if ([string]::IsNullOrWhiteSpace($ExpectedWarningText)) { throw '-ExpectedWarningText is mandatory for the main visual phase.' }
    $main = Find-UiaWindow -Title 'Excel Scenario Analysis Tool' -ProcessId $product.Id -TimeoutSeconds 20; $themes = if ($environment.highContrast) { @('System') } else { @('Light','Dark','System') }
    foreach ($theme in $themes) {
        Invoke-UiaElement -Element (Find-UiaElement -Root $main -AutomationId 'SettingsNavigationButton'); Start-Sleep -Milliseconds 250
        $combo = Find-UiaElement -Root $main -AutomationId 'ThemeComboBox'; Invoke-UiaElement -Element $combo; Start-Sleep -Milliseconds 150; Invoke-UiaElement -Element (Find-UiaElement -Root $main -Name $theme); Start-Sleep -Milliseconds 450; Assert-SelectedTheme -ComboBox $combo -Theme $theme
        foreach ($page in @('Dashboard','Workbooks','History','Settings','About')) {
            Invoke-UiaElement -Element (Find-UiaElement -Root $main -AutomationId "${page}NavigationButton"); Start-Sleep -Milliseconds 250; Assert-SelectedTheme -ComboBox $combo -Theme $theme; $controls = $mainControls.Clone()
            if ($page -eq 'Dashboard') { $controls.DashboardAddWorkbookButton = @{ name = 'Add workbook'; role = 'ControlType.Button' } }
            if ($page -eq 'Workbooks') { $controls.WorkbooksAddWorkbookButton = @{ name = 'Add workbook'; role = 'ControlType.Button' } }
            if ($page -eq 'History') { $controls.HistoryWorkbookFilter = @{ name = 'Filter history by workbook'; role = 'ControlType.ComboBox' }; $controls.ShowValueHistoryCheckBox = @{ name = 'Show value changes'; role = 'ControlType.CheckBox' }; $controls.ShowFormulaHistoryCheckBox = @{ name = 'Show formula changes'; role = 'ControlType.CheckBox' }; $controls.ShowSheetHistoryCheckBox = @{ name = 'Show sheet changes'; role = 'ControlType.CheckBox' }; $controls.ShowErrorsCheckBox = @{ name = 'Show capture errors'; role = 'ControlType.CheckBox' } }
            if ($page -eq 'Settings') { $controls.ThemeComboBox = @{ name = 'Application theme'; role = 'ControlType.ComboBox' }; $controls.StartWithWindowsCheckBox = @{ name = 'Start Excel Scenario Analysis Tool with Windows'; role = 'ControlType.CheckBox' }; $controls.ChooseDefaultReportFolderButton = @{ name = 'Choose default report folder'; role = 'ControlType.Button' } }
            Capture-State "$($theme.ToLowerInvariant())-$($page.ToLowerInvariant())-$($WindowSizeMode.ToLowerInvariant())" 'main-page' $main $theme $controls -RequireLongPath:($page -in @('Dashboard','Workbooks','History')) -RequireWarning:($page -eq 'History')
        }
    }
    Invoke-UiaElement -Element (Find-UiaElement -Root $main -AutomationId 'DashboardNavigationButton'); Invoke-UiaElement -Element (Find-UiaElement -Root $main -AutomationId 'DashboardAddWorkbookButton'); Capture-DialogAndCancel 'Choose a workbook to track' "main-workbook-picker-$($WindowSizeMode.ToLowerInvariant())" 'workbook-picker'
    Invoke-UiaElement -Element (Find-UiaElement -Root $main -AutomationId 'SettingsNavigationButton'); Invoke-UiaElement -Element (Find-UiaElement -Root $main -AutomationId 'ChooseDefaultReportFolderButton'); Capture-DialogAndCancel 'Choose where Excel Scenario Analysis Tool should save Markdown reports' "main-report-folder-picker-$($WindowSizeMode.ToLowerInvariant())" 'folder-picker'
    Invoke-UiaElement -Element (Find-UiaElement -Root $main -AutomationId 'WorkbooksNavigationButton')
    $more = Find-UiaElement -Root $main -Name 'Workbook actions' -Optional
    if (-not $more) { throw 'A populated workbook actions menu is required to capture the purge confirmation dialog.' }
    Invoke-UiaElement -Element $more; Start-Sleep -Milliseconds 200
    $purge = @([System.Windows.Automation.AutomationElement]::RootElement.FindAll([System.Windows.Automation.TreeScope]::Descendants,[System.Windows.Automation.Condition]::TrueCondition) | Where-Object {
        $_.Current.ProcessId -eq $product.Id -and -not $_.Current.IsOffscreen -and $_.Current.Name -eq ('Permanently purge history' + [char]0x2026) -and $_.Current.ControlType.ProgrammaticName -eq 'ControlType.MenuItem'
    }) | Select-Object -First 1
    if (-not $purge) { throw 'The open workbook actions menu does not expose Permanently purge history.' }
    Invoke-UiaElement -Element $purge; $purgeDialog = Find-UiaWindow -Title 'Permanently delete history?' -TimeoutSeconds 10; Capture-State "purge-confirmation-$($WindowSizeMode.ToLowerInvariant())" 'confirmation-dialog' $purgeDialog 'Windows'; Invoke-UiaElement -Element (Find-UiaElement -Root $purgeDialog -Name 'No')
}
else {
    if ([string]::IsNullOrWhiteSpace($StateName)) { throw '-StateName is mandatory for a supplemental capture.' }
    if ($UseDesktopRoot -and $StateName -ne 'tray-menu') { throw '-UseDesktopRoot is permitted only for a tray-menu capture.' }
    $target = if ($UseDesktopRoot) { [System.Windows.Automation.AutomationElement]::RootElement } else { Find-UiaWindow -Title $TargetWindowTitle -ProcessId $product.Id -TimeoutSeconds 15 }
    if ($StateName -eq 'tray-menu') {
        foreach ($menuText in @('Open Excel Scenario Analysis Tool',('Add workbook' + [char]0x2026),'Exit')) {
            $menuItems = @($target.FindAll([System.Windows.Automation.TreeScope]::Descendants,[System.Windows.Automation.Condition]::TrueCondition) | Where-Object { $_.Current.ProcessId -eq $product.Id -and $_.Current.Name -eq $menuText -and $_.Current.ControlType.ProgrammaticName -eq 'ControlType.MenuItem' })
            if ($menuItems.Count -ne 1) { throw "The product-owned open tray menu does not expose exactly one required MenuItem: $menuText" }
        }
        Capture-State "supplemental-$StateName-$($WindowSizeMode.ToLowerInvariant())" $StateName $target 'Observed' -ObservedStateText 'Open Excel Scenario Analysis Tool|Add workbook|Exit' -StateSemanticVerified
    }
    elseif ($StateName -eq 'toast') {
        $toast = Find-UiaElement -Root $target -AutomationId 'StatusToast' -Optional
        if ($null -eq $toast -or $toast.Current.IsOffscreen -or [string]::IsNullOrWhiteSpace($toast.Current.Name)) { throw 'No visible, named StatusToast is present; the supplemental label is not accepted as evidence.' }
        Capture-State "supplemental-$StateName-$($WindowSizeMode.ToLowerInvariant())" $StateName $target 'Observed' -RequiredText $toast.Current.Name -ObservedStateText $toast.Current.Name -StateSemanticVerified
    }
    elseif ($StateName -eq 'long-path') {
        if ([string]::IsNullOrWhiteSpace($ExpectedLongPath) -or $ExpectedLongPath.Length -lt 80) { throw '-ExpectedLongPath (at least 80 characters) is required for a supplemental long-path state.' }
        Capture-State "supplemental-$StateName-$($WindowSizeMode.ToLowerInvariant())" $StateName $target 'Observed' -RequireLongPath -ObservedStateText $ExpectedLongPath -StateSemanticVerified
    }
    else {
        if ($StateName -eq 'processing') {
            if ($ExpectedStateText -ne 'Processing' -or @(Get-ExactProductTextControls -RootElement $target -Text 'Processing' -ProductProcessId $product.Id).Count -lt 1) { throw "Processing evidence requires a visible product Text control whose exact state is 'Processing'." }
        }
        elseif ($StateName -eq 'warning') {
            if ($ExpectedStateText -ne 'Warning' -or @(Get-ExactProductTextControls -RootElement $target -Text 'Warning' -ProductProcessId $product.Id).Count -lt 1) { throw "Warning evidence requires a visible product Text control whose exact workbook status is 'Warning'." }
        }
        else {
            $allowedErrorPrefixes = @('Missing workbook','Unsupported workbook','Corrupt, encrypted, or unsafe workbook','Permission denied','Workbook temporarily unavailable','Capture failed')
            if ([string]::IsNullOrWhiteSpace($ExpectedStateText) -or -not @($allowedErrorPrefixes | Where-Object { $ExpectedStateText.StartsWith($_,[StringComparison]::Ordinal) }).Count) { throw 'Error evidence requires the exact visible capture-error category/stage, not arbitrary product text.' }
            $historyPage = Find-UiaElement -Root $target -AutomationId 'HistoryPage' -Optional
            if ($null -eq $historyPage -or $historyPage.Current.IsOffscreen -or @(Get-ExactProductTextControls -RootElement $target -Text $ExpectedStateText -ProductProcessId $product.Id).Count -ne 1) { throw 'Error evidence requires one exact visible error-category Text control on the active History page.' }
        }
        Capture-State "supplemental-$StateName-$($WindowSizeMode.ToLowerInvariant())" $StateName $target 'Observed' -RequiredText $ExpectedStateText -ObservedStateText $ExpectedStateText -StateSemanticVerified
    }
}

$failedResults = @($results | Where-Object { -not $_.passed })
$summary = [ordered]@{
    schemaVersion = 2; status = if ($failedResults.Count -eq 0) { 'MachineChecksPassedHumanReviewRequired' } else { 'Failed' }; captureSessionId = $sessionGuid; configuration = $configuration; operatorDisplayLabel = $OperatorDisplayLabel; expectedScalePercent = [int]$ExpectedScalePercent; windowsContrastThemeRequested = [bool]$WindowsContrastTheme
    installerSha256 = $expectedInstallerHash; installerPathAtCapture = $resolvedInstaller; applicationSha256 = $expectedApplicationHash; applicationPath = $product.Path; applicationProcessId = $product.Id; firstCapturedUtc = $firstCapturedUtc; capturedUtc = [DateTime]::UtcNow.ToString('O'); actualEnvironment = $environment; results = @($results)
    humanReview = [ordered]@{ required = $true; contrast = 'Pending'; clippingAndOverlap = 'Pending'; keyboardAndFocus = 'Pending'; themeTransitions = 'Pending'; reviewer = $null; reviewedUtc = $null; notes = $null }
}
Write-AcceptanceUtf8File -Path $summaryPath -Content ($summary | ConvertTo-Json -Depth 40)
if ($summary.status -eq 'Failed') { throw "VISUAL_MATRIX_MACHINE_CHECK_FAILED|$summaryPath" }
Write-Output "VISUAL_MATRIX_CAPTURED_REVIEW_REQUIRED|$summaryPath"
