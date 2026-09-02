Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms
if (-not ('ExcelDiffTrackerAcceptanceNative' -as [type])) {
Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class ExcelDiffTrackerAcceptanceNative
{
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool SetCursorPos(int x, int y);

    [DllImport("user32.dll")]
    public static extern void mouse_event(uint flags, uint dx, uint dy, uint data, UIntPtr extraInfo);
}
"@
}

function Get-UiaRoot {
    [System.Windows.Automation.AutomationElement]::RootElement
}

function Write-AcceptanceUtf8File {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Path,
        [Parameter(Mandatory)] [AllowEmptyString()] [string] $Content
    )
    $directory = Split-Path -Parent $Path
    if (-not [string]::IsNullOrWhiteSpace($directory)) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }
    $encoding = New-Object System.Text.UTF8Encoding -ArgumentList $false
    [System.IO.File]::WriteAllText([System.IO.Path]::GetFullPath($Path), $Content, $encoding)
}

function Get-AcceptanceRelativePath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $BasePath,
        [Parameter(Mandatory)] [string] $Path,
        [switch] $UseForwardSlash
    )
    $baseFullPath = [System.IO.Path]::GetFullPath($BasePath)
    $targetFullPath = [System.IO.Path]::GetFullPath($Path)
    $trimCharacters = [char[]]@(
        [System.IO.Path]::DirectorySeparatorChar,
        [System.IO.Path]::AltDirectorySeparatorChar)
    $baseWithoutSeparator = $baseFullPath.TrimEnd($trimCharacters)
    if ($targetFullPath.TrimEnd($trimCharacters).Equals($baseWithoutSeparator, [StringComparison]::OrdinalIgnoreCase)) {
        return '.'
    }
    $basePrefix = $baseWithoutSeparator + [System.IO.Path]::DirectorySeparatorChar
    if (-not $targetFullPath.StartsWith($basePrefix, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Path is outside the required base directory: '$targetFullPath' is not under '$baseFullPath'."
    }
    $relative = $targetFullPath.Substring($basePrefix.Length)
    if ($UseForwardSlash) { return $relative.Replace('\', '/') }
    $relative
}

function Get-AcceptanceFileSha256 {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [string] $Path)
    (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToUpperInvariant()
}

function Get-AcceptanceStreamSha256 {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [System.IO.Stream] $Stream)
    $algorithm = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = $algorithm.ComputeHash($Stream)
        -join @($bytes | ForEach-Object { $_.ToString('X2') })
    }
    finally {
        $algorithm.Dispose()
    }
}

function Find-UiaWindow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Title,
        [int] $ProcessId = 0,
        [int] $TimeoutSeconds = 20
    )

    $deadline = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
    do {
        $windows = (Get-UiaRoot).FindAll(
            [System.Windows.Automation.TreeScope]::Children,
            [System.Windows.Automation.Condition]::TrueCondition)
        foreach ($window in $windows) {
            if ($window.Current.Name -eq $Title -and ($ProcessId -eq 0 -or $window.Current.ProcessId -eq $ProcessId)) { return $window }
        }
        Start-Sleep -Milliseconds 200
    } while ([DateTime]::UtcNow -lt $deadline)
    throw "UI Automation window was not found within $TimeoutSeconds seconds: $Title"
}

function Get-UiaWindowFromHandle {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [long] $Handle)
    $window = [System.Windows.Automation.AutomationElement]::FromHandle([IntPtr]$Handle)
    if (-not $window) { throw "UI Automation could not resolve window handle $Handle." }
    $window
}

function Find-UiaElement {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [System.Windows.Automation.AutomationElement] $Root,
        [string] $AutomationId,
        [string] $Name,
        [switch] $Optional
    )

    if (-not $AutomationId -and -not $Name) { throw 'AutomationId or Name is required.' }
    $conditions = [System.Collections.Generic.List[System.Windows.Automation.Condition]]::new()
    if ($AutomationId) {
        $conditions.Add([System.Windows.Automation.PropertyCondition]::new(
            [System.Windows.Automation.AutomationElement]::AutomationIdProperty,
            $AutomationId))
    }
    if ($Name) {
        $conditions.Add([System.Windows.Automation.PropertyCondition]::new(
            [System.Windows.Automation.AutomationElement]::NameProperty,
            $Name))
    }
    $condition = if ($conditions.Count -eq 1) {
        $conditions[0]
    } else {
        [System.Windows.Automation.AndCondition]::new($conditions.ToArray())
    }
    $element = $Root.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $condition)
    if (-not $element -and -not $Optional) {
        throw "UI element was not found. AutomationId='$AutomationId' Name='$Name'."
    }
    $element
}

function Invoke-UiaElement {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [System.Windows.Automation.AutomationElement] $Element)
    $pattern = $null
    if ($Element.TryGetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern, [ref]$pattern)) {
        ([System.Windows.Automation.InvokePattern]$pattern).Invoke()
        return
    }
    if ($Element.TryGetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern, [ref]$pattern)) {
        ([System.Windows.Automation.SelectionItemPattern]$pattern).Select()
        return
    }
    if ($Element.TryGetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern, [ref]$pattern)) {
        ([System.Windows.Automation.TogglePattern]$pattern).Toggle()
        return
    }
    if ($Element.TryGetCurrentPattern([System.Windows.Automation.ExpandCollapsePattern]::Pattern, [ref]$pattern)) {
        $expand = [System.Windows.Automation.ExpandCollapsePattern]$pattern
        if ($expand.Current.ExpandCollapseState -eq [System.Windows.Automation.ExpandCollapseState]::Collapsed) { $expand.Expand() }
        else { $expand.Collapse() }
        return
    }
    throw "Element does not expose an invokable UIA pattern: $($Element.Current.AutomationId) $($Element.Current.Name)"
}

function Invoke-UiaMouseClick {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [System.Windows.Automation.AutomationElement] $Element,
        [ValidateSet('Left','Right','DoubleLeft')] [string] $Button = 'Left'
    )
    $bounds = $Element.Current.BoundingRectangle
    if ($bounds.IsEmpty -or $bounds.Width -le 0 -or $bounds.Height -le 0) {
        throw "Element has no clickable bounds: $($Element.Current.AutomationId) $($Element.Current.Name)"
    }
    $x = [int][math]::Round($bounds.Left + ($bounds.Width / 2.0))
    $y = [int][math]::Round($bounds.Top + ($bounds.Height / 2.0))
    if (-not [ExcelDiffTrackerAcceptanceNative]::SetCursorPos($x, $y)) { throw "Could not position the pointer at $x,$y." }
    Start-Sleep -Milliseconds 100
    if ($Button -eq 'Right') {
        [ExcelDiffTrackerAcceptanceNative]::mouse_event(0x0008, 0, 0, 0, [UIntPtr]::Zero)
        [ExcelDiffTrackerAcceptanceNative]::mouse_event(0x0010, 0, 0, 0, [UIntPtr]::Zero)
    } else {
        $count = if ($Button -eq 'DoubleLeft') { 2 } else { 1 }
        for ($index = 0; $index -lt $count; $index++) {
            [ExcelDiffTrackerAcceptanceNative]::mouse_event(0x0002, 0, 0, 0, [UIntPtr]::Zero)
            [ExcelDiffTrackerAcceptanceNative]::mouse_event(0x0004, 0, 0, 0, [UIntPtr]::Zero)
            if ($count -gt 1) { Start-Sleep -Milliseconds 80 }
        }
    }
    Start-Sleep -Milliseconds 200
}

function Set-UiaValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [System.Windows.Automation.AutomationElement] $Element,
        [Parameter(Mandatory)] [AllowEmptyString()] [string] $Value
    )
    $pattern = $null
    if (-not $Element.TryGetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern, [ref]$pattern)) {
        throw "Element does not expose ValuePattern: $($Element.Current.AutomationId) $($Element.Current.Name)"
    }
    ([System.Windows.Automation.ValuePattern]$pattern).SetValue($Value)
}

function Set-UiaForeground {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [System.Windows.Automation.AutomationElement] $Window)
    $handle = [IntPtr]$Window.Current.NativeWindowHandle
    if ($handle -eq [IntPtr]::Zero) { throw "Window has no native handle: $($Window.Current.Name)" }
    [void][ExcelDiffTrackerAcceptanceNative]::ShowWindow($handle, 9)
    if (-not [ExcelDiffTrackerAcceptanceNative]::SetForegroundWindow($handle)) {
        throw "Could not foreground window: $($Window.Current.Name)"
    }
    Start-Sleep -Milliseconds 250
}

function Send-UiaKeys {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [System.Windows.Automation.AutomationElement] $Window,
        [Parameter(Mandatory)] [string] $Keys
    )
    Set-UiaForeground -Window $Window
    [System.Windows.Forms.SendKeys]::SendWait($Keys)
    Start-Sleep -Milliseconds 150
}

function Convert-UiaNode {
    param(
        [Parameter(Mandatory)] [System.Windows.Automation.AutomationElement] $Element,
        [int] $Depth,
        [int] $MaxDepth
    )
    $rectangle = $Element.Current.BoundingRectangle
    $node = [ordered]@{
        name = $Element.Current.Name
        automationId = $Element.Current.AutomationId
        controlType = $Element.Current.ControlType.ProgrammaticName
        className = $Element.Current.ClassName
        enabled = $Element.Current.IsEnabled
        offscreen = $Element.Current.IsOffscreen
        bounds = [ordered]@{ x = $rectangle.X; y = $rectangle.Y; width = $rectangle.Width; height = $rectangle.Height }
        children = @()
    }
    if ($Depth -lt $MaxDepth) {
        $children = $Element.FindAll(
            [System.Windows.Automation.TreeScope]::Children,
            [System.Windows.Automation.Condition]::TrueCondition)
        $node.children = @($children | ForEach-Object { Convert-UiaNode -Element $_ -Depth ($Depth + 1) -MaxDepth $MaxDepth })
    }
    [pscustomobject]$node
}

function Export-UiaTree {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [System.Windows.Automation.AutomationElement] $Root,
        [Parameter(Mandatory)] [string] $Path,
        [int] $MaxDepth = 12
    )
    $directory = Split-Path -Parent $Path
    New-Item -ItemType Directory -Path $directory -Force | Out-Null
    $json = Convert-UiaNode -Element $Root -Depth 0 -MaxDepth $MaxDepth | ConvertTo-Json -Depth 30
    Write-AcceptanceUtf8File -Path $Path -Content $json
}

function Save-DesktopScreenshot {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [string] $Path)
    $directory = Split-Path -Parent $Path
    New-Item -ItemType Directory -Path $directory -Force | Out-Null
    $screen = [System.Windows.Forms.SystemInformation]::VirtualScreen
    $bitmap = [System.Drawing.Bitmap]::new($screen.Width, $screen.Height)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    try {
        $graphics.CopyFromScreen($screen.Left, $screen.Top, 0, 0, $bitmap.Size)
        $bitmap.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
    }
    finally {
        $graphics.Dispose()
        $bitmap.Dispose()
    }
}

function Wait-AcceptanceCondition {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [scriptblock] $Condition,
        [Parameter(Mandatory)] [string] $FailureMessage,
        [int] $TimeoutSeconds = 20,
        [int] $PollMilliseconds = 250
    )
    $deadline = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
    do {
        if (& $Condition) { return }
        Start-Sleep -Milliseconds $PollMilliseconds
    } while ([DateTime]::UtcNow -lt $deadline)
    throw $FailureMessage
}

Export-ModuleMember -Function Find-UiaWindow, Get-UiaWindowFromHandle, Find-UiaElement, Invoke-UiaElement, Invoke-UiaMouseClick, Set-UiaValue, Set-UiaForeground, Send-UiaKeys, Export-UiaTree, Save-DesktopScreenshot, Wait-AcceptanceCondition, Write-AcceptanceUtf8File, Get-AcceptanceRelativePath, Get-AcceptanceFileSha256, Get-AcceptanceStreamSha256
