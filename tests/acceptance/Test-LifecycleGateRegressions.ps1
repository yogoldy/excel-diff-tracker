[CmdletBinding()]
param(
    [string] $RepositoryRoot
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if ([string]::IsNullOrWhiteSpace($RepositoryRoot)) {
    $RepositoryRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
}
if ($PSVersionTable.PSVersion.Major -ne 5 -or $PSVersionTable.PSVersion.Minor -ne 1 -or $PSVersionTable.PSEdition -ne 'Desktop') {
    throw 'Run these isolated regressions with Windows PowerShell 5.1 Desktop.'
}
Add-Type -AssemblyName UIAutomationClient,UIAutomationTypes
$script:assertionCount = 0

function Assert-Test {
    param([bool] $Condition,[string] $Message)
    if (-not $Condition) { throw "Lifecycle regression failed: $Message" }
    $script:assertionCount++
}

function Assert-Throws {
    param([scriptblock] $Action,[string] $ExpectedMessage)
    $caught = $null
    try { $null = & $Action } catch { $caught = $_ }
    Assert-Test ($null -ne $caught) "Expected a failure containing '$ExpectedMessage'."
    Assert-Test ($caught.Exception.Message.Contains($ExpectedMessage)) "Unexpected failure: $($caught.Exception.Message)"
}

function Get-FunctionSource {
    param([string] $Path,[string] $Name)
    $tokens = $null
    $parseErrors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path,[ref]$tokens,[ref]$parseErrors)
    Assert-Test ($parseErrors.Count -eq 0) "$Path must parse under Windows PowerShell 5.1."
    $definitions = @($ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $Name },$true))
    Assert-Test ($definitions.Count -eq 1) "Expected one $Name definition in $Path."
    $definitions[0].Extent.Text
}

function Test-StartupReader {
    param([string] $Path)
    # Extract definitions only. Never dot-source a gate's top-level install/uninstall code.
    $assertionSource = Get-FunctionSource -Path $Path -Name 'Assert-GateCondition'
    $readerSource = Get-FunctionSource -Path $Path -Name 'Get-ProductStartupRegistration'
    . ([scriptblock]::Create($assertionSource))
    . ([scriptblock]::Create($readerSource))
    $runKey = 'HKCU:\IsolatedLifecycleMock\Run'
    $case = 'AbsentKey'

    function Test-Path {
        [CmdletBinding()]
        param([string] $LiteralPath)
        if ($LiteralPath -ne $runKey) { throw 'Unexpected mock registry path.' }
        if ($case -eq 'KeyReadError') { throw [UnauthorizedAccessException]::new('mock key access denied') }
        $case -ne 'AbsentKey'
    }

    function Get-ItemProperty {
        [CmdletBinding()]
        param([string] $LiteralPath)
        if ($LiteralPath -ne $runKey) { throw 'Unexpected mock registry path.' }
        switch ($case) {
            'AbsentKey' { throw 'An absent key must not be read.' }
            'AbsentValue' { [pscustomobject]@{ PSPath=$runKey }; break }
            'PresentValue' { [pscustomobject]@{ ExcelDiffTracker='"C:\Candidate\ExcelDiffTracker.exe" --background' }; break }
            'EmptyValue' { [pscustomobject]@{ ExcelDiffTracker='' }; break }
            'ValueReadError' { throw [UnauthorizedAccessException]::new('mock value access denied') }
            'NoReadResult' { return $null }
            default { throw 'Unexpected registry test case.' }
        }
    }

    foreach ($case in @('AbsentKey','AbsentValue')) {
        $registration = Get-ProductStartupRegistration
        Assert-Test (-not $registration.exists -and $null -eq $registration.value) "$case must safely report absence under StrictMode."
    }
    $case = 'PresentValue'
    $registration = Get-ProductStartupRegistration
    Assert-Test ($registration.exists -and $registration.value -ceq '"C:\Candidate\ExcelDiffTracker.exe" --background') 'The startup command must remain exact.'
    $case = 'EmptyValue'
    $registration = Get-ProductStartupRegistration
    Assert-Test ($registration.exists -and $registration.value -ceq '') 'A present empty startup value must not count as removed.'
    $case = 'KeyReadError'
    Assert-Throws { Get-ProductStartupRegistration } 'mock key access denied'
    $case = 'ValueReadError'
    Assert-Throws { Get-ProductStartupRegistration } 'mock value access denied'
    $case = 'NoReadResult'
    Assert-Throws { Get-ProductStartupRegistration } 'could not be read'
}

$startPath = Join-Path $RepositoryRoot 'scripts\acceptance\Start-InstalledLifecycleUpgradeGate.ps1'
$completePath = Join-Path $RepositoryRoot 'scripts\acceptance\Complete-InstalledLifecycleUpgradeGate.ps1'
Test-StartupReader -Path $startPath
Test-StartupReader -Path $completePath

$assertionSource = Get-FunctionSource -Path $startPath -Name 'Assert-GateCondition'
$selectorSource = Get-FunctionSource -Path $startPath -Name 'Select-PriorUiaElement'
. ([scriptblock]::Create($assertionSource))
. ([scriptblock]::Create($selectorSource))

function New-MockElement {
    param([int] $ProcessId,[string] $Name,[System.Windows.Automation.ControlType] $ControlType,[bool] $IsOffscreen=$false)
    [pscustomobject]@{ Current=[pscustomobject]@{ ProcessId=$ProcessId; Name=$Name; ControlType=$ControlType; IsOffscreen=$IsOffscreen } }
}

$button = [System.Windows.Automation.ControlType]::Button
$text = [System.Windows.Automation.ControlType]::Text
$target = New-MockElement -ProcessId 1234 -Name 'Continue' -ControlType $button
$decoys = @(
    (New-MockElement -ProcessId 4321 -Name 'Continue' -ControlType $button),
    (New-MockElement -ProcessId 1234 -Name 'Continue' -ControlType $text),
    (New-MockElement -ProcessId 1234 -Name 'continue' -ControlType $button),
    (New-MockElement -ProcessId 1234 -Name 'Continue' -ControlType $button -IsOffscreen $true)
)
$selected = Select-PriorUiaElement -Elements (@($target) + $decoys) -ProcessId 1234 -Name 'Continue' -ControlType $button
Assert-Test ([object]::ReferenceEquals($selected,$target)) 'Prior selection must reject other processes, types, captions, and hidden controls.'
Assert-Throws { Select-PriorUiaElement -Elements $decoys -ProcessId 1234 -Name 'Continue' -ControlType $button } 'found 0'
Assert-Test ($null -eq (Select-PriorUiaElement -Elements $decoys -ProcessId 1234 -Name 'Continue' -ControlType $button -Optional)) 'An optional missing completion heading must remain pending.'
$duplicate = New-MockElement -ProcessId 1234 -Name 'Continue' -ControlType $button
Assert-Throws { Select-PriorUiaElement -Elements @($target,$duplicate) -ProcessId 1234 -Name 'Continue' -ControlType $button } 'found 2'
Assert-Throws { Select-PriorUiaElement -Elements @($target,$duplicate) -ProcessId 1234 -Name 'Continue' -ControlType $button -Optional } 'found 2'

Write-Output "LIFECYCLE_GATE_REGRESSIONS_PASS|assertions=$script:assertionCount|powershell=$($PSVersionTable.PSVersion)|installedActions=none|registryWrites=none"
