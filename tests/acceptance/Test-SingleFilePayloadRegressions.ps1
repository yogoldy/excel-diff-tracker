[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repository = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$validator = Join-Path $repository 'scripts\acceptance\Test-SingleFilePayload.ps1'
$testRoot = Join-Path $repository ('TestResults\single-file-regression\' + [Guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path $testRoot
$checks = 0
foreach ($extra in @('', 'ExcelDiffTracker.Core.dll', 'fr\PresentationCore.resources.dll', 'hidden.dll', 'ExcelDiffTracker.deps.json', 'ExcelDiffTracker.runtimeconfig.json', 'ExcelDiffTracker.runtimeconfig.dev.json')) {
    $directory = Join-Path $testRoot ('case-' + $checks)
    $null = New-Item -ItemType Directory -Path $directory
    $executable = Join-Path $directory 'ExcelDiffTracker.exe'
    # Synthetic directory fixtures only; these files are never executed.
    [System.IO.File]::WriteAllText($executable, 'not-an-executable')
    [System.IO.File]::WriteAllText((Join-Path $directory 'symbols.pdb'), 'symbols')
    if ($extra) {
        $extraPath = Join-Path $directory $extra
        $null = New-Item -ItemType Directory -Path (Split-Path -Parent $extraPath) -Force
        [System.IO.File]::WriteAllText($extraPath, 'unhashed dependency')
        if ($extra -eq 'hidden.dll') { [System.IO.File]::SetAttributes($extraPath, [System.IO.FileAttributes]::Hidden) }
    }
    $failure = $null
    try { $null = & $validator -ExecutablePath $executable } catch { $failure = $_.Exception.Message }
    if ($extra -and ($null -eq $failure -or $failure -notlike 'Loose executable dependencies*')) { throw "Loose payload was not rejected: $extra" }
    if (-not $extra -and $null -ne $failure) { throw "Single-file directory was rejected: $failure" }
    $checks++
}
$allowlist = Join-Path $repository 'installer\LegacyRuntimeFiles-0.1.1.iss'
$ownedPaths = @{}
foreach ($line in Get-Content -LiteralPath $allowlist) {
    if ([string]::IsNullOrWhiteSpace($line) -or $line.StartsWith(';')) { continue }
    if ($line -notmatch '^Type: files; Name: "\{app\}\\([^"*?{}:]+)"$') { throw "Cleanup is not an exact application-owned file: $line" }
    $relative = $Matches[1]
    if ([System.IO.Path]::IsPathRooted($relative) -or $relative -match '(^|\\)\.\.(\\|$)' -or $ownedPaths.ContainsKey($relative)) { throw "Unsafe or duplicate cleanup path: $relative" }
    if (-not $relative.EndsWith('.dll', [StringComparison]::OrdinalIgnoreCase) -and $relative -notin @('createdump.exe', 'ExcelDiffTracker.deps.json', 'ExcelDiffTracker.runtimeconfig.json')) { throw "Cleanup includes a non-runtime file: $relative" }
    $ownedPaths[$relative] = $true
}
if ($ownedPaths.Count -ne 483) { throw 'The pinned legacy runtime inventory changed; review it against the public baseline.' }
$checks++
$installer = [System.IO.File]::ReadAllText((Join-Path $repository 'installer\ExcelDiffTracker.iss'))
if ($installer -notmatch '(?s)\[InstallDelete\].*?#include "LegacyRuntimeFiles-0.1.1.iss"') { throw 'Legacy cleanup is not included in the installer.' }
$checks++
Write-Output "SINGLE_FILE_REGRESSIONS_PASS|checks=$checks|legacyPaths=$($ownedPaths.Count)|evidence=$testRoot"
