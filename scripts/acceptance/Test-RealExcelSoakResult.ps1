[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string] $ResultPath,
    [Parameter(Mandatory)] [string] $InstallerPath,
    [Parameter(Mandatory)] [ValidatePattern('^[A-Fa-f0-9]{64}$')] [string] $ExpectedApplicationSha256
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$path = (Resolve-Path $ResultPath).Path
$root = Split-Path -Parent $path
$installerHash = (Get-FileHash (Resolve-Path $InstallerPath).Path -Algorithm SHA256).Hash.ToUpperInvariant()
$applicationHash = $ExpectedApplicationSha256.ToUpperInvariant()
$result = Get-Content $path -Raw | ConvertFrom-Json

function Require-Condition {
    param([bool] $Condition, [string] $Message)
    if (-not $Condition) { throw "Invalid real-Excel soak evidence: $Message" }
}

function Require-EvidenceFile {
    param([string] $RelativePath)
    Require-Condition (-not [string]::IsNullOrWhiteSpace($RelativePath)) 'an evidence path is empty'
    Require-Condition (-not [System.IO.Path]::IsPathRooted($RelativePath)) "evidence path must be relative: $RelativePath"
    $fullPath = [System.IO.Path]::GetFullPath((Join-Path $root $RelativePath))
    $rootWithSeparator = $root.TrimEnd('\') + '\'
    Require-Condition ($fullPath.StartsWith($rootWithSeparator, [StringComparison]::OrdinalIgnoreCase)) "evidence path escapes the soak directory: $RelativePath"
    Require-Condition (Test-Path $fullPath -PathType Leaf) "evidence file is missing: $RelativePath"
    $fullPath
}

Require-Condition ($result.schemaVersion -eq 1) 'schemaVersion must be 1'
Require-Condition ($result.gate -eq 'real-excel-ten-minute-soak') 'gate identity is wrong'
Require-Condition ($result.status -eq 'Passed') 'status must be Passed'
Require-Condition ($result.saveCount -eq 20) 'exactly 20 saves are required'
Require-Condition ($result.saveIntervalSeconds -ge 30) 'save interval must be at least 30 seconds'
Require-Condition ($result.durationSeconds -ge 600) 'soak duration must be at least ten minutes'
Require-Condition ($result.candidate.installerSha256 -eq $installerHash) 'installer hash does not match the exact candidate'
Require-Condition ($result.candidate.expectedInstallerSha256 -eq $installerHash) 'frozen installer hash differs'
Require-Condition ($result.candidate.applicationSha256 -eq $applicationHash) 'installed executable hash differs'
Require-Condition ($result.candidate.expectedApplicationSha256 -eq $applicationHash) 'frozen executable hash differs'

$workbooks = @($result.workbooks)
Require-Condition ($workbooks.Count -eq 2) 'exactly two workbooks are required'
Require-Condition (@($workbooks | Where-Object format -eq 'xlsx').Count -eq 1) 'one xlsx workbook is required'
Require-Condition (@($workbooks | Where-Object format -eq 'xlsm').Count -eq 1) 'one xlsm workbook is required'
Require-Condition (@($workbooks | Where-Object finalSequence -ne 10).Count -eq 0) 'each workbook must finish at sequence 10'
$macro = $workbooks | Where-Object format -eq 'xlsm'
Require-Condition (-not [string]::IsNullOrWhiteSpace($macro.macroHashBefore)) 'xlsm macro hash is missing'
Require-Condition ($macro.macroHashBefore -eq $macro.macroHashAfter) 'xlsm VBA project changed'

$saves = @($result.saves)
Require-Condition ($saves.Count -eq 20) 'exactly 20 save records are required'
$firstStarted = [DateTime]::Parse($saves[0].saveStartedUtc).ToUniversalTime()
$lastStarted = [DateTime]::Parse($saves[-1].saveStartedUtc).ToUniversalTime()
Require-Condition (($lastStarted - $firstStarted).TotalSeconds -ge 600) 'first-to-last real save span is under ten minutes'

foreach ($format in @('xlsx', 'xlsm')) {
    $formatSaves = @($saves | Where-Object format -eq $format)
    Require-Condition ($formatSaves.Count -eq 10) "$format must have ten saves"
    $expectedSequences = @(1..10)
    Require-Condition (@(Compare-Object $expectedSequences @($formatSaves.sequence)).Count -eq 0) "$format sequence set is not exactly 1 through 10"
    Require-Condition (@($formatSaves.sha256 | Select-Object -Unique).Count -eq 10) "$format stable hashes are missing or duplicated"
}

for ($index = 0; $index -lt $saves.Count; $index++) {
    $save = $saves[$index]
    $wantFormat = if (($index % 2) -eq 0) { 'xlsx' } else { 'xlsm' }
    Require-Condition ($save.index -eq ($index + 1)) "save index $index is out of order"
    Require-Condition ($save.format -eq $wantFormat) "save $($index + 1) does not alternate workbook formats"
    Require-Condition ($save.value -eq ('EDT-SOAK-{0:D2}' -f ($index + 1))) "save $($index + 1) value is wrong"
    Require-Condition ($save.sha256 -match '^[A-F0-9]{64}$') "save $($index + 1) hash is malformed"
    Require-Condition ($save.captureMilliseconds -le 20000) "save $($index + 1) exceeded the 20-second capture limit"
    $probePath = Require-EvidenceFile $save.probe
    $probe = Get-Content $probePath -Raw | ConvertFrom-Json
    Require-Condition $probe.passed "probe failed for save $($index + 1)"
    Require-Condition ($probe.workbookStatus -eq 'Active') "workbook was not Active after save $($index + 1)"
    Require-Condition ($probe.currentSequence -eq $save.sequence) "probe sequence differs for save $($index + 1)"
    Require-Condition ($probe.currentHash -eq $save.sha256) "probe hash differs for save $($index + 1)"
    Require-Condition ($probe.latestVersion.sha256 -eq $save.sha256) "latest-version hash differs for save $($index + 1)"
    Require-Condition ($probe.latestVersion.reportStatus -eq 'Ready') "report was not Ready for save $($index + 1)"
    Require-Condition ($probe.errorCount -eq 0) "capture errors exist after save $($index + 1)"
    Require-Condition ($probe.versionCount -eq $save.sequence) "version count differs after save $($index + 1)"
    Require-Condition ($probe.distinctVersionHashCount -eq $probe.versionCount) "duplicate version hash exists after save $($index + 1)"
}

$failedAssertions = @($result.assertions | Where-Object { -not $_.passed })
Require-Condition ($failedAssertions.Count -eq 0) 'result contains a failed assertion'
Require-Condition ([string]::IsNullOrWhiteSpace($result.failure)) 'result contains an unhandled failure'
Require-EvidenceFile 'screenshots/soak-history.png' | Out-Null
Require-EvidenceFile 'uia/soak-history.json' | Out-Null

Write-Output "REAL_EXCEL_SOAK_VALID|$path|$installerHash|$applicationHash"
