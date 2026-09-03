[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string] $ResultPath,
    [Parameter(Mandatory)] [string] $InstallerPath,
    [Parameter(Mandatory)] [ValidatePattern('^[A-Fa-f0-9]{64}$')] [string] $ExpectedApplicationSha256,
    [Parameter(Mandatory)] [ValidatePattern('^[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}$')] [string] $ExpectedOuterRunEvidenceId
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

function Get-ZipEntrySha256 {
    param([string] $Path, [string] $EntryName)
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $entry = $archive.GetEntry($EntryName)
        Require-Condition ($null -ne $entry) "archive entry is missing: $EntryName"
        $stream = $entry.Open()
        $algorithm = [System.Security.Cryptography.SHA256]::Create()
        try { [Convert]::ToBase64String($algorithm.ComputeHash($stream)) }
        finally { $algorithm.Dispose(); $stream.Dispose() }
    }
    finally { $archive.Dispose() }
}

Require-Condition ($result.schemaVersion -eq 2) 'schemaVersion must be 2'
Require-Condition ($result.evidenceId -match '^[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}$') 'evidence identity is missing or malformed'
Require-Condition ($result.outerRunEvidenceId -eq $ExpectedOuterRunEvidenceId.ToLowerInvariant()) 'outer run evidence identity differs'
Require-Condition ($result.gate -eq 'real-excel-ten-minute-soak') 'gate identity is wrong'
Require-Condition ($result.status -eq 'Passed') 'status must be Passed'
$resultStartedUtc = [DateTime]::Parse([string]$result.startedUtc).ToUniversalTime()
$resultFinishedUtc = [DateTime]::Parse([string]$result.finishedUtc).ToUniversalTime()
Require-Condition ($resultFinishedUtc -gt $resultStartedUtc -and $result.durationSeconds -gt 0) 'result timestamps are invalid'
Require-Condition ($result.saveCount -eq 20) 'exactly 20 saves are required'
Require-Condition ($result.saveIntervalSeconds -ge 30) 'save interval must be at least 30 seconds'
Require-Condition ($result.monotonicDurationSeconds -ge 600) 'monotonic soak duration must be at least ten minutes'
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
$xlsmEvidencePath = Require-EvidenceFile 'fixtures/Soak Macro.xlsm'
Require-Condition ((Get-ZipEntrySha256 $xlsmEvidencePath 'xl/vbaProject.bin') -eq $macro.macroHashAfter) 'retained xlsm evidence does not match the reported VBA hash'

$saves = @($result.saves)
Require-Condition ($saves.Count -eq 20) 'exactly 20 save records are required'
Require-Condition (($saves[-1].monotonicStartSeconds - $saves[0].monotonicStartSeconds) -ge 600) 'first-to-last real save span is under ten monotonic minutes'

foreach ($format in @('xlsx', 'xlsm')) {
    $formatSaves = @($saves | Where-Object format -eq $format)
    Require-Condition ($formatSaves.Count -eq 10) "$format must have ten saves"
    $expectedSequences = @(1..10)
    Require-Condition (@(Compare-Object $expectedSequences @($formatSaves.sequence)).Count -eq 0) "$format sequence set is not exactly 1 through 10"
    Require-Condition (@($formatSaves.sha256 | Select-Object -Unique).Count -eq 10) "$format stable hashes are missing or duplicated"

    $baselinePath = Require-EvidenceFile "probe/$format-baseline.json"
    $baseline = Get-Content $baselinePath -Raw | ConvertFrom-Json
    Require-Condition ($baseline.passed -eq $true -and @($baseline.failures).Count -eq 0) "$format baseline probe failed"
    Require-Condition ($baseline.workbookStatus -eq 'Active' -and $baseline.errorCount -eq 0 -and [string]::IsNullOrWhiteSpace($baseline.lastError)) "$format baseline was not Active without errors"
    Require-Condition ($baseline.currentSequence -eq 0 -and $baseline.versionCount -eq 0 -and $baseline.distinctVersionHashCount -eq 0 -and $null -eq $baseline.latestVersion) "$format baseline was not silent sequence zero"
    Require-Condition ($baseline.currentHash -match '^[A-Fa-f0-9]{64}$') "$format baseline hash is missing or malformed"

    $settledPath = Require-EvidenceFile "probe/$format-settled.json"
    $settled = Get-Content $settledPath -Raw | ConvertFrom-Json
    $lastSave = $formatSaves[-1]
    Require-Condition ($settled.passed -eq $true -and @($settled.failures).Count -eq 0) "$format settled probe failed"
    Require-Condition ($settled.workbookStatus -eq 'Active' -and $settled.errorCount -eq 0 -and [string]::IsNullOrWhiteSpace($settled.lastError)) "$format did not settle Active without errors"
    Require-Condition ($settled.currentSequence -eq 10 -and $settled.versionCount -eq 10 -and $settled.distinctVersionHashCount -eq 10) "$format settled sequence or version counts differ"
    Require-Condition ($settled.currentHash -eq $lastSave.sha256 -and $settled.latestVersion.sha256 -eq $lastSave.sha256) "$format settled hashes differ from the final save"
    Require-Condition ($settled.latestVersion.sequence -eq 10 -and $settled.latestVersion.reportStatus -eq 'Ready' -and $settled.latestVersion.cellChangeCount -eq 1 -and $settled.latestVersion.sheetChangeCount -eq 0) "$format settled version or report differs"
    Require-Condition ($settled.cellChange.address -eq 'Y1001' -and $settled.cellChange.kinds.Split(',') -contains 'LiteralChanged') "$format settled delta is not the final literal change"
    $settledBefore = $settled.cellChange.beforeJson | ConvertFrom-Json
    $settledAfter = $settled.cellChange.afterJson | ConvertFrom-Json
    Require-Condition ($settledAfter.literalValue -eq $lastSave.value -and $settledBefore.literalValue -eq ('EDT-SOAK-{0:D2}' -f ($lastSave.index - 2))) "$format settled literal values differ from the final save"
    $fixtureName = if ($format -eq 'xlsx') { 'Soak.xlsx' } else { 'Soak Macro.xlsm' }
    $fixturePath = Require-EvidenceFile "fixtures/$fixtureName"
    $fixtureHash = (Get-FileHash -LiteralPath $fixturePath -Algorithm SHA256).Hash.ToUpperInvariant()
    Require-Condition ($fixtureHash -eq $lastSave.sha256) "$format retained final workbook hash differs from the final save"
}

for ($index = 0; $index -lt $saves.Count; $index++) {
    $save = $saves[$index]
    $wantFormat = if (($index % 2) -eq 0) { 'xlsx' } else { 'xlsm' }
    Require-Condition ($save.index -eq ($index + 1)) "save index $index is out of order"
    Require-Condition ($save.format -eq $wantFormat) "save $($index + 1) does not alternate workbook formats"
    Require-Condition ($save.value -eq ('EDT-SOAK-{0:D2}' -f ($index + 1))) "save $($index + 1) value is wrong"
    Require-Condition ($save.monotonicStartSeconds -ge (($index * $result.saveIntervalSeconds) - 0.25)) "save $($index + 1) started before its monotonic schedule"
    Require-Condition ($save.sha256 -match '^[A-F0-9]{64}$') "save $($index + 1) hash is malformed"
    Require-Condition ($save.captureMilliseconds -ge 0 -and $save.captureMilliseconds -le 20000) "save $($index + 1) has an invalid capture duration"
    $scheduledUtc = [DateTime]::Parse([string]$save.scheduledUtc).ToUniversalTime()
    $saveStartedUtc = [DateTime]::Parse([string]$save.saveStartedUtc).ToUniversalTime()
    $ctrlSaveUtc = [DateTime]::Parse([string]$save.ctrlSaveUtc).ToUniversalTime()
    $capturedUtc = [DateTime]::Parse([string]$save.capturedUtc).ToUniversalTime()
    Require-Condition ($scheduledUtc -ge $resultStartedUtc -and $saveStartedUtc -ge $scheduledUtc.AddSeconds(-1) -and $ctrlSaveUtc -ge $saveStartedUtc -and $capturedUtc -ge $ctrlSaveUtc -and $capturedUtc -le $resultFinishedUtc) "save $($index + 1) timestamps are outside the soak or out of order"
    if ($index -gt 0) {
        $previousCapturedUtc = [DateTime]::Parse([string]$saves[$index - 1].capturedUtc).ToUniversalTime()
        Require-Condition ($saveStartedUtc -ge $previousCapturedUtc) "save $($index + 1) overlaps the previous save"
    }
    $probePath = Require-EvidenceFile $save.probe
    $probe = Get-Content $probePath -Raw | ConvertFrom-Json
    Require-Condition $probe.passed "probe failed for save $($index + 1)"
    Require-Condition ($probe.workbookStatus -eq 'Active') "workbook was not Active after save $($index + 1)"
    Require-Condition ($probe.currentSequence -eq $save.sequence) "probe sequence differs for save $($index + 1)"
    Require-Condition ($probe.currentHash -eq $save.sha256) "probe hash differs for save $($index + 1)"
    Require-Condition ($probe.latestVersion.sha256 -eq $save.sha256) "latest-version hash differs for save $($index + 1)"
    Require-Condition ($probe.latestVersion.reportStatus -eq 'Ready') "report was not Ready for save $($index + 1)"
    Require-Condition ($probe.latestVersion.cellChangeCount -eq 1) "save $($index + 1) did not contain exactly one cell change"
    Require-Condition ($probe.latestVersion.sheetChangeCount -eq 0) "save $($index + 1) unexpectedly changed a sheet"
    Require-Condition ($probe.errorCount -eq 0) "capture errors exist after save $($index + 1)"
    Require-Condition ($probe.versionCount -eq $save.sequence) "version count differs after save $($index + 1)"
    Require-Condition ($probe.distinctVersionHashCount -eq $probe.versionCount) "duplicate version hash exists after save $($index + 1)"
    Require-Condition ($probe.cellChange.address -eq 'Y1001') "save $($index + 1) changed the wrong cell"
    Require-Condition ($probe.cellChange.kinds.Split(',') -contains $save.expectedKind) "save $($index + 1) has the wrong change kind"
    $after = $probe.cellChange.afterJson | ConvertFrom-Json
    Require-Condition ($after.literalValue -eq $save.value) "save $($index + 1) has the wrong new literal"
    if ($save.sequence -eq 1) {
        Require-Condition ($null -eq $probe.cellChange.beforeJson) "save $($index + 1) should add a previously empty cell"
    } else {
        $before = $probe.cellChange.beforeJson | ConvertFrom-Json
        Require-Condition ($before.literalValue -eq ('EDT-SOAK-{0:D2}' -f ($index - 1))) "save $($index + 1) has the wrong old literal"
    }
    $reportPath = Require-EvidenceFile $save.report
    $markdown = [System.IO.File]::ReadAllText($reportPath)
    Require-Condition ($markdown.IndexOf('Y1001', [StringComparison]::Ordinal) -ge 0) "portable report omits the address for save $($index + 1)"
    Require-Condition ($markdown.IndexOf($save.value, [StringComparison]::Ordinal) -ge 0) "portable report omits the value for save $($index + 1)"
}

$failedAssertions = @($result.assertions | Where-Object { -not $_.passed })
Require-Condition ($failedAssertions.Count -eq 0) 'result contains a failed assertion'
Require-Condition ([string]::IsNullOrWhiteSpace($result.failure)) 'result contains an unhandled failure'
$screenshotPath = Require-EvidenceFile 'screenshots/soak-history.png'
$uiaPath = Require-EvidenceFile 'uia/soak-history.json'
Require-Condition ((Get-Item $screenshotPath).Length -gt 0) 'soak history screenshot is empty'
$uiaText = [System.IO.File]::ReadAllText($uiaPath)
Require-Condition ($uiaText.IndexOf('Soak.xlsx', [StringComparison]::Ordinal) -ge 0) 'history UIA evidence omits the xlsx workbook'
Require-Condition ($uiaText.IndexOf('Soak Macro.xlsm', [StringComparison]::Ordinal) -ge 0) 'history UIA evidence omits the xlsm workbook'

Write-Output "REAL_EXCEL_SOAK_VALID|$path|$installerHash|$applicationHash"
