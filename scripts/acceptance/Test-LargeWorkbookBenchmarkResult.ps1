[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string] $ResultPath,
    [Parameter(Mandatory)] [string] $InstallerPath,
    [Parameter(Mandatory)] [ValidatePattern('^[A-Fa-f0-9]{64}$')] [string] $ExpectedApplicationSha256
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'UiAutomation.psm1') -Force

$path = (Resolve-Path $ResultPath).Path
$root = Split-Path -Parent $path
$installer = (Resolve-Path $InstallerPath).Path
$installerHash = Get-AcceptanceFileSha256 -Path $installer
$expectedApplicationHash = $ExpectedApplicationSha256.ToUpperInvariant()
$result = Get-Content $path -Raw | ConvertFrom-Json

function Require-Condition {
    param([Parameter(Mandatory)] [bool] $Condition, [Parameter(Mandatory)] [string] $Message)
    if (-not $Condition) { throw "Invalid large-workbook benchmark evidence: $Message" }
}

function Resolve-EvidenceFile {
    param([Parameter(Mandatory)] [string] $RelativePath)
    Require-Condition -Condition (-not [string]::IsNullOrWhiteSpace($RelativePath)) -Message 'an evidence path is empty'
    Require-Condition -Condition (-not [System.IO.Path]::IsPathRooted($RelativePath)) -Message "evidence path must be relative: $RelativePath"
    $fullPath = [System.IO.Path]::GetFullPath((Join-Path $root $RelativePath))
    $null = Get-AcceptanceRelativePath -BasePath $root -Path $fullPath
    Require-Condition -Condition (Test-Path $fullPath -PathType Leaf) -Message "evidence file is missing: $RelativePath"
    $fullPath
}

function Require-EvidenceFile {
    param([Parameter(Mandatory)] [string] $RelativePath)
    $null = Resolve-EvidenceFile $RelativePath
}

function Get-UiaNames {
    param([Parameter(Mandatory)] [object] $Node)
    if (-not [string]::IsNullOrWhiteSpace([string]$Node.name)) { [string]$Node.name }
    foreach ($child in @($Node.children)) { Get-UiaNames $child }
}

Require-Condition ($result.schemaVersion -eq 1) 'schemaVersion must be 1'
Require-Condition ($result.gate -eq 'large-workbook-500k') 'gate must be large-workbook-500k'
Require-Condition ($result.status -eq 'Passed') 'status must be Passed'
Require-Condition ($result.candidate.installerSha256.ToUpperInvariant() -eq $installerHash) 'installer SHA-256 does not match the exact candidate'
Require-Condition ($result.candidate.expectedInstallerSha256.ToUpperInvariant() -eq $installerHash) 'frozen installer SHA-256 does not match the exact candidate'
Require-Condition ($result.candidate.applicationSha256.ToUpperInvariant() -eq $expectedApplicationHash) 'installed executable SHA-256 does not match the frozen candidate'
Require-Condition ($result.candidate.expectedApplicationSha256.ToUpperInvariant() -eq $expectedApplicationHash) 'frozen executable SHA-256 differs from the required hash'

Require-Condition ($result.fixture.rows -eq 25000) 'fixture rows must equal 25,000'
Require-Condition ($result.fixture.columns -eq 20) 'fixture columns must equal 20'
Require-Condition ($result.fixture.populatedCells -eq 500000) 'fixture populatedCells must equal 500,000'
Require-Condition ($result.fixture.usedRangeRows -eq 25000) 'visible Excel used range must have 25,000 rows'
Require-Condition ($result.fixture.usedRangeColumns -eq 20) 'visible Excel used range must have 20 columns'
Require-Condition ($result.fixture.usedRangeCells -eq 500000) 'visible Excel used range must contain 500,000 cells'
Require-Condition ($result.fixture.initialSha256 -match '^[A-F0-9]{64}$') 'fixture SHA-256 is missing or malformed'
Require-Condition ($result.fixture.bytes -gt 0) 'fixture byte count must be positive'
$fixturePath = Resolve-EvidenceFile $result.fixture.path

Require-Condition ($result.thresholds.phaseTimeoutSeconds -le 60) 'phase timeout is weaker than 60 seconds'
Require-Condition ($result.thresholds.maxUiaHeartbeatMilliseconds -le 2000) 'UI heartbeat threshold is weaker than two seconds'
Require-Condition ($result.thresholds.maxHeartbeatGapMilliseconds -le 2000) 'heartbeat gap threshold is weaker than two seconds'
Require-Condition ($result.thresholds.maxPrivateBytes -le 1610612736L) 'memory threshold is weaker than 1.5 GiB'
Require-Condition ($result.thresholds.requiredBulkChangedCells -eq 10000) 'bulk-change requirement must equal 10,000'
Require-Condition ($result.thresholds.automaticReportCellLimit -eq 5000) 'automatic report limit must equal 5,000'

$telemetryPath = Resolve-EvidenceFile $result.telemetry.path
Require-EvidenceFile $result.telemetry.monitorLog
$telemetry = @(Get-Content $telemetryPath | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_ | ConvertFrom-Json })
Require-Condition ($telemetry.Count -eq $result.telemetry.sampleCount) 'telemetry sample count differs from the raw JSONL'
Require-Condition ($telemetry.Count -ge 2) 'raw telemetry must contain at least two samples'
$monitorStartedUtc = [DateTime]::Parse($result.telemetry.monitorStartedUtc).ToUniversalTime()
$monitorStoppedUtc = [DateTime]::Parse($result.telemetry.monitorStoppedUtc).ToUniversalTime()
Require-Condition ($monitorStoppedUtc -gt $monitorStartedUtc) 'telemetry monitor timestamps are invalid'
$rawRequiredDurationMilliseconds = ($monitorStoppedUtc - $monitorStartedUtc).TotalMilliseconds
$rawInitialLagMilliseconds = [math]::Max(0, ([DateTime]::Parse($telemetry[0].startedUtc).ToUniversalTime() - $monitorStartedUtc).TotalMilliseconds)
$rawTailLagMilliseconds = ($monitorStoppedUtc - [DateTime]::Parse($telemetry[-1].startedUtc).ToUniversalTime()).TotalMilliseconds
$rawExpectedMinimumSamples = [math]::Max(2, [math]::Floor($rawRequiredDurationMilliseconds / $result.thresholds.maxHeartbeatGapMilliseconds))
Require-Condition ($result.telemetry.expectedMinimumSamples -eq $rawExpectedMinimumSamples) 'reported minimum sample count differs from monitor timestamps'
Require-Condition ([math]::Round($result.telemetry.requiredDurationMilliseconds, 3) -eq [math]::Round($rawRequiredDurationMilliseconds, 3)) 'reported monitor duration differs from monitor timestamps'
Require-Condition ([math]::Round($result.telemetry.initialLagMilliseconds, 3) -eq [math]::Round($rawInitialLagMilliseconds, 3)) 'reported initial telemetry lag differs from raw evidence'
Require-Condition ([math]::Round($result.telemetry.tailLagMilliseconds, 3) -eq [math]::Round($rawTailLagMilliseconds, 3)) 'reported tail telemetry lag differs from raw evidence'
Require-Condition ($result.telemetry.sampleCount -ge $result.telemetry.expectedMinimumSamples -and $result.telemetry.sampleCount -ge 2) 'heartbeat sample coverage is insufficient'
Require-Condition ($result.telemetry.initialLagMilliseconds -le $result.thresholds.maxHeartbeatGapMilliseconds) 'heartbeat monitor started too late'
Require-Condition ($result.telemetry.tailLagMilliseconds -le $result.thresholds.maxHeartbeatGapMilliseconds) 'heartbeat monitor stopped before the benchmark ended'
Require-Condition ($result.telemetry.failedSamples -eq 0) 'telemetry contains a failed responsiveness or memory sample'
Require-Condition ($result.telemetry.maxUiaDurationMilliseconds -le $result.thresholds.maxUiaHeartbeatMilliseconds) 'a UI Automation query exceeded two seconds'
Require-Condition ($result.telemetry.maxGapMilliseconds -le $result.thresholds.maxHeartbeatGapMilliseconds) 'UI heartbeat coverage has a gap over two seconds'
Require-Condition ($result.telemetry.peakPrivateBytes -lt $result.thresholds.maxPrivateBytes) 'peak private bytes reached the 1.5 GiB limit'
Require-Condition ($null -ne $result.telemetry.peakPrivateWorkingSetBytes -and $result.telemetry.peakPrivateWorkingSetBytes -lt $result.thresholds.maxPrivateBytes) 'peak private working set reached the 1.5 GiB limit or was not measured'
Require-Condition ($result.telemetry.peakWorkingSetBytes -lt $result.thresholds.maxPrivateBytes) 'peak total working set reached the 1.5 GiB limit'
Require-Condition (-not $result.telemetry.monitorForcedStop) 'telemetry monitor did not stop cleanly'
Require-Condition ($result.telemetry.monitorJobState -eq 'Completed') 'telemetry monitor job did not complete successfully'
$rawTelemetryFailures = @($telemetry | Where-Object {
    -not $_.processFound -or -not $_.responding -or -not $_.uiaOk -or
    $null -eq $_.privateWorkingSetBytes -or
    [double]$_.privateWorkingSetAgeMilliseconds -gt $result.thresholds.maxHeartbeatGapMilliseconds -or
    [double]$_.durationMilliseconds -gt $result.thresholds.maxUiaHeartbeatMilliseconds -or
    [double]$_.gapMilliseconds -gt $result.thresholds.maxHeartbeatGapMilliseconds
})
Require-Condition ($rawTelemetryFailures.Count -eq 0) 'raw telemetry contains an unresponsive, missing, stale-memory, slow-query, or heartbeat-gap sample'
$rawMaxUiaDuration = ($telemetry | Measure-Object durationMilliseconds -Maximum).Maximum
$rawMaxGap = ($telemetry | Select-Object -Skip 1 | Measure-Object gapMilliseconds -Maximum).Maximum
$rawPeakPrivateBytes = ($telemetry | Measure-Object privateBytes -Maximum).Maximum
$rawPeakPrivateWorkingSetBytes = ($telemetry | Measure-Object privateWorkingSetBytes -Maximum).Maximum
$rawPeakWorkingSetBytes = ($telemetry | Measure-Object peakWorkingSetBytes -Maximum).Maximum
Require-Condition ($rawMaxUiaDuration -eq $result.telemetry.maxUiaDurationMilliseconds) 'reported maximum UIA duration differs from raw telemetry'
Require-Condition ($rawMaxGap -eq $result.telemetry.maxGapMilliseconds) 'reported maximum heartbeat gap differs from raw telemetry'
Require-Condition ($rawPeakPrivateBytes -eq $result.telemetry.peakPrivateBytes) 'reported peak private bytes differ from raw telemetry'
Require-Condition ($rawPeakPrivateWorkingSetBytes -eq $result.telemetry.peakPrivateWorkingSetBytes) 'reported private working-set peak differs from raw telemetry'
Require-Condition ($rawPeakWorkingSetBytes -eq $result.telemetry.peakWorkingSetBytes) 'reported total working-set peak differs from raw telemetry'

$phases = @($result.phases)
Require-Condition ($phases.Count -eq 5) 'exactly five phases are required'
$expected = @(
    @{ Name = 'baseline'; Sequence = 0L; Address = $null; CellChanges = 0L; Headings = $null; Truncated = $null },
    @{ Name = 'beginning-cell save'; Sequence = 1L; Address = 'A1'; CellChanges = 1L; Headings = 1L; Truncated = $false },
    @{ Name = 'middle-cell save'; Sequence = 2L; Address = 'J12500'; CellChanges = 1L; Headings = 1L; Truncated = $false },
    @{ Name = 'end-cell save'; Sequence = 3L; Address = 'T25000'; CellChanges = 1L; Headings = 1L; Truncated = $false },
    @{ Name = '10,000-cell truncation save'; Sequence = 4L; Address = 'A2'; CellChanges = 10000L; Headings = 5000L; Truncated = $true }
)
for ($index = 0; $index -lt $expected.Count; $index++) {
    $actual = $phases[$index]
    $want = $expected[$index]
    Require-Condition ($actual.name -eq $want.Name) "phase $index name must be $($want.Name)"
    Require-Condition ($actual.sequence -eq $want.Sequence) "phase $($want.Name) sequence must be $($want.Sequence)"
    Require-Condition ($actual.expectedCellChanges -eq $want.CellChanges) "phase $($want.Name) cell-change count is wrong"
    Require-Condition ($actual.elapsedMilliseconds -le 60000) "phase $($want.Name) exceeded 60 seconds"
    Require-Condition ($actual.workbookSha256 -match '^[A-F0-9]{64}$') "phase $($want.Name) workbook SHA-256 is malformed"
    $probePath = Resolve-EvidenceFile $actual.probe
    $probeResult = Get-Content $probePath -Raw | ConvertFrom-Json
    Require-Condition ($probeResult.passed) "phase $($want.Name) external database probe did not pass"
    Require-Condition ($probeResult.workbookStatus -eq 'Active') "phase $($want.Name) workbook was not Active"
    Require-Condition ($probeResult.currentSequence -eq $want.Sequence) "phase $($want.Name) probe sequence differs"
    Require-Condition ($probeResult.currentHash -eq $actual.workbookSha256) "phase $($want.Name) current hash differs from the phase record"
    Require-Condition ($probeResult.errorCount -eq 0) "phase $($want.Name) has capture errors"
    Require-Condition ([string]::IsNullOrWhiteSpace($probeResult.lastError)) "phase $($want.Name) has a current capture error"
    Require-EvidenceFile $actual.uiEvidence.screenshot
    $uiaPath = Resolve-EvidenceFile $actual.uiEvidence.uiaTree
    $uiaNames = @(Get-UiaNames (Get-Content $uiaPath -Raw | ConvertFrom-Json))
    Require-Condition ($uiaNames -contains $fixturePath) "phase $($want.Name) UI evidence does not identify the benchmark workbook"
    if ($want.Sequence -gt 0) {
        Require-Condition ($actual.evidenceAddress -eq $want.Address) "phase $($want.Name) evidence address is wrong"
        Require-Condition ($actual.expectedKind -eq 'FormulaRemoved') "phase $($want.Name) expected kind is wrong"
        Require-Condition ($probeResult.latestVersion.sequence -eq $want.Sequence) "phase $($want.Name) latest version sequence differs"
        Require-Condition ($probeResult.latestVersion.sha256 -eq $actual.workbookSha256) "phase $($want.Name) version hash differs"
        Require-Condition ($probeResult.latestVersion.reportStatus -eq 'Ready') "phase $($want.Name) report was not Ready"
        Require-Condition ($probeResult.latestVersion.cellChangeCount -eq $want.CellChanges) "phase $($want.Name) raw SQLite cell-change count is wrong"
        Require-Condition ($probeResult.latestVersion.sheetChangeCount -eq 0) "phase $($want.Name) raw SQLite sheet-change count is not zero"
        Require-Condition ($probeResult.cellChange.sheetName -eq 'Performance' -and $probeResult.cellChange.address -eq $want.Address) "phase $($want.Name) raw SQLite delta address is wrong"
        Require-Condition (@(([string]$probeResult.cellChange.kinds).Split(',')) -contains 'FormulaRemoved') "phase $($want.Name) raw SQLite delta lacks FormulaRemoved"
        $after = $probeResult.cellChange.afterJson | ConvertFrom-Json
        Require-Condition ($after.literalValue -eq $actual.expectedValue) "phase $($want.Name) raw SQLite new literal value is wrong"
        Require-Condition ($actual.automaticReport.addressOccurrences -eq 1) "phase $($want.Name) address is not represented exactly once"
        Require-Condition ($actual.automaticReport.cellHeadingCount -eq $want.Headings) "phase $($want.Name) automatic report detail count is wrong"
        Require-Condition ($actual.automaticReport.valueOccurrences -eq $want.Headings) "phase $($want.Name) included values are missing or duplicated"
        Require-Condition ($actual.automaticReport.truncated -eq $want.Truncated) "phase $($want.Name) truncation state is wrong"
        $automaticReportPath = Resolve-EvidenceFile $actual.automaticReport.path
        $automaticMarkdown = [System.IO.File]::ReadAllText($automaticReportPath)
        $addressPattern = '(?m)^' + [regex]::Escape("### Performance!$($want.Address)") + '\r?$'
        $rawAddressCount = [regex]::Matches($automaticMarkdown, $addressPattern).Count
        $rawHeadingCount = [regex]::Matches($automaticMarkdown, '(?m)^### Performance!').Count
        $rawValueCount = [regex]::Matches($automaticMarkdown, [regex]::Escape($actual.expectedValue)).Count
        $rawTruncated = $automaticMarkdown.IndexOf('> This automatic report includes the first ', [StringComparison]::Ordinal) -ge 0
        Require-Condition ($rawAddressCount -eq 1) "phase $($want.Name) automatic Markdown address is missing or duplicated"
        Require-Condition ($rawHeadingCount -eq $want.Headings) "phase $($want.Name) automatic Markdown detail count differs"
        Require-Condition ($rawValueCount -eq $want.Headings) "phase $($want.Name) automatic Markdown values are missing or duplicated"
        Require-Condition ($rawTruncated -eq $want.Truncated) "phase $($want.Name) automatic Markdown truncation marker differs"
        $expectedSummary = "$($want.CellChanges.ToString('N0')) cells, 0 sheet changes"
        Require-Condition ($uiaNames -contains $expectedSummary) "phase $($want.Name) UI summary is missing"
        Require-Condition (@($uiaNames | Where-Object { $_ -match "^Version $($want.Sequence)(?:\s|$)" }).Count -gt 0) "phase $($want.Name) UI version is missing"
    }
    else {
        Require-Condition ($null -eq $probeResult.latestVersion) 'baseline probe unexpectedly contains a version record'
    }
}

$hashes = @($phases | Select-Object -ExpandProperty workbookSha256 -Unique)
Require-Condition ($hashes.Count -eq 5) 'baseline and four saves must have five distinct stable hashes'
Require-Condition ((Get-AcceptanceFileSha256 -Path $fixturePath) -eq $phases[4].workbookSha256) 'final fixture bytes do not match sequence 4'
$bulk = $phases[4]
Require-Condition ($bulk.fullExport.exportedViaInstalledUi) 'full Markdown export was not driven through the installed UI'
Require-Condition ($bulk.fullExport.cellHeadingCount -eq 10000) 'full export does not contain all 10,000 cell details'
Require-Condition ($bulk.fullExport.valueOccurrences -eq 10000) 'full export does not contain all 10,000 new values'
Require-Condition (-not $bulk.fullExport.truncated) 'full export is truncated'
Require-Condition ($bulk.fullExport.elapsedMilliseconds -le 60000) 'full export exceeded 60 seconds'
Require-Condition ($bulk.fullExport.defaultFileName -match 'version-000004-full\.md$') 'full export dialog did not target sequence 4'
Require-EvidenceFile $bulk.fullExport.uiEvidence.screenshot
$exportUiaPath = Resolve-EvidenceFile $bulk.fullExport.uiEvidence.uiaTree
$exportUia = Get-Content $exportUiaPath -Raw | ConvertFrom-Json
Require-Condition ($exportUia.name -eq 'Export complete Markdown report') 'full-export UI evidence is not the installed Save dialog'
$fullExportPath = Resolve-EvidenceFile $bulk.fullExport.path
$fullMarkdown = [System.IO.File]::ReadAllText($fullExportPath)
Require-Condition ([regex]::Matches($fullMarkdown, '(?m)^### Performance!').Count -eq 10000) 'raw full export does not contain exactly 10,000 cell details'
Require-Condition ([regex]::Matches($fullMarkdown, 'EDT-LARGE-BULK').Count -eq 10000) 'raw full export does not contain exactly 10,000 new values'
Require-Condition ($fullMarkdown.IndexOf('> This automatic report includes the first ', [StringComparison]::Ordinal) -lt 0) 'raw full export contains a truncation marker'

$failedAssertions = @($result.assertions | Where-Object { -not $_.passed })
Require-Condition ($failedAssertions.Count -eq 0) 'result contains a failed assertion'
Require-Condition (@($result.assertions).Count -ge 40) 'result does not contain the complete assertion set'
Require-Condition ([string]::IsNullOrWhiteSpace($result.failure)) 'result contains an unhandled failure'

Write-Output "LARGE_WORKBOOK_BENCHMARK_VALID|$path|$installerHash|$expectedApplicationHash"
