[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$acceptanceRoot = Split-Path -Parent $PSScriptRoot
$repositoryRoot = Split-Path -Parent (Split-Path -Parent $acceptanceRoot)
$testRoot = Join-Path $repositoryRoot ('artifacts\validator-regression\' + [Guid]::NewGuid().ToString('D'))
$null = New-Item -ItemType Directory -Path $testRoot
$applicationHash = 'A' * 64
$outerRunEvidenceId = '12345678-1234-1234-1234-123456789012'
$testStartedUtc = [DateTime]::Parse('2026-09-01T00:00:00Z').ToUniversalTime()
$testCount = 0
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.IO.Compression

# These are synthetic validator inputs, not product acceptance evidence. No installed
# process, UI, registry, user history, or real workbook is read or modified.
function Write-TestText {
    param([string] $Path, [AllowEmptyString()] [string] $Text)
    $null = New-Item -ItemType Directory -Path (Split-Path -Parent $Path) -Force
    [System.IO.File]::WriteAllText($Path, $Text, [System.Text.UTF8Encoding]::new($false))
}

function Write-TestJson {
    param([string] $Path, [object] $Value)
    Write-TestText $Path ($Value | ConvertTo-Json -Depth 20)
}

function Get-TestHash {
    param([string] $Path)
    (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToUpperInvariant()
}

function Add-TestZipEntry {
    param([string] $Path, [string] $EntryName, [string] $Content)
    $archive = [System.IO.Compression.ZipFile]::Open($Path, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
        $stream = $archive.CreateEntry($EntryName).Open()
        try {
            $bytes = [System.Text.Encoding]::UTF8.GetBytes($Content)
            $stream.Write($bytes, 0, $bytes.Length)
        } finally { $stream.Dispose() }
    } finally { $archive.Dispose() }
}

function New-TestCandidate {
    param([string] $Directory)
    $installer = Join-Path $Directory 'installer.bin'
    Write-TestText $installer 'synthetic installer identity; never executable'
    $hash = Get-TestHash $installer
    [pscustomobject]@{
        installer = $installer
        identity = [pscustomobject]@{
            installerSha256 = $hash; expectedInstallerSha256 = $hash
            applicationSha256 = $applicationHash; expectedApplicationSha256 = $applicationHash
        }
    }
}

function New-TestProbe {
    param([int] $Sequence, [string] $Hash, [string] $Address, [string] $Kind, [string] $Value, [string] $BeforeValue)
    $latest = $null
    $cell = $null
    if ($Sequence -gt 0) {
        $latest = [pscustomobject]@{ sequence = $Sequence; sha256 = $Hash; reportStatus = 'Ready'; cellChangeCount = 1; sheetChangeCount = 0 }
        $before = if ($Sequence -eq 1) { $null } else { @{ literalValue = $BeforeValue } | ConvertTo-Json -Compress }
        $cell = [pscustomobject]@{ sheetName = 'Performance'; address = $Address; kinds = $Kind; beforeJson = $before; afterJson = (@{ literalValue = $Value } | ConvertTo-Json -Compress) }
    }
    [pscustomobject]@{
        passed = $true; failures = @(); workbookStatus = 'Active'; currentSequence = $Sequence
        currentHash = $Hash; lastError = $null; errorCount = 0; versionCount = $Sequence
        distinctVersionHashCount = $Sequence; latestVersion = $latest; cellChange = $cell
    }
}

function New-BenchmarkBundle {
    param([string] $Directory)
    $candidate = New-TestCandidate $Directory
    $fixtureRelative = 'fixtures/Acceptance Large 500k.xlsx'
    $fixture = Join-Path $Directory $fixtureRelative
    Write-TestText $fixture 'synthetic final benchmark fixture'
    $finalHash = Get-TestHash $fixture
    $telemetry = @(for ($index = 0; $index -lt 40; $index++) {
        [pscustomobject]@{
            sample = $index; startedUtc = $testStartedUtc.AddSeconds(1).AddMilliseconds($index * 250).ToString('O')
            gapMilliseconds = $(if ($index -eq 0) { 0 } else { 250 }); durationMilliseconds = 10
            processFound = $true; responding = $true; uiaOk = $true
            privateBytes = 1000000; privateWorkingSetBytes = 800000; privateWorkingSetAgeMilliseconds = 10; peakWorkingSetBytes = 1200000
        }
    })
    Write-TestText (Join-Path $Directory 'logs/monitor.txt') ''
    Write-TestText (Join-Path $Directory 'screenshots/state.png') 'synthetic screenshot sentinel'
    $states = @(
        @{ Name = 'baseline'; Address = ''; Value = ''; Count = 0; Headings = 0 },
        @{ Name = 'beginning-cell save'; Address = 'A1'; Value = 'EDT-LARGE-BEGIN'; Count = 1; Headings = 1 },
        @{ Name = 'middle-cell save'; Address = 'J12500'; Value = 'EDT-LARGE-MIDDLE'; Count = 1; Headings = 1 },
        @{ Name = 'end-cell save'; Address = 'T25000'; Value = 'EDT-LARGE-END'; Count = 1; Headings = 1 },
        @{ Name = '10,000-cell truncation save'; Address = 'A2'; Value = 'EDT-LARGE-BULK'; Count = 10000; Headings = 5000 }
    )
    $phases = @(for ($sequence = 0; $sequence -lt $states.Count; $sequence++) {
        $state = $states[$sequence]
        $hash = if ($sequence -eq 4) { $finalHash } else { ($sequence + 1).ToString('X64') }
        $probe = New-TestProbe $sequence $hash $state.Address 'FormulaRemoved' $state.Value ''
        if ($sequence -gt 0) { $probe.latestVersion.cellChangeCount = $state.Count }
        Write-TestJson (Join-Path $Directory "probe/$sequence.json") $probe
        $names = @($fixture, "Version $sequence", "$($state.Count.ToString('N0')) cells, 0 sheet changes")
        Write-TestJson (Join-Path $Directory "uia/$sequence.json") ([pscustomobject]@{ name = 'Excel Diff Tracker'; children = @($names | ForEach-Object { [pscustomobject]@{ name = $_; children = @() } }) })
        $phase = [pscustomobject]@{
            name = $state.Name; sequence = $sequence; expectedCellChanges = $state.Count; elapsedMilliseconds = 1000
            workbookSha256 = $hash; probe = "probe/$sequence.json"
            uiEvidence = [pscustomobject]@{ screenshot = 'screenshots/state.png'; uiaTree = "uia/$sequence.json" }
        }
        if ($sequence -gt 0) {
            $lines = @("### Performance!$($state.Address)", $state.Value)
            if ($state.Headings -gt 1) { $lines += @(for ($i = 1; $i -lt $state.Headings; $i++) { "### Performance!B$($i + 1)"; $state.Value }) }
            if ($sequence -eq 4) { $lines += '> This automatic report includes the first 5,000 of 10,000 cell changes.' }
            Write-TestText (Join-Path $Directory "reports/$sequence.md") ($lines -join "`n")
            $phase | Add-Member evidenceAddress $state.Address
            $phase | Add-Member expectedKind 'FormulaRemoved'
            $phase | Add-Member expectedValue $state.Value
            $phase | Add-Member automaticReport ([pscustomobject]@{ path = "reports/$sequence.md"; addressOccurrences = 1; cellHeadingCount = $state.Headings; valueOccurrences = $state.Headings; truncated = ($sequence -eq 4) })
        }
        $phase
    })
    $fullLines = @(for ($i = 1; $i -le 10000; $i++) { "### Performance!B$i"; 'EDT-LARGE-BULK' })
    Write-TestText (Join-Path $Directory 'reports/full.md') ($fullLines -join "`n")
    Write-TestJson (Join-Path $Directory 'uia/export.json') @{ name = 'Export complete Markdown report' }
    $phases[4] | Add-Member fullExport ([pscustomobject]@{
        path = 'reports/full.md'; exportedViaInstalledUi = $true; cellHeadingCount = 10000; valueOccurrences = 10000
        truncated = $false; elapsedMilliseconds = 1000; defaultFileName = 'version-000004-full.md'
        uiEvidence = [pscustomobject]@{ screenshot = 'screenshots/state.png'; uiaTree = 'uia/export.json' }
    })
    $result = [pscustomobject]@{
        schemaVersion = 2; evidenceId = [Guid]::NewGuid().ToString('D'); outerRunEvidenceId = $outerRunEvidenceId
        gate = 'large-workbook-500k'; status = 'Passed'; startedUtc = $testStartedUtc.ToString('O'); finishedUtc = $testStartedUtc.AddSeconds(20).ToString('O')
        candidate = $candidate.identity
        fixture = [pscustomobject]@{ rows = 25000; columns = 20; populatedCells = 500000; usedRangeRows = 25000; usedRangeColumns = 20; usedRangeCells = 500000; initialSha256 = (1).ToString('X64'); bytes = 1; path = $fixtureRelative }
        thresholds = [pscustomobject]@{ phaseTimeoutSeconds = 60; maxUiaHeartbeatMilliseconds = 2000; maxHeartbeatGapMilliseconds = 2000; maxPrivateBytes = 1610612736; requiredBulkChangedCells = 10000; automaticReportCellLimit = 5000 }
        telemetry = [pscustomobject]@{
            path = 'telemetry/samples.jsonl'; monitorLog = 'logs/monitor.txt'; sampleCount = 40
            monitorStartedUtc = $testStartedUtc.AddSeconds(1).ToString('O'); monitorStoppedUtc = $testStartedUtc.AddSeconds(11).ToString('O')
            expectedMinimumSamples = 5; requiredDurationMilliseconds = 10000; initialLagMilliseconds = 0; tailLagMilliseconds = 250
            failedSamples = 0; maxUiaDurationMilliseconds = 10; maxGapMilliseconds = 250; peakPrivateBytes = 1000000
            peakPrivateWorkingSetBytes = 800000; peakWorkingSetBytes = 1200000; monitorForcedStop = $false; monitorJobState = 'Completed'
        }
        phases = $phases; assertions = @(1..40 | ForEach-Object { [pscustomobject]@{ name = "synthetic check $_"; passed = $true } }); failure = $null
    }
    [pscustomobject]@{ directory = $Directory; installer = $candidate.installer; result = $result; resultPath = (Join-Path $Directory 'large-workbook-benchmark.json'); telemetry = $telemetry }
}

function New-SoakBundle {
    param([string] $Directory)
    $candidate = New-TestCandidate $Directory
    $fixtures = Join-Path $Directory 'fixtures'
    $null = New-Item -ItemType Directory -Path $fixtures
    Write-TestText (Join-Path $fixtures 'Soak.xlsx') 'synthetic final xlsx fixture'
    $xlsmPath = Join-Path $fixtures 'Soak Macro.xlsm'
    Add-TestZipEntry $xlsmPath 'xl/vbaProject.bin' 'synthetic VBA bytes; never executable'
    $algorithm = [System.Security.Cryptography.SHA256]::Create()
    try { $macroHash = [Convert]::ToBase64String($algorithm.ComputeHash([System.Text.Encoding]::UTF8.GetBytes('synthetic VBA bytes; never executable'))) }
    finally { $algorithm.Dispose() }
    $saves = @(for ($index = 1; $index -le 20; $index++) {
        $format = if ($index % 2 -eq 1) { 'xlsx' } else { 'xlsm' }
        $sequence = [int][math]::Ceiling($index / 2.0)
        $fixtureName = if ($format -eq 'xlsx') { 'Soak.xlsx' } else { 'Soak Macro.xlsm' }
        $hash = if ($sequence -eq 10) { Get-TestHash (Join-Path $fixtures $fixtureName) } else { $index.ToString('X64') }
        $value = 'EDT-SOAK-{0:D2}' -f $index
        $previous = if ($sequence -eq 1) { '' } else { 'EDT-SOAK-{0:D2}' -f ($index - 2) }
        $kind = if ($sequence -eq 1) { 'LiteralAdded' } else { 'LiteralChanged' }
        $probe = New-TestProbe $sequence $hash 'Y1001' $kind $value $previous
        $probeRelative = "probe/$format-sequence-$sequence.json"
        Write-TestJson (Join-Path $Directory $probeRelative) $probe
        if ($sequence -eq 10) { Write-TestJson (Join-Path $Directory "probe/$format-settled.json") $probe }
        $reportRelative = "reports/$index.md"
        Write-TestText (Join-Path $Directory $reportRelative) "Y1001 $value"
        $saveTime = $testStartedUtc.AddSeconds(($index - 1) * 32)
        [pscustomobject]@{
            index = $index; format = $format; sequence = $sequence; value = $value; expectedKind = $kind
            monotonicStartSeconds = (($index - 1) * 32); sha256 = $hash; captureMilliseconds = 1000
            scheduledUtc = $saveTime.ToString('O'); saveStartedUtc = $saveTime.ToString('O'); ctrlSaveUtc = $saveTime.AddMilliseconds(100).ToString('O'); capturedUtc = $saveTime.AddSeconds(1).ToString('O')
            probe = $probeRelative; report = $reportRelative
        }
    })
    foreach ($format in @('xlsx', 'xlsm')) {
        Write-TestJson (Join-Path $Directory "probe/$format-baseline.json") (New-TestProbe 0 ('B' * 64) '' '' '' '')
    }
    Write-TestText (Join-Path $Directory 'screenshots/soak-history.png') 'synthetic screenshot sentinel'
    Write-TestJson (Join-Path $Directory 'uia/soak-history.json') @{ name = 'Soak.xlsx Soak Macro.xlsm' }
    $result = [pscustomobject]@{
        schemaVersion = 2; evidenceId = [Guid]::NewGuid().ToString('D'); outerRunEvidenceId = $outerRunEvidenceId
        gate = 'real-excel-ten-minute-soak'; status = 'Passed'; startedUtc = $testStartedUtc.ToString('O'); finishedUtc = $testStartedUtc.AddSeconds(625).ToString('O')
        durationSeconds = 625; monotonicDurationSeconds = 625; saveCount = 20; saveIntervalSeconds = 32; candidate = $candidate.identity
        workbooks = @(
            [pscustomobject]@{ format = 'xlsx'; finalSequence = 10; macroHashBefore = $null; macroHashAfter = $null },
            [pscustomobject]@{ format = 'xlsm'; finalSequence = 10; macroHashBefore = $macroHash; macroHashAfter = $macroHash }
        )
        saves = $saves; assertions = @(); failure = $null
    }
    [pscustomobject]@{ directory = $Directory; installer = $candidate.installer; result = $result; resultPath = (Join-Path $Directory 'real-excel-soak.json') }
}

function Invoke-ValidatorCase {
    param([string] $Name, [ValidateSet('Benchmark', 'Soak')] [string] $Kind, [scriptblock] $Mutation, [string] $ExpectedFailure = '')
    $directory = Join-Path $testRoot $Name
    $bundle = if ($Kind -eq 'Benchmark') { New-BenchmarkBundle $directory } else { New-SoakBundle $directory }
    if ($null -ne $Mutation) { & $Mutation $bundle }
    Write-TestJson $bundle.resultPath $bundle.result
    if ($Kind -eq 'Benchmark') {
        Write-TestText (Join-Path $directory $bundle.result.telemetry.path) ((@($bundle.telemetry | ForEach-Object { $_ | ConvertTo-Json -Compress }) -join "`n") + "`n")
    }
    $validatorName = if ($Kind -eq 'Benchmark') { 'Test-LargeWorkbookBenchmarkResult.ps1' } else { 'Test-RealExcelSoakResult.ps1' }
    $passMarker = if ($Kind -eq 'Benchmark') { 'LARGE_WORKBOOK_BENCHMARK_VALID|' } else { 'REAL_EXCEL_SOAK_VALID|' }
    $validationFailure = $null
    $output = @()
    try {
        $output = @(& (Join-Path $acceptanceRoot $validatorName) -ResultPath $bundle.resultPath -InstallerPath $bundle.installer -ExpectedApplicationSha256 $applicationHash -ExpectedOuterRunEvidenceId $outerRunEvidenceId)
    } catch { $validationFailure = $_.Exception.Message }
    if ([string]::IsNullOrEmpty($ExpectedFailure)) {
        if ($null -ne $validationFailure) { throw "$Name should pass: $validationFailure" }
        if ($output.Count -ne 1 -or -not ([string]$output[0]).StartsWith($passMarker, [StringComparison]::Ordinal)) { throw "$Name did not emit the exact validator pass marker: $output" }
    } elseif ($null -eq $validationFailure -or $validationFailure.IndexOf($ExpectedFailure, [StringComparison]::Ordinal) -lt 0) {
        throw "$Name did not fail for '$ExpectedFailure': $validationFailure"
    }
    $script:testCount++
    Write-Output "VALIDATOR_REGRESSION_PASS|$Name"
}

Invoke-ValidatorCase 'benchmark-valid' Benchmark
Invoke-ValidatorCase 'benchmark-missing-samples' Benchmark {
    param($bundle)
    $bundle.telemetry = @($bundle.telemetry | Where-Object { $_.sample -lt 10 -or $_.sample -gt 20 })
    $bundle.result.telemetry.sampleCount = $bundle.telemetry.Count
} 'missing, duplicated, or out of order'
Invoke-ValidatorCase 'benchmark-concealed-time-gap' Benchmark {
    param($bundle)
    $bundle.telemetry = @($bundle.telemetry | Where-Object { $_.sample -lt 10 -or $_.sample -gt 20 })
    for ($i = 0; $i -lt $bundle.telemetry.Count; $i++) { $bundle.telemetry[$i].sample = $i }
    $bundle.result.telemetry.sampleCount = $bundle.telemetry.Count
} 'actual heartbeat gap over two seconds'
Invoke-ValidatorCase 'benchmark-forged-recorded-gap' Benchmark { param($bundle) $bundle.telemetry[4].gapMilliseconds = 1 } 'recorded gap differs from its timestamps'
Invoke-ValidatorCase 'benchmark-reversed-timestamp' Benchmark { param($bundle) $bundle.telemetry[4].startedUtc = $bundle.telemetry[2].startedUtc } 'timestamp is duplicated or out of order'
Invoke-ValidatorCase 'benchmark-duplicate-index' Benchmark { param($bundle) $bundle.telemetry[4].sample = 3 } 'missing, duplicated, or out of order'

Invoke-ValidatorCase 'soak-valid' Soak
Invoke-ValidatorCase 'soak-missing-xlsx-settled' Soak { param($bundle) [System.IO.File]::Delete((Join-Path $bundle.directory 'probe/xlsx-settled.json')) } 'evidence file is missing: probe/xlsx-settled.json'
Invoke-ValidatorCase 'soak-missing-xlsm-settled' Soak { param($bundle) [System.IO.File]::Delete((Join-Path $bundle.directory 'probe/xlsm-settled.json')) } 'evidence file is missing: probe/xlsm-settled.json'
Invoke-ValidatorCase 'soak-missing-baseline' Soak { param($bundle) [System.IO.File]::Delete((Join-Path $bundle.directory 'probe/xlsx-baseline.json')) } 'evidence file is missing: probe/xlsx-baseline.json'
Invoke-ValidatorCase 'soak-inexact-baseline' Soak {
    param($bundle)
    $path = Join-Path $bundle.directory 'probe/xlsx-baseline.json'
    $probe = Get-Content $path -Raw | ConvertFrom-Json
    $probe.currentSequence = 1
    Write-TestJson $path $probe
} 'baseline was not silent sequence zero'
Invoke-ValidatorCase 'soak-duplicate-settled-version' Soak {
    param($bundle)
    $path = Join-Path $bundle.directory 'probe/xlsx-settled.json'
    $probe = Get-Content $path -Raw | ConvertFrom-Json
    $probe.versionCount = 11
    Write-TestJson $path $probe
} 'settled sequence or version counts differ'
Invoke-ValidatorCase 'soak-settled-warning' Soak {
    param($bundle)
    $path = Join-Path $bundle.directory 'probe/xlsm-settled.json'
    $probe = Get-Content $path -Raw | ConvertFrom-Json
    $probe.workbookStatus = 'Warning'
    Write-TestJson $path $probe
} 'did not settle Active without errors'
Invoke-ValidatorCase 'soak-settled-wrong-hash' Soak {
    param($bundle)
    $path = Join-Path $bundle.directory 'probe/xlsm-settled.json'
    $probe = Get-Content $path -Raw | ConvertFrom-Json
    $probe.currentHash = 'C' * 64
    Write-TestJson $path $probe
} 'settled hashes differ from the final save'
Invoke-ValidatorCase 'soak-settled-wrong-value' Soak {
    param($bundle)
    $path = Join-Path $bundle.directory 'probe/xlsx-settled.json'
    $probe = Get-Content $path -Raw | ConvertFrom-Json
    $probe.cellChange.afterJson = '{"literalValue":"wrong"}'
    Write-TestJson $path $probe
} 'settled literal values differ from the final save'
Invoke-ValidatorCase 'soak-missing-final-xlsx' Soak { param($bundle) [System.IO.File]::Delete((Join-Path $bundle.directory 'fixtures/Soak.xlsx')) } 'evidence file is missing: fixtures/Soak.xlsx'
Invoke-ValidatorCase 'soak-changed-final-xlsx' Soak { param($bundle) Write-TestText (Join-Path $bundle.directory 'fixtures/Soak.xlsx') 'changed unrecorded fixture bytes' } 'xlsx retained final workbook hash differs from the final save'
Invoke-ValidatorCase 'soak-changed-final-xlsm' Soak { param($bundle) Add-TestZipEntry (Join-Path $bundle.directory 'fixtures/Soak Macro.xlsm') 'changed.txt' 'changed non-VBA bytes' } 'xlsm retained final workbook hash differs from the final save'

Write-Output "BENCHMARK_SOAK_VALIDATOR_REGRESSIONS_PASS|tests=$testCount|artifacts=$testRoot"
