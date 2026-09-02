[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string] $ResultPath,
    [Parameter(Mandatory)] [string] $InstallerPath,
    [Parameter(Mandatory)] [ValidatePattern('^[A-Fa-f0-9]{64}$')] [string] $ExpectedApplicationSha256,
    [Parameter(Mandatory)] [string] $ProbePath,
    [Parameter(Mandatory)] [string] $XlsmFixture
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$path = (Resolve-Path $ResultPath).Path
$root = Split-Path -Parent $path
$installer = (Resolve-Path $InstallerPath).Path
$probe = (Resolve-Path $ProbePath).Path
$sourceXlsm = (Resolve-Path $XlsmFixture).Path
if ([System.IO.Path]::GetExtension($sourceXlsm) -ne '.xlsm') { throw 'XlsmFixture must use the .xlsm extension.' }
$installerHash = (Get-FileHash $installer -Algorithm SHA256).Hash.ToUpperInvariant()
$applicationHash = $ExpectedApplicationSha256.ToUpperInvariant()
$probeHash = (Get-FileHash $probe -Algorithm SHA256).Hash.ToUpperInvariant()
$result = Get-Content $path -Raw | ConvertFrom-Json
$emDash = [char]0x2014

function Require-Condition {
    param([bool] $Condition, [string] $Message)
    if (-not $Condition) { throw "Invalid installed semantic-matrix evidence: $Message" }
}

function Resolve-EvidenceFile {
    param([string] $RelativePath, [string] $ExpectedSha256)
    Require-Condition (-not [string]::IsNullOrWhiteSpace($RelativePath)) 'an evidence path is empty'
    Require-Condition (-not [System.IO.Path]::IsPathRooted($RelativePath)) "evidence path must be relative: $RelativePath"
    $fullPath = [System.IO.Path]::GetFullPath((Join-Path $root $RelativePath))
    $rootPrefix = $root.TrimEnd('\') + '\'
    Require-Condition ($fullPath.StartsWith($rootPrefix, [StringComparison]::OrdinalIgnoreCase)) "evidence path escapes the matrix directory: $RelativePath"
    Require-Condition (Test-Path $fullPath -PathType Leaf) "evidence file is missing: $RelativePath"
    Require-Condition ((Get-Item $fullPath).Length -gt 0) "evidence file is empty: $RelativePath"
    if (-not [string]::IsNullOrWhiteSpace($ExpectedSha256)) {
        $actualHash = (Get-FileHash $fullPath -Algorithm SHA256).Hash.ToUpperInvariant()
        Require-Condition ($actualHash -eq $ExpectedSha256.ToUpperInvariant()) "evidence hash differs for $RelativePath"
    }
    $fullPath
}

function Get-ZipEntrySha256 {
    param([string] $Path, [string] $EntryName, [switch] $Optional)
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $entry = $archive.GetEntry($EntryName)
        if (-not $entry) {
            if ($Optional) { return $null }
            throw "Archive entry not found: $EntryName in $Path"
        }
        $stream = $entry.Open()
        $algorithm = [System.Security.Cryptography.SHA256]::Create()
        try { -join @($algorithm.ComputeHash($stream) | ForEach-Object { $_.ToString('X2') }) }
        finally { $algorithm.Dispose(); $stream.Dispose() }
    }
    finally { $archive.Dispose() }
}

function Read-Probe {
    param([object] $Phase, [int] $Index)
    $probePaths = @($Phase.probes)
    Require-Condition ($Index -ge 0 -and $Index -lt $probePaths.Count) "phase $($Phase.sequence) is missing probe index $Index"
    $probePath = Resolve-EvidenceFile -RelativePath $probePaths[$Index].path -ExpectedSha256 $probePaths[$Index].sha256
    $probeResult = Get-Content $probePath -Raw | ConvertFrom-Json
    Require-Condition ($probeResult.passed) "phase $($Phase.sequence) probe $($Index + 1) failed"
    Require-Condition ($probeResult.workbookStatus -eq 'Active') "phase $($Phase.sequence) workbook was not Active"
    Require-Condition ($probeResult.currentSequence -eq $Phase.sequence) "phase $($Phase.sequence) probe sequence differs"
    Require-Condition ($probeResult.currentHash -eq $Phase.workbookSha256) "phase $($Phase.sequence) probe hash differs"
    Require-Condition ($probeResult.errorCount -eq 0) "phase $($Phase.sequence) has capture errors"
    Require-Condition ([string]::IsNullOrWhiteSpace($probeResult.lastError)) "phase $($Phase.sequence) has a current capture error"
    Require-Condition ($probeResult.versionCount -eq $Phase.sequence) "phase $($Phase.sequence) version count differs"
    Require-Condition ($probeResult.distinctVersionHashCount -eq $probeResult.versionCount) "phase $($Phase.sequence) has duplicate captured hashes"
    if ($Phase.sequence -eq 0) {
        Require-Condition ($null -eq $probeResult.latestVersion) 'silent baseline unexpectedly has a version record'
    }
    else {
        Require-Condition ($probeResult.latestVersion.sequence -eq $Phase.sequence) "phase $($Phase.sequence) latest version differs"
        Require-Condition ($probeResult.latestVersion.sha256 -eq $Phase.workbookSha256) "phase $($Phase.sequence) latest-version hash differs"
        Require-Condition ($probeResult.latestVersion.reportStatus -eq 'Ready') "phase $($Phase.sequence) report was not Ready"
        Require-Condition ($probeResult.latestVersion.cellChangeCount -eq $Phase.expectedCellChangeCount) "phase $($Phase.sequence) SQLite cell count differs"
        Require-Condition ($probeResult.latestVersion.sheetChangeCount -eq $Phase.expectedSheetChangeCount) "phase $($Phase.sequence) SQLite sheet count differs"
    }
    $probeResult
}

function Require-Kinds {
    param([string] $Kinds, [string[]] $ExpectedKinds, [string] $Context)
    $actual = @($Kinds.Split(',') | Sort-Object)
    $expected = @($ExpectedKinds | Sort-Object)
    Require-Condition ($actual.Count -eq $expected.Count) "$Context has an unexpected number of change kinds: $Kinds"
    Require-Condition (@(Compare-Object $expected $actual).Count -eq 0) "$Context kinds differ: expected $($expected -join ','), found $Kinds"
}

function Read-CellState {
    param([object] $Json)
    if ($null -eq $Json -or [string]::IsNullOrWhiteSpace([string]$Json)) { return $null }
    [string]$Json | ConvertFrom-Json
}

function Require-CellDelta {
    param(
        [object] $ProbeResult,
        [string] $Address,
        [string[]] $Kinds,
        [object] $Before,
        [object] $After,
        [string] $Context
    )
    Require-Condition ($ProbeResult.cellChange.sheetName -eq 'Matrix') "$Context sheet name differs"
    Require-Condition ($ProbeResult.cellChange.address -eq $Address) "$Context address differs"
    Require-Kinds -Kinds $ProbeResult.cellChange.kinds -ExpectedKinds $Kinds -Context $Context
    $beforeState = Read-CellState $ProbeResult.cellChange.beforeJson
    $afterState = Read-CellState $ProbeResult.cellChange.afterJson
    if ($null -eq $Before) {
        Require-Condition ($null -eq $beforeState) "$Context before state should be absent"
    }
    else {
        Require-Condition ($null -ne $beforeState) "$Context before state is absent"
        foreach ($property in $Before.PSObject.Properties) {
            Require-Condition ($beforeState.($property.Name) -eq $property.Value) "$Context before $($property.Name) differs"
        }
    }
    if ($null -eq $After) {
        Require-Condition ($null -eq $afterState) "$Context after state should be absent"
    }
    else {
        Require-Condition ($null -ne $afterState) "$Context after state is absent"
        foreach ($property in $After.PSObject.Properties) {
            Require-Condition ($afterState.($property.Name) -eq $property.Value) "$Context after $($property.Name) differs"
        }
    }
}

function Require-SheetDelta {
    param(
        [object] $ProbeResult,
        [string] $Kind,
        [object] $Before,
        [object] $After,
        [string] $Context
    )
    Require-Condition ($ProbeResult.sheetChange.kind -eq $Kind) "$Context kind differs"
    $beforeState = Read-CellState $ProbeResult.sheetChange.beforeJson
    $afterState = Read-CellState $ProbeResult.sheetChange.afterJson
    if ($null -eq $Before) {
        Require-Condition ($null -eq $beforeState) "$Context before state should be absent"
    }
    else {
        Require-Condition ($null -ne $beforeState) "$Context before state is absent"
        foreach ($property in $Before.PSObject.Properties) {
            Require-Condition ($beforeState.($property.Name) -eq $property.Value) "$Context before $($property.Name) differs"
        }
    }
    if ($null -eq $After) {
        Require-Condition ($null -eq $afterState) "$Context after state should be absent"
    }
    else {
        Require-Condition ($null -ne $afterState) "$Context after state is absent"
        foreach ($property in $After.PSObject.Properties) {
            Require-Condition ($afterState.($property.Name) -eq $property.Value) "$Context after $($property.Name) differs"
        }
    }
}

function Require-Report {
    param(
        [object] $Phase,
        [long] $LiteralCount,
        [long] $FormulaCount,
        [long] $ResultCount,
        [long] $TypeCount,
        [string[]] $RequiredText = @(),
        [switch] $NoTrackedChanges
    )
    Require-Condition ($null -ne $Phase.report) "phase $($Phase.sequence) has no portable report record"
    $reportPath = Resolve-EvidenceFile -RelativePath $Phase.report.path -ExpectedSha256 $Phase.report.sha256
    $markdown = [System.IO.File]::ReadAllText($reportPath)
    Require-Condition ($markdown.IndexOf("- Version: $($Phase.sequence)", [StringComparison]::Ordinal) -ge 0) "phase $($Phase.sequence) Markdown version differs"
    Require-Condition ($markdown.IndexOf("- Current SHA-256: ``$($Phase.workbookSha256)``", [StringComparison]::Ordinal) -ge 0) "phase $($Phase.sequence) Markdown hash differs"
    foreach ($text in $RequiredText) {
        Require-Condition ($markdown.IndexOf($text, [StringComparison]::Ordinal) -ge 0) "phase $($Phase.sequence) Markdown omits: $text"
    }
    if ($NoTrackedChanges) {
        Require-Condition ($markdown.IndexOf('No tracked changes. The workbook file changed, but only ignored content or package metadata differed.', [StringComparison]::Ordinal) -ge 0) "phase $($Phase.sequence) is not the exact no-tracked-changes report"
        Require-Condition ($markdown.IndexOf('| Category | Count |', [StringComparison]::Ordinal) -lt 0) "phase $($Phase.sequence) no-change report unexpectedly has tracked summary rows"
    }
    else {
        $rows = @{
            'Sheet changes' = $Phase.expectedSheetChangeCount
            'Changed cells' = $Phase.expectedCellChangeCount
            'Literal value changes' = $LiteralCount
            'Formula text changes' = $FormulaCount
            'Calculated result changes' = $ResultCount
            'Cell type changes' = $TypeCount
        }
        foreach ($entry in $rows.GetEnumerator()) {
            $pattern = '(?m)^\| ' + [regex]::Escape([string]$entry.Key) + ' \| ' + [regex]::Escape([string]$entry.Value) + ' \|\r?$'
            Require-Condition ([regex]::Matches($markdown, $pattern).Count -eq 1) "phase $($Phase.sequence) Markdown summary row differs for $($entry.Key)"
        }
    }
}

function Test-FormatSemanticMatrix {
    param([string] $Format, [object[]] $FormatPhases)

    $p0 = Read-Probe $FormatPhases[0] 0
    Require-Condition (@($FormatPhases[0].probes).Count -eq 1) "$Format baseline must have one probe"

    $p1 = Read-Probe $FormatPhases[1] 0
    Require-CellDelta $p1 'A1' @('FormulaAdded') $null ([pscustomobject]@{ formulaText='1+1'; cachedResult='2'; cellType='formula:Number' }) "$Format formula add"
    Require-Report $FormatPhases[1] 0 1 0 0 @('### Matrix!A1','Formula added','<pre>1+1</pre>','<pre>2</pre>')

    $p2 = Read-Probe $FormatPhases[2] 0
    Require-CellDelta $p2 'A1' @('FormulaChanged','FormulaResultChanged') ([pscustomobject]@{ formulaText='1+1'; cachedResult='2' }) ([pscustomobject]@{ formulaText='1+2'; cachedResult='3' }) "$Format formula edit"
    Require-Report $FormatPhases[2] 0 1 1 0 @('### Matrix!A1','Formula changed','Formula result changed','<pre>1+1</pre>','<pre>1+2</pre>','<pre>3</pre>')

    $p3 = Read-Probe $FormatPhases[3] 0
    Require-CellDelta $p3 'A1' @('FormulaRemoved') ([pscustomobject]@{ formulaText='1+2'; cachedResult='3' }) $null "$Format formula delete"
    Require-Report $FormatPhases[3] 0 1 0 0 @('### Matrix!A1','Formula removed','<pre>1+2</pre>')

    Require-Condition (@($FormatPhases[4].probes).Count -eq 2) "$Format formula-result phase must have independent formula and precedent probes"
    $p4Formula = Read-Probe $FormatPhases[4] 0
    $p4Precedent = Read-Probe $FormatPhases[4] 1
    Require-CellDelta $p4Formula 'C1' @('FormulaResultChanged') ([pscustomobject]@{ formulaText='D1'; cachedResult='1' }) ([pscustomobject]@{ formulaText='D1'; cachedResult='2' }) "$Format formula result only"
    Require-CellDelta $p4Precedent 'D1' @('LiteralChanged') ([pscustomobject]@{ literalValue='1'; cellType='Number' }) ([pscustomobject]@{ literalValue='2'; cellType='Number' }) "$Format formula-result precedent"
    Require-Report $FormatPhases[4] 1 0 1 0 @('### Matrix!C1','### Matrix!D1','Formula result changed','Literal changed','<pre>D1</pre>','<pre>1</pre>','<pre>2</pre>')

    $p5 = Read-Probe $FormatPhases[5] 0
    Require-CellDelta $p5 'E1' @('CellTypeChanged','LiteralChanged') ([pscustomobject]@{ literalValue='123'; cellType='Number' }) ([pscustomobject]@{ literalValue='EDT-TEXT-123'; cellType='SharedString' }) "$Format literal and type transition"
    Require-Report $FormatPhases[5] 1 0 0 1 @('### Matrix!E1','Cell type changed','Literal changed','<pre>123</pre>','<pre>EDT-TEXT-123</pre>')

    $p6 = Read-Probe $FormatPhases[6] 0
    Require-Condition ($p6.latestVersion.cellChangeCount -eq 0 -and $p6.latestVersion.sheetChangeCount -eq 0) "$Format style-only save has a semantic delta"
    Require-Report $FormatPhases[6] 0 0 0 0 @('None.') -NoTrackedChanges

    $p7 = Read-Probe $FormatPhases[7] 0
    Require-SheetDelta $p7 'Added' $null ([pscustomobject]@{ name='Matrix Added'; position=1; visibility='Visible' }) "$Format sheet add"
    Require-Report $FormatPhases[7] 0 0 0 0 @("| Added | $emDash | Matrix Added (position 2, Visible) |")

    $p8 = Read-Probe $FormatPhases[8] 0
    Require-SheetDelta $p8 'Renamed' ([pscustomobject]@{ name='Matrix Added'; position=1; visibility='Visible' }) ([pscustomobject]@{ name='Matrix Renamed'; position=1; visibility='Visible' }) "$Format sheet rename"
    Require-Condition ($p8.sheetChange.sheetId -eq $p7.sheetChange.sheetId) "$Format sheet rename did not preserve the stable sheet identity"
    Require-Report $FormatPhases[8] 0 0 0 0 @('| Renamed | Matrix Added (position 2, Visible) | Matrix Renamed (position 2, Visible) |')

    $p9 = Read-Probe $FormatPhases[9] 0
    Require-SheetDelta $p9 'Reordered' ([pscustomobject]@{ name='Matrix Renamed'; position=1; visibility='Visible' }) ([pscustomobject]@{ name='Matrix Renamed'; position=0; visibility='Visible' }) "$Format sheet reorder"
    Require-Condition ($p9.sheetChange.sheetId -eq $p7.sheetChange.sheetId) "$Format sheet reorder did not preserve the stable sheet identity"
    Require-Report $FormatPhases[9] 0 0 0 0 @('| Reordered | Matrix Renamed (position 2, Visible) | Matrix Renamed (position 1, Visible) |','| Reordered | Matrix (position 1, Visible) | Matrix (position 2, Visible) |')

    $p10 = Read-Probe $FormatPhases[10] 0
    Require-SheetDelta $p10 'VisibilityChanged' ([pscustomobject]@{ name='Matrix Renamed'; position=0; visibility='Visible' }) ([pscustomobject]@{ name='Matrix Renamed'; position=0; visibility='Hidden' }) "$Format sheet hide"
    Require-Condition ($p10.sheetChange.sheetId -eq $p7.sheetChange.sheetId) "$Format sheet hide did not preserve the stable sheet identity"
    Require-Report $FormatPhases[10] 0 0 0 0 @('| Visibility changed | Matrix Renamed (position 1, Visible) | Matrix Renamed (position 1, Hidden) |')

    $p11 = Read-Probe $FormatPhases[11] 0
    Require-SheetDelta $p11 'VisibilityChanged' ([pscustomobject]@{ name='Matrix Renamed'; position=0; visibility='Hidden' }) ([pscustomobject]@{ name='Matrix Renamed'; position=0; visibility='Visible' }) "$Format sheet unhide"
    Require-Condition ($p11.sheetChange.sheetId -eq $p7.sheetChange.sheetId) "$Format sheet unhide did not preserve the stable sheet identity"
    Require-Report $FormatPhases[11] 0 0 0 0 @('| Visibility changed | Matrix Renamed (position 1, Hidden) | Matrix Renamed (position 1, Visible) |')

    $p12 = Read-Probe $FormatPhases[12] 0
    Require-SheetDelta $p12 'Removed' ([pscustomobject]@{ name='Matrix Renamed'; position=0; visibility='Visible' }) $null "$Format sheet remove"
    Require-Condition ($p12.sheetChange.sheetId -eq $p7.sheetChange.sheetId) "$Format sheet removal did not preserve the stable sheet identity"
    Require-Report $FormatPhases[12] 0 0 0 0 @(
        "| Removed | Matrix Renamed (position 1, Visible) | $emDash |",
        '| Reordered | Matrix (position 2, Visible) | Matrix (position 1, Visible) |')

    foreach ($phase in $FormatPhases) {
        Require-Condition (@($phase.probes).Count -eq $(if ($phase.sequence -eq 4) { 2 } else { 1 })) "$Format phase $($phase.sequence) has an unexpected probe count"
    }
}

Require-Condition ($result.schemaVersion -eq 1) 'schemaVersion must be 1'
Require-Condition ($result.gate -eq 'installed-real-excel-semantic-matrix') 'gate identity is wrong'
Require-Condition ($result.status -eq 'Passed') 'status must be Passed'
Require-Condition ($result.evidenceId -match '^[0-9a-fA-F-]{36}$') 'evidenceId is missing or malformed'
$startedUtc = [DateTime]::Parse($result.startedUtc).ToUniversalTime()
$finishedUtc = [DateTime]::Parse($result.finishedUtc).ToUniversalTime()
Require-Condition ($finishedUtc -gt $startedUtc) 'run timestamps are invalid'
Require-Condition ($result.durationSeconds -gt 0) 'duration must be positive'
Require-Condition ($result.candidate.installerSha256 -eq $installerHash) 'installer hash does not match the exact candidate'
Require-Condition ($result.candidate.expectedInstallerSha256 -eq $installerHash) 'frozen installer hash differs'
Require-Condition ($result.candidate.applicationSha256 -eq $applicationHash) 'installed executable hash differs'
Require-Condition ($result.candidate.expectedApplicationSha256 -eq $applicationHash) 'frozen executable hash differs'
Require-Condition ($result.candidate.probeSha256 -eq $probeHash) 'external acceptance-probe hash differs'
$workbooks = @($result.workbooks)
Require-Condition ($workbooks.Count -eq 2) 'exactly one xlsx and one xlsm workbook record are required'
Require-Condition (@($workbooks | Where-Object format -eq 'xlsx').Count -eq 1) 'xlsx workbook record is missing or duplicated'
Require-Condition (@($workbooks | Where-Object format -eq 'xlsm').Count -eq 1) 'xlsm workbook record is missing or duplicated'
Require-Condition (@($workbooks | Where-Object finalSequence -ne 12).Count -eq 0) 'both workbook final sequences must be 12'
$xlsxRecord = $workbooks | Where-Object format -eq 'xlsx'
$xlsmRecord = $workbooks | Where-Object format -eq 'xlsm'
$fixtureMacroSha256 = Get-ZipEntrySha256 -Path $sourceXlsm -EntryName 'xl/vbaProject.bin'
Require-Condition ($null -eq $xlsxRecord.sourceMacroSha256 -and $null -eq $xlsxRecord.baselineMacroSha256) 'xlsx workbook record unexpectedly has a VBA hash'
Require-Condition ($xlsmRecord.sourceMacroSha256 -eq $fixtureMacroSha256) 'xlsm source VBA hash differs from the supplied deterministic fixture'
Require-Condition ($xlsmRecord.baselineMacroSha256 -eq $fixtureMacroSha256) 'xlsm fixture preparation changed vbaProject.bin'

$phases = @($result.phases)
Require-Condition ($phases.Count -eq 26) 'two silent baselines plus exactly twenty-four saves are required'
$expectedNames = @(
    'silent baseline',
    'formula add',
    'formula edit',
    'formula delete',
    'formula result only',
    'literal and type transition',
    'style only',
    'sheet add',
    'sheet rename',
    'sheet reorder',
    'sheet hide',
    'sheet unhide',
    'sheet remove')
$expectedCellCounts = @(0,1,1,1,2,1,0,0,0,0,0,0,0)
$expectedSheetCounts = @(0,0,0,0,0,0,0,1,1,2,1,1,2)
$previousCapturedUtc = $startedUtc
$allPaths = [System.Collections.Generic.List[string]]::new()

foreach ($format in @('xlsx','xlsm')) {
    $formatPhases = @($phases | Where-Object format -eq $format)
    Require-Condition ($formatPhases.Count -eq 13) "$format must contain one baseline and twelve saves"
    for ($index = 0; $index -lt $formatPhases.Count; $index++) {
        $phase = $formatPhases[$index]
        Require-Condition ($phase.sequence -eq $index) "$format phase index $index has the wrong sequence"
        Require-Condition ($phase.name -eq $expectedNames[$index]) "$format phase $index name differs"
        Require-Condition ($phase.expectedCellChangeCount -eq $expectedCellCounts[$index]) "$format phase $index expected cell count differs"
        Require-Condition ($phase.expectedSheetChangeCount -eq $expectedSheetCounts[$index]) "$format phase $index expected sheet count differs"
        if ($index -eq 0) {
            Require-Condition ($phase.transport -eq 'Excel COM fixture setup before tracking; installed-app UIA baseline registration') "$format baseline transport record differs"
            Require-Condition ($null -eq $phase.ctrlSaveUtc) "$format baseline unexpectedly records Ctrl+S"
        }
        else {
            Require-Condition ($phase.transport -eq 'visible Excel keyboard/UIA mutation and Ctrl+S') "$format phase $index was not recorded as a visible Excel keyboard/UIA save"
            Require-Condition (-not [string]::IsNullOrWhiteSpace($phase.ctrlSaveUtc)) "$format phase $index does not record Ctrl+S"
            $ctrlSaveUtc = [DateTime]::Parse($phase.ctrlSaveUtc).ToUniversalTime()
            Require-Condition ($ctrlSaveUtc -ge $startedUtc -and $ctrlSaveUtc -le $finishedUtc) "$format phase $index Ctrl+S timestamp is outside the run"
        }
        $actionUtc = [DateTime]::Parse($phase.actionStartedUtc).ToUniversalTime()
        $capturedUtc = [DateTime]::Parse($phase.capturedUtc).ToUniversalTime()
        Require-Condition ($actionUtc -ge $startedUtc -and $actionUtc -le $finishedUtc) "$format phase $index action timestamp is outside the fresh run"
        Require-Condition ($capturedUtc -ge $previousCapturedUtc -and $capturedUtc -le $finishedUtc) "$format phase $index capture timestamp is out of order"
        Require-Condition ($phase.captureMilliseconds -ge 0 -and $phase.captureMilliseconds -le 60000) "$format phase $index capture duration is invalid"
        $previousCapturedUtc = $capturedUtc
        Require-Condition ($phase.workbookSha256 -match '^[A-F0-9]{64}$') "$format phase $index workbook hash is malformed"
        Require-Condition ($phase.workbookEvidence.sha256 -eq $phase.workbookSha256) "$format phase $index portable workbook record differs"
        Require-Condition ([System.IO.Path]::GetExtension($phase.workbookEvidence.path) -eq ".$format") "$format phase $index workbook evidence has the wrong extension"
        $workbookPath = Resolve-EvidenceFile -RelativePath $phase.workbookEvidence.path -ExpectedSha256 $phase.workbookEvidence.sha256
        if ($format -eq 'xlsm') {
            $retainedMacroSha256 = Get-ZipEntrySha256 -Path $workbookPath -EntryName 'xl/vbaProject.bin'
            Require-Condition ($phase.macroSha256 -eq $fixtureMacroSha256) "xlsm phase $index reported VBA hash differs"
            Require-Condition ($retainedMacroSha256 -eq $fixtureMacroSha256) "xlsm phase $index changed vbaProject.bin"
        }
        else {
            Require-Condition ($null -eq $phase.macroSha256) "xlsx phase $index unexpectedly reports a VBA hash"
            Require-Condition ($null -eq (Get-ZipEntrySha256 -Path $workbookPath -EntryName 'xl/vbaProject.bin' -Optional)) "xlsx phase $index unexpectedly contains vbaProject.bin"
        }
        $allPaths.Add($phase.workbookEvidence.path)
        $screenshot = Resolve-EvidenceFile -RelativePath $phase.uiEvidence.screenshot -ExpectedSha256 $phase.uiEvidence.screenshotSha256
        $uiaPath = Resolve-EvidenceFile -RelativePath $phase.uiEvidence.uiaTree -ExpectedSha256 $phase.uiEvidence.uiaTreeSha256
        $allPaths.Add($phase.uiEvidence.screenshot)
        $allPaths.Add($phase.uiEvidence.uiaTree)
        Require-Condition ((Get-Item $screenshot).Length -gt 1024) "$format phase $index screenshot is implausibly small"
        $uiaText = [System.IO.File]::ReadAllText($uiaPath)
        Require-Condition ($uiaText.IndexOf("Installed Semantic Matrix.$format", [StringComparison]::Ordinal) -ge 0) "$format phase $index UIA evidence omits the tracked workbook"
        foreach ($probeRecord in @($phase.probes)) { $allPaths.Add([string]$probeRecord.path) }
        if ($null -ne $phase.report) { $allPaths.Add([string]$phase.report.path) }
    }
    Require-Condition (@($formatPhases.workbookSha256 | Select-Object -Unique).Count -eq 13) "$format baseline and saves do not have thirteen unique stable hashes"
}

Require-Condition (@($allPaths | Select-Object -Unique).Count -eq $allPaths.Count) 'an evidence path is reused by more than one phase'

Test-FormatSemanticMatrix -Format xlsx -FormatPhases @($phases | Where-Object format -eq 'xlsx')
Test-FormatSemanticMatrix -Format xlsm -FormatPhases @($phases | Where-Object format -eq 'xlsm')

$failedAssertions = @($result.assertions | Where-Object { -not $_.passed })
Require-Condition ($failedAssertions.Count -eq 0) 'result contains a failed assertion'
Require-Condition (@($result.assertions).Count -ge 90) 'result does not contain the complete dual-format fail-closed assertion set'
Require-Condition ([string]::IsNullOrWhiteSpace($result.failure)) 'result contains an unhandled failure'
$transcript = Resolve-EvidenceFile -RelativePath $result.transcript.path -ExpectedSha256 $result.transcript.sha256
Require-Condition ((Get-Item $transcript).Length -gt 0) 'transcript is empty'

Write-Output "INSTALLED_SEMANTIC_MATRIX_VALID|$path|$installerHash|$applicationHash|$probeHash|$($result.evidenceId)"
