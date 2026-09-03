[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string] $ResultPath,
    [Parameter(Mandatory)] [string] $InstallerPath,
    [Parameter(Mandatory)] [ValidatePattern('^[A-Fa-f0-9]{64}$')] [string] $ExpectedApplicationSha256,
    [Parameter(Mandatory)] [string] $ProbePath,
    [Parameter(Mandatory)] [ValidatePattern('^[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}$')] [string] $ExpectedOuterRunEvidenceId
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$path = (Resolve-Path $ResultPath).Path
$root = Split-Path -Parent $path
$installer = (Resolve-Path $InstallerPath).Path
$probe = (Resolve-Path $ProbePath).Path
$installerHash = (Get-FileHash -LiteralPath $installer -Algorithm SHA256).Hash.ToUpperInvariant()
$probeHash = (Get-FileHash -LiteralPath $probe -Algorithm SHA256).Hash.ToUpperInvariant()
$applicationHash = $ExpectedApplicationSha256.ToUpperInvariant()
$result = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json

function Require-Condition {
    param([bool] $Condition, [string] $Message)
    if (-not $Condition) { throw "Invalid installed recovery-matrix evidence: $Message" }
}

function Resolve-EvidenceFile {
    param([object] $Record)
    Require-Condition ($null -ne $Record) 'an evidence record is missing'
    Require-Condition (-not [string]::IsNullOrWhiteSpace([string]$Record.path)) 'an evidence path is empty'
    Require-Condition (-not [System.IO.Path]::IsPathRooted([string]$Record.path)) "evidence path must be relative: $($Record.path)"
    $full = [System.IO.Path]::GetFullPath((Join-Path $root ([string]$Record.path).Replace('/','\')))
    $prefix = $root.TrimEnd('\') + '\'
    Require-Condition ($full.StartsWith($prefix,[StringComparison]::OrdinalIgnoreCase)) "evidence path escapes the run directory: $($Record.path)"
    Require-Condition (Test-Path $full -PathType Leaf) "evidence file is missing: $($Record.path)"
    $item = Get-Item -LiteralPath $full
    Require-Condition ($item.Length -gt 0) "evidence file is empty: $($Record.path)"
    Require-Condition ($item.Length -eq [long]$Record.bytes) "evidence byte count differs: $($Record.path)"
    $hash = (Get-FileHash -LiteralPath $full -Algorithm SHA256).Hash.ToUpperInvariant()
    Require-Condition ($hash -eq ([string]$Record.sha256).ToUpperInvariant()) "evidence hash differs: $($Record.path)"
    $full
}

function Get-Scenario {
    param([string] $Id)
    $matches = @($result.scenarios | Where-Object { $_.id -eq $Id })
    Require-Condition ($matches.Count -eq 1) "scenario '$Id' is missing or duplicated"
    $matches[0]
}

function Get-Phase {
    param([object] $Scenario, [string] $Name)
    $matches = @($Scenario.phases | Where-Object { $_.name -eq $Name })
    Require-Condition ($matches.Count -eq 1) "scenario '$($Scenario.id)' phase '$Name' is missing or duplicated"
    $matches[0]
}

function Validate-PhaseEvidence {
    param([object] $Scenario, [object] $Phase)
    $observed = [DateTime]::Parse($Phase.observedUtc).ToUniversalTime()
    Require-Condition ($observed -ge [DateTime]::Parse($Scenario.startedUtc).ToUniversalTime() -and $observed -le [DateTime]::Parse($Scenario.finishedUtc).ToUniversalTime()) "$($Scenario.id)/$($Phase.name) timestamp is outside its scenario"
    foreach ($record in @($Phase.database)) { $null = Resolve-EvidenceFile $record }
    Require-Condition (@($Phase.database).Count -ge 1) "$($Scenario.id)/$($Phase.name) has no retained database bytes"
    if ($null -ne $Phase.probe) {
        $probePath = Resolve-EvidenceFile $Phase.probe
        $raw = Get-Content -LiteralPath $probePath -Raw | ConvertFrom-Json
        Require-Condition ($raw.passed) "$($Scenario.id)/$($Phase.name) raw external probe did not pass"
        Require-Condition (($raw | ConvertTo-Json -Depth 20 -Compress) -eq ($Phase.result | ConvertTo-Json -Depth 20 -Compress)) "$($Scenario.id)/$($Phase.name) summary differs from raw external probe"
    }
    if ($null -ne $Phase.source -and $Phase.source.exists -and -not $Phase.source.inaccessible) {
        $sourcePath = Resolve-EvidenceFile $Phase.source
        $null = $sourcePath
        if (-not [string]::IsNullOrWhiteSpace([string]$Phase.expectedSourceSha256)) {
            Require-Condition ($Phase.source.sha256 -eq $Phase.expectedSourceSha256) "$($Scenario.id)/$($Phase.name) retained source differs from its expected hash"
        }
    }
    if ($null -ne $Phase.report) { $null = Resolve-EvidenceFile $Phase.report }
}

function Require-State {
    param([object] $Phase, [long] $Sequence, [long] $Versions, [long] $Errors, [string] $Status, [string] $ReportStatus = '')
    Require-Condition ($null -ne $Phase.probe) "$($Phase.name) does not contain an external probe"
    Require-Condition ($Phase.result.currentSequence -eq $Sequence) "$($Phase.name) sequence differs"
    Require-Condition ($Phase.result.versionCount -eq $Versions) "$($Phase.name) version count differs"
    Require-Condition ($Phase.result.distinctVersionHashCount -eq $Versions) "$($Phase.name) version hashes are not unique"
    Require-Condition ($Phase.result.errorCount -eq $Errors) "$($Phase.name) error count differs"
    Require-Condition ($Phase.result.workbookStatus -eq $Status) "$($Phase.name) status differs"
    if ($Sequence -eq 0) {
        Require-Condition ($null -eq $Phase.result.latestVersion) "$($Phase.name) silent baseline unexpectedly has a version"
    }
    else {
        Require-Condition ($Phase.result.latestVersion.sequence -eq $Sequence -and $Phase.result.latestVersion.sha256 -ceq $Phase.result.currentHash) "$($Phase.name) latest version differs from the current sequence or hash"
    }
    if (-not [string]::IsNullOrWhiteSpace($ReportStatus)) {
        Require-Condition ($Phase.result.latestVersion.reportStatus -eq $ReportStatus) "$($Phase.name) report status differs"
        if ($ReportStatus -eq 'Ready') {
            Require-Condition ($null -ne $Phase.report) "$($Phase.name) does not retain its ready Markdown report"
            $null = Resolve-EvidenceFile $Phase.report
        }
        elseif ($ReportStatus -eq 'Pending') {
            Require-Condition ($null -eq $Phase.report) "$($Phase.name) unexpectedly retains a ready report while marked Pending"
        }
    }
}

function Require-LiteralDelta {
    param(
        [object] $Phase,
        [string] $BeforeValue,
        [string] $AfterValue,
        [string] $PreviousHash,
        [string] $Address = 'A1',
        [string] $ExpectedCurrentHash = ''
    )
    $raw = Get-Content -LiteralPath (Resolve-EvidenceFile $Phase.probe) -Raw | ConvertFrom-Json
    Require-Condition ($raw.passed) "$($Phase.name) external semantic probe failed"
    Require-Condition (($raw | ConvertTo-Json -Depth 20 -Compress) -eq ($Phase.result | ConvertTo-Json -Depth 20 -Compress)) "$($Phase.name) semantic summary differs from its raw probe"
    if ([string]::IsNullOrWhiteSpace($ExpectedCurrentHash)) { $ExpectedCurrentHash = [string]$Phase.expectedSourceSha256 }
    Require-Condition ($ExpectedCurrentHash -match '^[A-F0-9]{64}$' -and $raw.currentHash -ceq $ExpectedCurrentHash) "$($Phase.name) captured the wrong source hash"
    Require-Condition ($raw.latestVersion.sequence -eq $raw.currentSequence -and $raw.latestVersion.sha256 -ceq $raw.currentHash) "$($Phase.name) latest semantic version differs"
    Require-Condition ($raw.latestVersion.cellChangeCount -eq 1 -and $raw.latestVersion.sheetChangeCount -eq 0) "$($Phase.name) must have exactly one cell delta and zero sheet deltas"
    Require-Condition ($null -ne $raw.cellChange -and $raw.cellChange.sheetName -ceq 'Recovery' -and $raw.cellChange.address -ceq $Address) "$($Phase.name) changed the wrong cell"
    Require-Condition ($raw.cellChange.kinds -ceq 'LiteralChanged') "$($Phase.name) cell delta kinds differ"
    $before = $raw.cellChange.beforeJson | ConvertFrom-Json
    $after = $raw.cellChange.afterJson | ConvertFrom-Json
    Require-Condition ($null -ne $before -and $null -ne $after -and $before.literalValue -ceq $BeforeValue -and $after.literalValue -ceq $AfterValue) "$($Phase.name) exact before/after literal values differ"
    Require-Condition ($before.sheetId -gt 0 -and $before.sheetId -eq $after.sheetId) "$($Phase.name) changed the stable sheet identity"
    foreach ($state in @($before,$after)) {
        Require-Condition ($state.sheetName -ceq 'Recovery' -and $state.address -ceq $Address -and $state.cellType -ceq 'SharedString') "$($Phase.name) before/after identity or type differs"
        foreach ($field in @('formulaText','formulaType','formulaReference','formulaSharedIndex','formulaXml','cachedResult')) {
            Require-Condition ($null -eq $state.$field) "$($Phase.name) literal state contains unexpected $field"
        }
    }
    if ($raw.latestVersion.reportStatus -eq 'Ready') {
        $markdown = [System.IO.File]::ReadAllText((Resolve-EvidenceFile $Phase.report))
        $lines = @(
            "- Version: $($raw.currentSequence)",
            "- Previous SHA-256: ``$PreviousHash``",
            "- Current SHA-256: ``$ExpectedCurrentHash``",
            '| Sheet changes | 0 |','| Changed cells | 1 |','| Literal value changes | 1 |',
            '| Formula text changes | 0 |','| Calculated result changes | 0 |','| Cell type changes | 0 |',
            "### Recovery!$Address",'**Changes:** Literal changed',
            '| Type | <pre>SharedString</pre> | <pre>SharedString</pre> |',
            "| Literal value | <pre>$BeforeValue</pre> | <pre>$AfterValue</pre> |")
        foreach ($line in $lines) {
            Require-Condition ([regex]::Matches($markdown,('(?m)^' + [regex]::Escape($line) + '\r?$')).Count -eq 1) "$($Phase.name) Markdown line differs or is duplicated: $line"
        }
        foreach ($field in @('Formula','Formula type','Formula reference','Shared formula index','Stored formula XML','Calculated result')) {
            $line = '| ' + $field + ' | `<none>` | `<none>` |'
            Require-Condition ([regex]::Matches($markdown,('(?m)^' + [regex]::Escape($line) + '\r?$')).Count -eq 1) "$($Phase.name) Markdown contains unexpected formula data: $field"
        }
        Require-Condition ([regex]::Matches($markdown,'(?m)^### ').Count -eq 1) "$($Phase.name) Markdown has an extra cell detail"
        Require-Condition ($markdown.IndexOf('No tracked changes.',[StringComparison]::Ordinal) -lt 0) "$($Phase.name) Markdown incorrectly reports no tracked changes"
    }
}

Require-Condition ($result.schemaVersion -eq 2) 'schema version must be 2'
Require-Condition ($result.evidenceId -match '^[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}$') 'evidence identity is missing or malformed'
Require-Condition ($result.outerRunEvidenceId -eq $ExpectedOuterRunEvidenceId.ToLowerInvariant()) 'outer run evidence identity differs'
Require-Condition ($result.success) 'runner did not pass'
Require-Condition ([string]::IsNullOrWhiteSpace([string]$result.failure)) 'runner retained a failure'
Require-Condition ($result.confirmation.disposableCleanVm) 'disposable clean VM confirmation is absent'
Require-Condition ($result.confirmation.freshEvidenceDirectory) 'fresh evidence confirmation is absent'
Require-Condition ($result.confirmation.localEvidenceDirectory) 'evidence was not local to the VM'
Require-Condition ($result.environment.powershellVersion -match '^5\.1\.') 'runner was not Windows PowerShell 5.1'
Require-Condition ($result.environment.powershellEdition -eq 'Desktop') 'runner was not Desktop PowerShell'
Require-Condition ($result.candidate.installerSha256 -eq $installerHash) 'installer hash differs from the file supplied to the validator'
Require-Condition ($result.candidate.expectedInstallerSha256 -eq $installerHash) 'frozen installer hash differs'
Require-Condition ($result.candidate.applicationSha256 -eq $applicationHash) 'installed executable hash differs from the frozen hash'
Require-Condition ($result.candidate.expectedApplicationSha256 -eq $applicationHash) 'frozen installed executable hash differs'
Require-Condition ($result.candidate.probeSha256 -eq $probeHash) 'external AcceptanceProbe hash differs'
Require-Condition ($result.thresholds.shortLockSeconds -ge 5 -and $result.thresholds.shortLockSeconds -le 15) 'short-lock threshold is outside the accepted approximately-five-second range'
Require-Condition ($result.thresholds.failureExposureSeconds -ge 65) 'failure exposure is shorter than the stable-copy timeout'
Require-Condition ($result.thresholds.largeWorkbookRows -ge 5000) 'large workbook fixture is too small'
Require-Condition ($result.thresholds.largeWorkbookColumns -eq 20 -and $result.thresholds.largeWorkbookCells -eq ($result.thresholds.largeWorkbookRows * 20)) 'large workbook dimensions differ'
$runStarted = [DateTime]::Parse([string]$result.startedUtc).ToUniversalTime()
$runFinished = [DateTime]::Parse([string]$result.finishedUtc).ToUniversalTime()
Require-Condition ($runFinished -gt $runStarted -and $result.durationSeconds -gt 0) 'run timestamps are invalid'

$requiredIds = @(
    'compatible-exclusive-lock-recovery','atomic-replacement','save-as-path-behavior',
    'autosave-write-burst-dedup-order','two-workbook-isolation','missing-file-restoration',
    'stopped-app-restart-reconciliation','unwritable-report-recovery',
    'interrupted-slow-large-capture-recovery','corrupt-package-rejection','encrypted-package-rejection')
$actualIds = @($result.scenarios | ForEach-Object { [string]$_.id })
Require-Condition ($actualIds.Count -eq $requiredIds.Count) 'the matrix must contain exactly eleven scenarios'
Require-Condition (@(Compare-Object $requiredIds $actualIds).Count -eq 0) 'the required scenario set differs'
for ($index = 0; $index -lt $requiredIds.Count; $index++) {
    Require-Condition ($actualIds[$index] -eq $requiredIds[$index]) "scenario index $index is reordered"
}

$previousScenarioFinished = $runStarted
foreach ($scenario in @($result.scenarios)) {
    Require-Condition ($scenario.passed) "scenario '$($scenario.id)' did not pass"
    Require-Condition ([string]::IsNullOrWhiteSpace([string]$scenario.failure)) "scenario '$($scenario.id)' retained a failure"
    Require-Condition (@($scenario.phases).Count -ge 2) "scenario '$($scenario.id)' has too few phases"
    $scenarioStarted = [DateTime]::Parse([string]$scenario.startedUtc).ToUniversalTime()
    $scenarioFinished = [DateTime]::Parse([string]$scenario.finishedUtc).ToUniversalTime()
    Require-Condition ($scenarioStarted -ge $previousScenarioFinished -and $scenarioFinished -gt $scenarioStarted -and $scenarioFinished -le $runFinished) "scenario '$($scenario.id)' timestamps are outside the run or reordered"
    $previousObserved = $scenarioStarted
    foreach ($phase in @($scenario.phases)) {
        $observed = [DateTime]::Parse([string]$phase.observedUtc).ToUniversalTime()
        Require-Condition ($observed -ge $previousObserved) "scenario '$($scenario.id)' phase '$($phase.name)' is reordered"
        Validate-PhaseEvidence -Scenario $scenario -Phase $phase
        $previousObserved = $observed
    }
    $previousScenarioFinished = $scenarioFinished
    $null = Resolve-EvidenceFile $scenario.uiEvidence.screenshot
    $null = Resolve-EvidenceFile $scenario.uiEvidence.tree
}

$expectedPhases = @{
    'compatible-exclusive-lock-recovery' = @('baseline','compatible-lock-captured','compatible-lock-window','exclusive-lock-held','exclusive-lock-window','exclusive-lock-recovered')
    'atomic-replacement' = @('baseline','replacement-captured')
    'save-as-path-behavior' = @('baseline','original-path-unchanged','save-as-destination-retained')
    'autosave-write-burst-dedup-order' = @('baseline','ordered-write-manifest','burst-deduplicated')
    'two-workbook-isolation' = @('baseline-a','baseline-b','unlocked-b-captured','locked-a-not-advanced','locked-a-recovered')
    'missing-file-restoration' = @('baseline','missing-detected','restoration-captured')
    'stopped-app-restart-reconciliation' = @('baseline','change-while-stopped','restart-reconciled')
    'unwritable-report-recovery' = @('baseline','report-pending','report-target-obstruction','pending-report-recovered')
    'interrupted-slow-large-capture-recovery' = @('baseline','large-capture-interrupted','large-capture-recovered')
    'corrupt-package-rejection' = @('baseline','corrupt-rejected','valid-package-restored')
    'encrypted-package-rejection' = @('baseline','encrypted-rejected','encrypted-fixture-provenance','unencrypted-package-restored')
}
foreach ($scenario in @($result.scenarios)) {
    $actualPhases = @($scenario.phases | ForEach-Object { [string]$_.name })
    $wantedPhases = @($expectedPhases[$scenario.id])
    Require-Condition ($actualPhases.Count -eq $wantedPhases.Count -and @(Compare-Object $wantedPhases $actualPhases).Count -eq 0) "scenario '$($scenario.id)' phase set differs"
    for ($index = 0; $index -lt $wantedPhases.Count; $index++) {
        Require-Condition ($actualPhases[$index] -eq $wantedPhases[$index]) "scenario '$($scenario.id)' phase index $index is reordered"
    }
}

$s = Get-Scenario 'compatible-exclusive-lock-recovery'
Require-State (Get-Phase $s 'baseline') 0 0 0 'Active'
Require-State (Get-Phase $s 'compatible-lock-captured') 1 1 0 'Active' 'Ready'
Require-LiteralDelta (Get-Phase $s 'compatible-lock-captured') 'baseline' 'compatible-lock-save' (Get-Phase $s 'baseline').result.currentHash
$compatibleWindow = Get-Phase $s 'compatible-lock-window'
Require-Condition ($compatibleWindow.detail.heldSeconds -ge $result.thresholds.shortLockSeconds -and $compatibleWindow.detail.heldSeconds -le ($result.thresholds.shortLockSeconds + 2)) 'compatible lock was not retained for the required approximately-five-second interval'
Require-State (Get-Phase $s 'exclusive-lock-held') 1 1 0 'Processing' 'Ready'
Require-LiteralDelta (Get-Phase $s 'exclusive-lock-held') 'baseline' 'compatible-lock-save' (Get-Phase $s 'baseline').result.currentHash -ExpectedCurrentHash (Get-Phase $s 'compatible-lock-captured').result.currentHash
$exclusiveWindow = Get-Phase $s 'exclusive-lock-window'
Require-Condition ($exclusiveWindow.detail.heldSeconds -ge $result.thresholds.shortLockSeconds -and $exclusiveWindow.detail.heldSeconds -le ($result.thresholds.shortLockSeconds + 2)) 'exclusive lock was not retained for the required approximately-five-second interval'
Require-State (Get-Phase $s 'exclusive-lock-recovered') 2 2 0 'Active' 'Ready'
Require-LiteralDelta (Get-Phase $s 'exclusive-lock-recovered') 'compatible-lock-save' 'exclusive-lock-save' (Get-Phase $s 'compatible-lock-captured').result.currentHash

$s = Get-Scenario 'atomic-replacement'
Require-State (Get-Phase $s 'replacement-captured') 1 1 0 'Active' 'Ready'
Require-LiteralDelta (Get-Phase $s 'replacement-captured') 'baseline' 'atomic-replacement-value' (Get-Phase $s 'baseline').result.currentHash

$s = Get-Scenario 'save-as-path-behavior'
$p = Get-Phase $s 'original-path-unchanged'
Require-State $p 0 0 0 'Active'
Require-Condition ($p.result.currentHash -eq $p.expectedSourceSha256) 'Save As altered or advanced the original tracked path'
$null = Get-Phase $s 'save-as-destination-retained'

$s = Get-Scenario 'autosave-write-burst-dedup-order'
$manifest = Get-Phase $s 'ordered-write-manifest'
Require-Condition (@($manifest.detail.writes).Count -eq 5) 'write burst does not contain exactly five ordered states'
for ($index = 0; $index -lt 5; $index++) {
    Require-Condition ($manifest.detail.writes[$index].order -eq ($index + 1)) 'write-burst ordering differs'
    $null = Resolve-EvidenceFile $manifest.detail.writes[$index]
}
$burst = Get-Phase $s 'burst-deduplicated'
Require-State $burst 1 1 0 'Active' 'Ready'
Require-LiteralDelta $burst 'baseline' 'burst-5' (Get-Phase $s 'baseline').result.currentHash
Require-Condition ($burst.result.currentHash -eq $manifest.detail.writes[4].sha256) 'write burst did not settle on its final ordered state'

$s = Get-Scenario 'two-workbook-isolation'
Require-State (Get-Phase $s 'unlocked-b-captured') 1 1 0 'Active' 'Ready'
Require-LiteralDelta (Get-Phase $s 'unlocked-b-captured') 'baseline' 'free-b' (Get-Phase $s 'baseline-b').result.currentHash
Require-State (Get-Phase $s 'locked-a-not-advanced') 0 0 0 'Processing'
Require-State (Get-Phase $s 'locked-a-recovered') 1 1 0 'Active' 'Ready'
Require-LiteralDelta (Get-Phase $s 'locked-a-recovered') 'baseline' 'locked-a' (Get-Phase $s 'baseline-a').result.currentHash

$s = Get-Scenario 'missing-file-restoration'
$missing = Get-Phase $s 'missing-detected'
Require-State $missing 0 0 1 'Missing'
Require-Condition (-not $missing.source.exists) 'missing-file evidence says the source existed'
Require-State (Get-Phase $s 'restoration-captured') 1 1 1 'Active' 'Ready'
Require-LiteralDelta (Get-Phase $s 'restoration-captured') 'baseline' 'restored-after-missing' (Get-Phase $s 'baseline').result.currentHash

$s = Get-Scenario 'stopped-app-restart-reconciliation'
$stopped = Get-Phase $s 'change-while-stopped'
Require-Condition ($stopped.detail.databaseSha256Before -eq $stopped.detail.databaseSha256After) 'database changed while the app was stopped'
Require-Condition ((Resolve-EvidenceFile $stopped.detail.source) -ne $null) 'stopped-app source evidence is missing'
Require-Condition ($stopped.detail.source.sha256 -eq $stopped.detail.sourceSha256) 'stopped-app source hash differs'
Require-State (Get-Phase $s 'restart-reconciled') 1 1 0 'Active' 'Ready'
Require-LiteralDelta (Get-Phase $s 'restart-reconciled') 'baseline' 'changed-while-stopped' (Get-Phase $s 'baseline').result.currentHash

$s = Get-Scenario 'unwritable-report-recovery'
Require-State (Get-Phase $s 'report-pending') 1 1 1 'Warning' 'Pending'
Require-LiteralDelta (Get-Phase $s 'report-pending') 'baseline' 'pending-report' (Get-Phase $s 'baseline').result.currentHash
$obstruction = Get-Phase $s 'report-target-obstruction'
Require-Condition ((Resolve-EvidenceFile $obstruction.detail.obstruction) -ne $null) 'report-target obstruction evidence is missing'
Require-Condition ($obstruction.detail.reportStatus -eq 'Pending') 'report-target obstruction was not retained while the report was Pending'
Require-State (Get-Phase $s 'pending-report-recovered') 1 1 1 'Active' 'Ready'
Require-LiteralDelta (Get-Phase $s 'pending-report-recovered') 'baseline' 'pending-report' (Get-Phase $s 'baseline').result.currentHash

$s = Get-Scenario 'interrupted-slow-large-capture-recovery'
$interrupted = Get-Phase $s 'large-capture-interrupted'
Require-State $interrupted 0 0 1 'Warning'
Require-Condition ($interrupted.result.currentHash -ne $interrupted.expectedSourceSha256) 'interrupted large capture advanced the baseline hash'
Require-State (Get-Phase $s 'large-capture-recovered') 1 1 1 'Active' 'Ready'
Require-LiteralDelta (Get-Phase $s 'large-capture-recovered') "large-baseline-$($result.thresholds.largeWorkbookRows - 1)-19" 'large-recovery-change' (Get-Phase $s 'baseline').result.currentHash -Address "T$($result.thresholds.largeWorkbookRows)"

foreach ($id in @('corrupt-package-rejection','encrypted-package-rejection')) {
    $s = Get-Scenario $id
    $rejectionName = if ($id -eq 'corrupt-package-rejection') { 'corrupt-rejected' } else { 'encrypted-rejected' }
    $restorationName = if ($id -eq 'corrupt-package-rejection') { 'valid-package-restored' } else { 'unencrypted-package-restored' }
    $baseline = Get-Phase $s 'baseline'
    $rejected = Get-Phase $s $rejectionName
    Require-State $rejected 0 0 1 'Warning'
    Require-Condition ($rejected.result.currentHash -eq $baseline.result.currentHash) "$id advanced its baseline hash"
    Require-Condition ($rejected.result.currentHash -ne $rejected.expectedSourceSha256) "$id accepted the rejected package hash"
    $restored = Get-Phase $s $restorationName
    Require-State $restored 0 0 1 'Active'
    Require-Condition ($restored.result.currentHash -eq $baseline.result.currentHash) "$id did not preserve the baseline after restoration"
}
$encrypted = Get-Scenario 'encrypted-package-rejection'
$provenance = Get-Phase $encrypted 'encrypted-fixture-provenance'
Require-Condition ($provenance.detail.method -in @('excel-saveas-password','supplied-fixture')) 'encrypted fixture provenance is not fail-closed'
Require-Condition ((Resolve-EvidenceFile $provenance.detail.fixture) -ne $null) 'encrypted fixture evidence is missing'
Require-Condition ($provenance.detail.fixture.sha256 -eq $provenance.detail.sha256) 'encrypted fixture provenance hash differs'

$null = Resolve-EvidenceFile $result.transcript
Write-Output "INSTALLED_RECOVERY_MATRIX_VALID|result=$path|evidenceId=$($result.evidenceId)"
