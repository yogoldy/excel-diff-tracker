[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if ($PSVersionTable.PSVersion.Major -ne 5 -or $PSVersionTable.PSEdition -ne 'Desktop') {
    throw 'Run these isolated regressions with Windows PowerShell 5.1 Desktop.'
}

# Only selected AST expressions/functions are executed. No installed app, Excel,
# workbook, profile, database, UI, or acceptance-runner entry point is accessed.
$repository = Split-Path -Parent $PSScriptRoot
$runnerPath = Join-Path $repository 'scripts\acceptance\Invoke-InstalledRecoveryMatrix.ps1'
$validatorPath = Join-Path $repository 'scripts\acceptance\Test-InstalledRecoveryMatrixResult.ps1'
$semanticPath = Join-Path $repository 'scripts\acceptance\Invoke-InstalledSemanticMatrix.ps1'
$asts = @{}
foreach ($scriptPath in @($runnerPath,$validatorPath,$semanticPath)) {
    $tokens = $null
    $errors = $null
    $asts[$scriptPath] = [System.Management.Automation.Language.Parser]::ParseFile($scriptPath,[ref]$tokens,[ref]$errors)
    if ($errors.Count -gt 0) { throw ($errors | Out-String) }
}

$checks = 0
function Assert-Test {
    param([bool] $Condition, [string] $Name)
    if (-not $Condition) { throw "Regression failed: $Name" }
    $script:checks++
}
function Assert-Rejected {
    param([scriptblock] $Action, [string] $ExpectedMessage, [string] $Name)
    $caught = $null
    try { & $Action } catch { $caught = $_.Exception.Message }
    Assert-Test ($null -ne $caught -and $caught -match $ExpectedMessage) "$Name; observed: $caught"
}

foreach ($entry in @(
    @{ path=$runnerPath; names=@('Assert-RecoveryCondition','Assert-RecoveryLiteralDelta') },
    @{ path=$validatorPath; names=@('Require-Condition','Resolve-EvidenceFile','Require-LiteralDelta','Require-State') }
)) {
    foreach ($name in $entry.names) {
        $functionAst = @($asts[$entry.path].FindAll({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] },$true) | Where-Object Name -eq $name)
        Assert-Test ($functionAst.Count -eq 1) "exact function found: $name"
        . ([scriptblock]::Create($functionAst[0].Extent.Text))
    }
}

$dateEntries = @($asts[$semanticPath].FindAll({ param($node) $node -is [System.Management.Automation.Language.HashtableAst] },$true) |
    ForEach-Object { $_.KeyValuePairs } | Where-Object { $_.Item1.Extent.Text -eq 'ctrlSaveUtc' })
Assert-Test ($dateEntries.Count -eq 1) 'exact semantic timestamp expression found'
$recordDate = [scriptblock]::Create('param([Nullable[DateTime]] $CtrlSaveUtc) ' + $dateEntries[0].Item2.Extent.Text)
Assert-Test ($null -eq (& $recordDate $null)) 'semantic baseline records a null Ctrl+S timestamp'
$savedUtc = [DateTime]::SpecifyKind([DateTime]'2026-09-03T12:34:56',[DateTimeKind]::Utc)
Assert-Test ((& $recordDate $savedUtc) -ceq $savedUtc.ToString('O')) 'semantic save records the exact UTC timestamp'

$missingCommands = @($asts[$runnerPath].FindAll({ param($node) $node -is [System.Management.Automation.Language.CommandAst] -and $node.GetCommandName() -eq 'Invoke-RecoveryScenario' },$true) |
    Where-Object { $_.CommandElements[1].Value -eq 'missing-file-restoration' })
Assert-Test ($missingCommands.Count -eq 1) 'missing-file scenario found'
$missingBody = $missingCommands[0].CommandElements[2].ScriptBlock.Extent.Text
$missingAction = [scriptblock]::Create($missingBody.Substring(1,$missingBody.Length - 2))
& {
    $activeScenario = [pscustomobject]@{ id='missing-file-restoration'; phases=[System.Collections.Generic.List[object]]::new() }
    $fixtures = 'C:\isolated-mocked-fixtures'
    $FailureExposureSeconds = 65
    $script:restorationObserved = $false
    function Register-ScenarioWorkbook { 'C:\isolated-mocked-fixtures\missing.xlsx' }
    function Move-Item { param($LiteralPath,$Destination) }
    function Set-WorkbookValue { param($Path,$Value) }
    function Get-AcceptanceFileSha256 { param($Path) 'B' * 64 }
    function Add-ProbePhase {
        param($Name,$WorkbookPath,$Arguments,$TimeoutSeconds,$ExpectedSourceSha256,$ExpectedBeforeValue,$ExpectedAfterValue)
        if ($Name -eq 'missing-detected') {
            $activeScenario.phases.Add([pscustomobject]@{ name=$Name; source=[pscustomobject]@{ exists=$false } })
            # Real AcceptanceProbe results have no source property.
            return [pscustomobject]@{ currentSequence=0; workbookStatus='Missing' }
        }
        $script:restorationObserved = ($Name -eq 'restoration-captured' -and $ExpectedBeforeValue -ceq 'baseline' -and $ExpectedAfterValue -ceq 'restored-after-missing')
    }
    & $missingAction
    Assert-Test $script:restorationObserved 'missing-file scenario reaches restoration with a raw probe that has no source property'
}

# Require every scenario phase that has a captured version to request exact probe
# contents. Evaluate only its Add-ProbePhase command against a parameter mock.
$expectedCalls = @{
    'compatible-lock-captured'=@('A1','baseline','compatible-lock-save')
    'exclusive-lock-held'=@('A1','baseline','compatible-lock-save')
    'exclusive-lock-recovered'=@('A1','compatible-lock-save','exclusive-lock-save')
    'replacement-captured'=@('A1','baseline','atomic-replacement-value')
    'burst-deduplicated'=@('A1','baseline','burst-5')
    'unlocked-b-captured'=@('A1','baseline','free-b')
    'locked-a-recovered'=@('A1','baseline','locked-a')
    'restoration-captured'=@('A1','baseline','restored-after-missing')
    'restart-reconciled'=@('A1','baseline','changed-while-stopped')
    'report-pending'=@('A1','baseline','pending-report')
    'pending-report-recovered'=@('A1','baseline','pending-report')
    'large-capture-recovered'=@('T25000','large-baseline-24999-19','large-recovery-change')
}
$deltaCommands = @($asts[$runnerPath].FindAll({ param($node) $node -is [System.Management.Automation.Language.CommandAst] -and $node.GetCommandName() -eq 'Add-ProbePhase' },$true) |
    Where-Object { @($_.CommandElements | Where-Object { $_ -is [System.Management.Automation.Language.CommandParameterAst] -and $_.ParameterName -eq 'ExpectedAfterValue' }).Count -eq 1 })
Assert-Test ($deltaCommands.Count -eq $expectedCalls.Count) 'all 12 captured-version phases request exact probe values'
& {
    $path=$pathA=$pathB='C:\isolated-mocked-fixtures\source.xlsx'
    $hash=$hash1=$hash2=$hashA=$hashB=$finalHash='B' * 64
    $ShortLockSeconds=5
    $LargeWorkbookRows=25000
    function Add-ProbePhase {
        param($Name,$WorkbookPath,$Arguments,$TimeoutSeconds,$ExpectedSourceSha256,$ExpectedBeforeValue,$ExpectedAfterValue,$ExpectedAddress='A1',[switch]$SourceInaccessible)
        $wanted=$expectedCalls[$Name]
        Assert-Test ($null -ne $wanted -and $ExpectedAddress -ceq $wanted[0] -and $ExpectedBeforeValue -ceq $wanted[1] -and $ExpectedAfterValue -ceq $wanted[2]) "exact fixture delta requested: $Name"
    }
    foreach ($command in $deltaCommands) { & ([scriptblock]::Create($command.Extent.Text)) }
}

$root = Join-Path ([System.IO.Path]::GetTempPath()) ('edt-semantic-recovery-regression-' + [Guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path $root
$evidence = $root
$previousHash = 'A' * 64
$currentHash = 'B' * 64
function Write-TestFile {
    param([string] $Name,[string] $Text)
    $filePath = Join-Path $root $Name
    [System.IO.File]::WriteAllText($filePath,$Text,(New-Object System.Text.UTF8Encoding -ArgumentList $false))
    [pscustomobject]@{ path=$Name; sha256=(Get-FileHash -LiteralPath $filePath -Algorithm SHA256).Hash; bytes=(Get-Item -LiteralPath $filePath).Length }
}
function New-TestPhase {
    param([string] $Address='A1',[string] $BeforeValue='baseline',[string] $AfterValue='atomic-replacement-value',[switch] $Pending)
    $before=[ordered]@{ sheetId=1; sheetName='Recovery'; address=$Address; cellType='SharedString'; literalValue=$BeforeValue; formulaText=$null; formulaType=$null; formulaReference=$null; formulaSharedIndex=$null; formulaXml=$null; cachedResult=$null }
    $after=[ordered]@{}; foreach($key in $before.Keys) { $after[$key]=$before[$key] }; $after.literalValue=$AfterValue
    $raw=[pscustomobject]@{
        passed=$true; currentSequence=1; currentHash=$currentHash; versionCount=1; distinctVersionHashCount=1; errorCount=0; workbookStatus='Active'
        latestVersion=[pscustomobject]@{ sequence=1; sha256=$currentHash; cellChangeCount=1; sheetChangeCount=0; reportStatus=$(if($Pending){'Pending'}else{'Ready'}) }
        cellChange=[pscustomobject]@{ sheetName='Recovery'; address=$Address; kinds='LiteralChanged'; beforeJson=($before|ConvertTo-Json -Compress); afterJson=($after|ConvertTo-Json -Compress) }
    }
    $reportLines=@(
        '- Version: 1',"- Previous SHA-256: ``$previousHash``","- Current SHA-256: ``$currentHash``",
        '| Sheet changes | 0 |','| Changed cells | 1 |','| Literal value changes | 1 |',
        '| Formula text changes | 0 |','| Calculated result changes | 0 |','| Cell type changes | 0 |',
        "### Recovery!$Address",'**Changes:** Literal changed','| Type | <pre>SharedString</pre> | <pre>SharedString</pre> |',
        "| Literal value | <pre>$BeforeValue</pre> | <pre>$AfterValue</pre> |")
    foreach($field in @('Formula','Formula type','Formula reference','Shared formula index','Stored formula XML','Calculated result')) { $reportLines += '| ' + $field + ' | `<none>` | `<none>` |' }
    [pscustomobject]@{ name='synthetic-capture'; expectedSourceSha256=$currentHash; result=$raw; probe=(Write-TestFile 'probe.json' ($raw|ConvertTo-Json -Depth 20)); report=$(if($Pending){$null}else{Write-TestFile 'report.md' ($reportLines -join "`r`n")}) }
}
function Update-TestProbe {
    param([object] $Phase)
    $Phase.probe=Write-TestFile 'probe.json' ($Phase.result|ConvertTo-Json -Depth 20)
}
function Assert-BothReject {
    param([object] $Phase,[string] $Pattern,[string] $Name)
    Assert-Rejected { Assert-RecoveryLiteralDelta $Phase 'A1' 'baseline' 'atomic-replacement-value' } $Pattern "runner rejects $Name"
    Assert-Rejected { Require-LiteralDelta $Phase 'baseline' 'atomic-replacement-value' $previousHash } $Pattern "validator rejects $Name"
}

$phase=New-TestPhase
Assert-RecoveryLiteralDelta $phase 'A1' 'baseline' 'atomic-replacement-value'
Require-LiteralDelta $phase 'baseline' 'atomic-replacement-value' $previousHash
Assert-Test $true 'valid Ready exact delta and Markdown pass both checks'
$phase=New-TestPhase -Pending
Assert-RecoveryLiteralDelta $phase 'A1' 'baseline' 'atomic-replacement-value'
Require-LiteralDelta $phase 'baseline' 'atomic-replacement-value' $previousHash
Assert-Test $true 'valid Pending exact delta passes without a report'
$phase=New-TestPhase -Address 'T25000' -BeforeValue 'large-baseline-24999-19' -AfterValue 'large-recovery-change'
Assert-RecoveryLiteralDelta $phase 'T25000' 'large-baseline-24999-19' 'large-recovery-change'
Require-LiteralDelta $phase 'large-baseline-24999-19' 'large-recovery-change' $previousHash -Address 'T25000'
Assert-Test $true 'large-fixture final-cell delta passes'

$phase=New-TestPhase; $phase.result.latestVersion.cellChangeCount=0; Update-TestProbe $phase
Assert-BothReject $phase 'exactly one' 'zero-change version'
$phase=New-TestPhase; $state=$phase.result.cellChange.beforeJson|ConvertFrom-Json; $state.literalValue='BASELINE'; $phase.result.cellChange.beforeJson=$state|ConvertTo-Json -Compress; Update-TestProbe $phase
Assert-BothReject $phase 'literal' 'incorrect old literal including case'
$phase=New-TestPhase; $state=$phase.result.cellChange.afterJson|ConvertFrom-Json; $state.literalValue='wrong-value'; $phase.result.cellChange.afterJson=$state|ConvertTo-Json -Compress; Update-TestProbe $phase
Assert-BothReject $phase 'literal' 'incorrect new literal'
$phase=New-TestPhase; $phase.result.cellChange.kinds='LiteralChanged,FormulaChanged'; Update-TestProbe $phase
Assert-BothReject $phase 'kind' 'extra change kind'
$phase=New-TestPhase; $phase.result.cellChange.address='B1'; Update-TestProbe $phase
Assert-BothReject $phase 'wrong cell' 'incorrect cell address'
$phase=New-TestPhase; $state=$phase.result.cellChange.afterJson|ConvertFrom-Json; $state.cachedResult='2'; $phase.result.cellChange.afterJson=$state|ConvertTo-Json -Compress; Update-TestProbe $phase
Assert-BothReject $phase 'formula|cachedResult' 'unexpected formula result'
$phase=New-TestPhase; $phase.report=Write-TestFile 'report.md' ([System.IO.File]::ReadAllText((Join-Path $root 'report.md')).Replace('<pre>baseline</pre> | <pre>atomic-replacement-value</pre>','<pre>atomic-replacement-value</pre> | <pre>baseline</pre>'))
Assert-BothReject $phase 'line' 'reversed Markdown before/after values'
$phase=New-TestPhase; $phase.report=Write-TestFile 'report.md' ([System.IO.File]::ReadAllText((Join-Path $root 'report.md')).Replace($previousHash,('C'*64)))
Assert-Rejected { Require-LiteralDelta $phase 'baseline' 'atomic-replacement-value' $previousHash } 'Previous SHA' 'validator rejects wrong report ancestry'
$phase=New-TestPhase; $phase.result.currentHash='C'*64; $phase.result.latestVersion.sha256='C'*64; Update-TestProbe $phase
Assert-Rejected { Require-LiteralDelta $phase 'baseline' 'atomic-replacement-value' $previousHash } 'wrong source hash' 'validator rejects incorrect captured hash'
$phase=New-TestPhase; $phase.result.cellChange.kinds='LiteralAdded'
Assert-Rejected { Require-LiteralDelta $phase 'baseline' 'atomic-replacement-value' $previousHash } 'summary differs' 'validator rejects summary/raw disagreement'
$phase=New-TestPhase; $phase.result.currentSequence=0; $phase.result.versionCount=0; $phase.result.distinctVersionHashCount=0
Assert-Rejected { Require-State $phase 0 0 0 'Active' } 'silent baseline' 'validator rejects a hidden version during baseline'

Write-Output "SEMANTIC_RECOVERY_ISOLATED_REGRESSIONS_PASS|checks=$checks|syntheticFiles=$root"
