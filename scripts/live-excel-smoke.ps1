param(
    [string]$Configuration = "Release",
    [string]$RepositoryRoot
)

$ErrorActionPreference = "Stop"
$root = if ($RepositoryRoot) { (Resolve-Path $RepositoryRoot).Path } else { Split-Path -Parent $PSScriptRoot }
$artifactRoot = Join-Path $root "artifacts\live-excel-smoke"
$toolOutput = Join-Path $root "artifacts\smoke-tool"
$workbookPath = Join-Path $artifactRoot "LiveExcel.xlsx"
$reportPath = Join-Path $artifactRoot "reports"
$databasePath = Join-Path $artifactRoot "history.db"
$readyPath = Join-Path $artifactRoot "ready.signal"
$resultPath = Join-Path $artifactRoot "result.json"
$stdoutPath = Join-Path $artifactRoot "smoke.stdout.log"
$stderrPath = Join-Path $artifactRoot "smoke.stderr.log"

if (Test-Path $artifactRoot) {
    Remove-Item $artifactRoot -Recurse -Force
}
New-Item $artifactRoot -ItemType Directory | Out-Null
New-Item $reportPath -ItemType Directory | Out-Null

function Close-ComObject([object]$value) {
    if ($null -ne $value) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($value)
    }
}

function New-BaselineWorkbook([string]$path) {
    $excel = $null
    $workbook = $null
    $sheet = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Add()
        $sheet = $workbook.Worksheets.Item(1)
        $sheet.Name = "Live Test"
        $sheet.Cells.Item(1, 1).Value2 = 1
        $workbook.SaveAs($path, 51)
        $workbook.Close($false)
    }
    finally {
        if ($null -ne $excel) { $excel.Quit() }
        Close-ComObject $sheet
        Close-ComObject $workbook
        Close-ComObject $excel
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Save-ChangedWorkbook([string]$path) {
    $excel = $null
    $workbook = $null
    $sheet = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Open($path)
        $sheet = $workbook.Worksheets.Item(1)
        $sheet.Cells.Item(1, 1).Value2 = 2
        $sheet.Cells.Item(2, 1).Formula = "=A1+5"
        $workbook.Save()
        $workbook.Close($false)
    }
    finally {
        if ($null -ne $excel) { $excel.Quit() }
        Close-ComObject $sheet
        Close-ComObject $workbook
        Close-ComObject $excel
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

New-BaselineWorkbook $workbookPath
$smokeExe = Join-Path $toolOutput "ExcelDiffTracker.Smoke.exe"
if (-not (Test-Path $smokeExe)) {
    throw "Smoke tool was not published: $smokeExe"
}

$arguments = @($workbookPath, $reportPath, $databasePath, $readyPath, $resultPath)
$process = Start-Process -FilePath $smokeExe -ArgumentList $arguments -PassThru -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath
$deadline = [DateTime]::UtcNow.AddSeconds(30)
while (-not (Test-Path $readyPath) -and [DateTime]::UtcNow -lt $deadline) {
    Start-Sleep -Milliseconds 200
}
if (-not (Test-Path $readyPath)) {
    Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
    throw "Tracker did not establish a baseline. $((Get-Content $stderrPath -Raw -ErrorAction SilentlyContinue))"
}

Save-ChangedWorkbook $workbookPath
if (-not $process.WaitForExit(60000)) {
    Stop-Process -Id $process.Id -Force
    throw "Tracker did not capture the Excel save within 60 seconds."
}
if (-not (Test-Path $resultPath)) {
    throw "Smoke tracker produced no result file. $((Get-Content $stderrPath -Raw -ErrorAction SilentlyContinue))"
}

$result = Get-Content $resultPath -Raw | ConvertFrom-Json
if ($result.kind -ne "Captured" -or $result.versionCount -ne 1 -or $result.errorCount -ne 0 -or $result.cellChanges -lt 2) {
    throw "Unexpected smoke result: $(Get-Content $resultPath -Raw)"
}
if (-not (Test-Path $result.ReportPath)) {
    throw "The expected Markdown report was not created: $($result.ReportPath)"
}
$markdown = Get-Content $result.ReportPath -Raw
if ($markdown -notmatch "Live Test!A1" -or $markdown -notmatch "Live Test!A2" -or $markdown -notmatch "Formula added") {
    throw "The Markdown report does not contain the expected real-Excel changes."
}

Write-Output "LIVE_EXCEL_SMOKE_PASS"
Get-Content $resultPath -Raw
