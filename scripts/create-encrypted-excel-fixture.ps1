[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'
$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()
    $workbook.Worksheets.Item(1).Cells.Item(1, 1).Value2 = 'encrypted acceptance fixture'
    $workbook.SaveAs($OutputPath, 51, 'ExcelDiffTracker-Test-Only')
    $workbook.Close($false)
    $workbook = $null
    Write-Host "ENCRYPTED_FIXTURE_CREATED|$OutputPath"
}
finally {
    if ($workbook) { $workbook.Close($false) }
    if ($excel) { $excel.Quit() }
    if ($workbook) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook) }
    if ($excel) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
