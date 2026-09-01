[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$OutputPath,
    [int]$Rows = 25000,
    [int]$Columns = 20
)

$ErrorActionPreference = 'Stop'
$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()
    $worksheet = $workbook.Worksheets.Item(1)
    $worksheet.Name = 'Performance'
    $range = $worksheet.Range($worksheet.Cells.Item(1, 1), $worksheet.Cells.Item($Rows, $Columns))
    $range.FormulaR1C1 = '=ROW()+COLUMN()'
    $workbook.SaveAs($OutputPath, 51)
    $workbook.Close($false)
    $workbook = $null
    Write-Host "LARGE_FIXTURE_CREATED|rows=$Rows|columns=$Columns|cells=$($Rows * $Columns)|$OutputPath"
}
finally {
    if ($workbook) { $workbook.Close($false) }
    if ($excel) { $excel.Quit() }
    if ($workbook) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook) }
    if ($excel) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
