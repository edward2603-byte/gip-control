param(
    [string]$WorkbookPath = (Join-Path (Split-Path -Parent $PSScriptRoot) 'templates\GIP_Navigator_TEMPLATE.xlsm')
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent $PSScriptRoot
$WorkbookPath = [IO.Path]::GetFullPath($WorkbookPath)
if (-not (Test-Path -LiteralPath $WorkbookPath)) { throw "Workbook not found: $WorkbookPath" }

$backupPath = Join-Path (Split-Path -Parent $WorkbookPath) ("{0}_before_gantt_fix_{1}{2}" -f [IO.Path]::GetFileNameWithoutExtension($WorkbookPath), (Get-Date -Format 'yyyyMMdd_HHmmss'), [IO.Path]::GetExtension($WorkbookPath))
$excel = $null
$book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Open($WorkbookPath, 0, $false)
    $book.SaveCopyAs($backupPath)

    $moduleText = [IO.File]::ReadAllText((Join-Path $repoRoot 'vba\GIPNavigator.bas'), [Text.UTF8Encoding]::new($false))
    $module = $book.VBProject.VBComponents.Item('GIPNavigator').CodeModule
    if ($module.CountOfLines -gt 0) { $module.DeleteLines(1, $module.CountOfLines) }
    $module.AddFromString($moduleText)

    $workbookClassText = [IO.File]::ReadAllText((Join-Path $repoRoot 'vba\ThisWorkbook.cls'), [Text.UTF8Encoding]::new($false))
    $workbookModule = $book.VBProject.VBComponents.Item($book.CodeName).CodeModule
    if ($workbookModule.CountOfLines -gt 0) { $workbookModule.DeleteLines(1, $workbookModule.CountOfLines) }
    $workbookModule.AddFromString($workbookClassText)

    $book.ForceFullCalculation = $true
    $excel.Calculation = -4105 # xlCalculationAutomatic
    $excel.CalculateFullRebuild()
    $excel.Run("'$($book.Name)'!RefreshGantt")
    $book.Save()
    Write-Output "UPDATED=$WorkbookPath"
    Write-Output "BACKUP=$backupPath"
}
finally {
    if ($book) { $book.Close($false) }
    if ($excel) { $excel.Quit() }
    foreach ($item in @($book, $excel)) {
        if ($null -ne $item -and [Runtime.InteropServices.Marshal]::IsComObject($item)) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($item) }
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
