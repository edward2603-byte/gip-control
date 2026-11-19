param(
    [string]$WorkbookPath = (Join-Path (Split-Path -Parent $PSScriptRoot) 'templates\GIP_Navigator_TEMPLATE.xlsm')
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

function Get-ContentCount($Range) {
    $count = 0
    foreach ($row in $Range.Value2) {
        foreach ($value in $row) {
            if ($null -ne $value -and "$value".Trim().Length -gt 0) { $count++ }
        }
    }
    return $count
}

$repoRoot = Split-Path -Parent $PSScriptRoot
$WorkbookPath = [IO.Path]::GetFullPath($WorkbookPath)
if (-not (Test-Path -LiteralPath $WorkbookPath)) { throw "Не найден файл: $WorkbookPath" }

$temporaryWorkbook = Join-Path ([IO.Path]::GetTempPath()) ("GIP_Navigator_stage123_{0}.xlsm" -f [guid]::NewGuid().ToString('N'))
$backupPath = Join-Path (Split-Path -Parent $WorkbookPath) ("{0}_before_stage123_{1}{2}" -f [IO.Path]::GetFileNameWithoutExtension($WorkbookPath), (Get-Date -Format 'yyyyMMdd_HHmmss'), [IO.Path]::GetExtension($WorkbookPath))
$excel = $null
$oldBook = $null
$newBook = $null

try {
    # Do not silently remove historical input.  Empty legacy sheets may be
    # replaced by the redesigned structure; populated ones require a separate
    # data migration decision.
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $oldBook = $excel.Workbooks.Open($WorkbookPath, 0, $true)
    $legacyRanges = @{
        '07_Акты' = 'A5:H304'
        '08_Внешние_согласования' = 'A5:F54'
        '03_СР_проекта' = 'A5:O304'
        '04_Согласование_плана' = 'A5:F304'
    }
    foreach ($sheetName in $legacyRanges.Keys) {
        $exists = $false
        foreach ($candidateSheet in $oldBook.Worksheets) { if ($candidateSheet.Name -eq $sheetName) { $exists = $true; break } }
        if ($exists -and (Get-ContentCount $oldBook.Worksheets.Item($sheetName).Range($legacyRanges[$sheetName])) -gt 0) {
            throw "Лист '$sheetName' содержит пользовательские записи. Автоматическая замена остановлена: данные не будут потеряны."
        }
    }
    $kpValues = $oldBook.Worksheets.Item('01_КП').Range('A5:I79').Value2
    $oldBook.Close($false)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($oldBook)
    $oldBook = $null

    & (Join-Path $repoRoot 'tools\build_xlsm.ps1') -OutputPath $temporaryWorkbook
    if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $temporaryWorkbook)) { throw 'Не удалось подготовить обновлённую XLSM-книгу.' }

    $newBook = $excel.Workbooks.Open($temporaryWorkbook)
    $targetKp = $newBook.Worksheets.Item('01_КП')
    for ($row = 1; $row -le 75; $row++) {
        for ($column = 1; $column -le 9; $column++) {
            $value = $kpValues[$row, $column]
            if ($null -ne $value -and "$value".Length -gt 0) {
                $address = "{0}{1}" -f [char](64 + $column), ($row + 4)
                # Excel's COM binder on this Windows build only accepts a
                # string VARIANT for cells inside the formatted KP table.  The
                # worksheet converts numeric text back to a number, retaining
                # the date and currency formats of the target cells.
                $targetKp.Range($address).Value2 = [string]$value
            }
        }
    }
    $newBook.Worksheets.Item('01_КП').Range('H5:K79').Calculate()
    $excel.CalculateFullRebuild()
    $excel.Run("'$($newBook.Name)'!RefreshGantt")
    $newBook.Save()
    $newBook.Close($true)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($newBook)
    $newBook = $null
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    $excel = $null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()

    # Atomic replacement: the original goes to the timestamped backup only if
    # the fully built replacement already exists and Excel has closed it.
    [IO.File]::Replace($temporaryWorkbook, $WorkbookPath, $backupPath, $true)
    Write-Output "UPDATED=$WorkbookPath"
    Write-Output "BACKUP=$backupPath"
}
finally {
    if ($oldBook) { $oldBook.Close($false) }
    if ($newBook) { $newBook.Close($false) }
    if ($excel) { $excel.Quit() }
    foreach ($item in @($oldBook, $newBook, $excel)) {
        if ($null -ne $item -and [Runtime.InteropServices.Marshal]::IsComObject($item)) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($item) }
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if (Test-Path -LiteralPath $temporaryWorkbook) { Remove-Item -LiteralPath $temporaryWorkbook -Force }
}
