[CmdletBinding()]
param(
    [string]$WorkbookPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if ([string]::IsNullOrWhiteSpace($WorkbookPath)) {
    $WorkbookPath = Join-Path $PSScriptRoot '..\templates\GIP_Navigator_STAGE0_PROTOTYPE.xlsx'
}

$source = [System.IO.Path]::GetFullPath($WorkbookPath)
if (-not (Test-Path -LiteralPath $source)) { throw "Не найдена книга: $source. Сначала запустите tools\\New-Stage0Prototype.ps1." }
$copy = Join-Path ([System.IO.Path]::GetTempPath()) ("GIP_Navigator_STAGE0_{0}.xlsx" -f [Guid]::NewGuid())
Copy-Item -LiteralPath $source -Destination $copy

$excel = $null; $workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Open($copy)
    $excel.Calculation = -4105
    $sr = $workbook.Worksheets.Item('03_СР_проекта')
    $control = $workbook.Worksheets.Item('06_Контроль')

    # Базовый прогноз гидравлического расчёта после ожидаемых ИД 07.02: 07.02–10.02.
    if ($sr.Range('N4').Text -ne '07.02.2026' -or $sr.Range('O4').Text -ne '10.02.2026') {
        throw "Базовый прогноз SR-03 не совпал: $($sr.Range('N4').Text)–$($sr.Range('O4').Text)"
    }

    # Ожидаемое и фактическое завершения заменяют расчётную дату окончания.
    $sr.Range('K4').Formula = '=DATE(2026,2,16)'
    $excel.CalculateFullRebuild()
    if ($sr.Range('O4').Text -ne '16.02.2026') { throw 'Ожидаемое окончание не изменило прогноз SR-03.' }
    $sr.Range('K4').ClearContents()
    $sr.Range('J4').Formula = '=DATE(2026,2,17)'
    $excel.CalculateFullRebuild()
    if ($sr.Range('O4').Text -ne '17.02.2026') { throw 'Фактическое окончание не изменило прогноз SR-03.' }
    $sr.Range('J4').ClearContents()

    # Фактическое начало SR-03 переносит зависимую цепочку и создаёт риск по КП-2.
    $sr.Range('I4').Formula = '=DATE(2026,2,20)'
    $excel.CalculateFullRebuild()
    if ($sr.Range('O4').Text -ne '23.02.2026') { throw 'Фактическое начало не изменило прогноз SR-03.' }
    if ($control.Range('D6').Text -ne 'Критично: позднее КП') { throw "После задержки нет критичного сигнала: $($control.Range('D6').Text)" }
    if ($workbook.Worksheets.Item('05_График').Range('C6').Text -ne '20.02.2026') { throw 'Источник диаграммы Ганта не обновился после задержки.' }

    # Нет даты ИД не превращается в выдуманную дату и помечается как неопределённость.
    $sr.Range('I4:K4').ClearContents()
    $sr.Range('L4').Value2 = 'Нет данных'
    $sr.Range('M4').ClearContents()
    $excel.CalculateFullRebuild()
    if ($sr.Range('O4').Text -ne '') { throw 'При отсутствии ИД появился фиктивный прогноз.' }
    if ($control.Range('D4').Text -ne 'Неопределённость: нет даты ИД') { throw "Нет сигнала неопределённости: $($control.Range('D4').Text)" }

    if ($sr.Range('O2:O7').FormatConditions.Count -lt 1 -or $workbook.Worksheets.Item('05_График').ChartObjects().Count -lt 1) {
        throw 'В прототипе нет условного форматирования или диаграммы Ганта.'
    }

    $workbook.Close($false)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook); $workbook=$null
    Write-Output 'PASS: Excel автоматически пересчитал факт, риск КП и отсутствие ИД.'
}
finally {
    if ($workbook) { $workbook.Close($false); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) }
    if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
    if (Test-Path -LiteralPath $copy) { Remove-Item -LiteralPath $copy -Force }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
