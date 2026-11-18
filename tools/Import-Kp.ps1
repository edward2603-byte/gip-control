[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$SourcePath,

    [string]$MasterPath,

    [switch]$PreviewOnly,
    [switch]$Apply
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($MasterPath)) {
    $MasterPath = Join-Path $PSScriptRoot '..\templates\GIP_Navigator_TEMPLATE.xlsx'
}

function Get-FullPath([string]$Path, [string]$Label) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "$Label не найден: $Path"
    }
    return (Resolve-Path -LiteralPath $Path).Path
}

$source = Get-FullPath $SourcePath 'Файл КП'
$master = Get-FullPath $MasterPath 'Основная книга GIP Navigator'
if ([IO.Path]::GetFullPath($source) -eq [IO.Path]::GetFullPath($master)) {
    throw 'Файл КП и основная книга не могут быть одним файлом.'
}

# Предпросмотр выполняется тем же валидатором, который задаёт структуру импорта.
$python = (Get-Command python -ErrorAction Stop).Source
$validator = Join-Path $PSScriptRoot 'gipnav_import.py'
$previousOutputEncoding = [Console]::OutputEncoding
try {
    [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
    $previewJson = & $python $validator $source
    if ($LASTEXITCODE -ne 0) { throw 'Проверка структуры КП завершилась с ошибкой.' }
    $preview = $previewJson | ConvertFrom-Json
}
finally {
    [Console]::OutputEncoding = $previousOutputEncoding
}

$previewText = [PSCustomObject]@{
    Project = $preview.project
    ContractOrOrder = $preview.contract_or_order
    ImportedStages = @($preview.stages).Count
    TotalRubWithoutVat = $preview.total_rub_without_vat
    IgnoredServiceRows = @($preview.warnings).Count
    DuplicateStageNumbers = @($preview.duplicate_stage_numbers) -join ', '
}

if ($PreviewOnly -or -not $Apply) {
    $previewText | Format-List | Out-String | Write-Output
    if (-not $PreviewOnly) {
        Write-Output 'Изменения не внесены. После проверки запустите команду с ключом -Apply.'
    }
    return
}

# A full rebuild would erase the GIP's selected SR, facts, acts and agreed
# baseline. It is intentionally disabled until the safe in-place updater is
# introduced in Step 5. Preview remains available and never changes a file.
throw 'Безопасное обновление КП в существующей книге пока не включено. Изменения не внесены: используйте -PreviewOnly. Полная пересборка книги намеренно запрещена, чтобы не потерять данные ГИПа.'

$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $workbook = $excel.Workbooks.Open($master)
    $kp = $workbook.Worksheets.Item('01_КП')
    $sr = $workbook.Worksheets.Item('03_СР_проекта')
    $graph = $workbook.Worksheets.Item('05_График')
    $dashboardSheet = $workbook.Worksheets.Item('00_Главная')

    # 01_КП: исходные значения сохраняются типизированными и отображаются полностью.
    $kp.Cells.Clear()
    $kp.Range('A1:G1').Merge()
    $kp.Range('A1').Value = "Календарный план — $($preview.project)"
    $kp.Range('A2:G2').Merge()
    $kp.Range('A2').Value = "Источник: $([IO.Path]::GetFileName($source)) · исходный файл не изменялся · импорт: $(Get-Date -Format 'dd.MM.yyyy HH:mm')"
    $headers = @(
        'Номер договора / Номер Наряд-Заказа',
        'Наименование проекта',
        '№ Этапа',
        'Наименование этапа',
        'Начало работ',
        'Окончание работ',
        'Стоимость этапа, руб. без НДС'
    )
    for ($column = 1; $column -le $headers.Count; $column++) {
        $kp.Cells.Item(3, $column).Value = $headers[$column - 1]
    }

    $row = 4
    $graphRows = @()
    foreach ($item in @($preview.stages)) {
        $kp.Cells.Item($row, 1).Value = [string]$item.contract_or_order
        $kp.Cells.Item($row, 2).Value = [string]$item.project
        $kp.Cells.Item($row, 3).NumberFormat = '@'
        $kp.Cells.Item($row, 3).Value = [string]$item.stage
        $kp.Cells.Item($row, 4).Value = [string]$item.name
        if ($item.start) {
            ($kp.Cells.Item($row, 5)).Value = [double]([datetime]::ParseExact([string]$item.start, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)).ToOADate()
        }
        if ($item.end) {
            ($kp.Cells.Item($row, 6)).Value = [double]([datetime]::ParseExact([string]$item.end, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)).ToOADate()
        }
        if ($null -ne $item.amount_rub_without_vat -and [string]$item.amount_rub_without_vat -ne '') {
            $amountValue = [double]::Parse([string]$item.amount_rub_without_vat, [Globalization.CultureInfo]::InvariantCulture)
            ($kp.Cells.Item($row, 7)).FormulaLocal = '=' + $amountValue.ToString('0.00', [Globalization.CultureInfo]::CurrentCulture)
        }
        if ($item.start -and $item.end) {
            $graphRows += [PSCustomObject]@{ SourceRow = $row; Name = [string]$item.name }
        }
        $row++
    }
    $lastKpRow = $row - 1

    $titleRange = $kp.Range('A1:G1')
    $titleRange.Interior.Color = 0x00B65A12
    $titleRange.Font.Color = 0x00FFFFFF
    $titleRange.Font.Bold = $true
    $titleRange.Font.Size = 16
    $titleRange.HorizontalAlignment = -4108
    $titleRange.VerticalAlignment = -4108
    $kp.Rows.Item(1).RowHeight = 30
    $noteRange = $kp.Range('A2:G2')
    $noteRange.Interior.Color = 0x00F7F4EF
    $noteRange.Font.Color = 0x006B7280
    $noteRange.Font.Italic = $true
    $noteRange.WrapText = $true
    $noteRange.HorizontalAlignment = -4131
    $kp.Rows.Item(2).RowHeight = 24
    $headerRange = $kp.Range("A3:G3")
    $headerRange.Interior.Color = 0x00983F7B
    $headerRange.Font.Color = 0x00FFFFFF
    $headerRange.Font.Bold = $true
    $headerRange.WrapText = $true
    $headerRange.HorizontalAlignment = -4108
    $headerRange.VerticalAlignment = -4108
    $kp.Rows.Item(3).RowHeight = 38
    $bodyRange = $kp.Range("A4:G$lastKpRow")
    $bodyRange.WrapText = $true
    $bodyRange.VerticalAlignment = -4160
    $bodyRange.Font.Name = 'Aptos'
    $bodyRange.Font.Size = 10
    $bodyRange.Borders.LineStyle = 1
    $bodyRange.Borders.Color = 0x00E5E7EB
    $kp.Range("E4:F$lastKpRow").NumberFormatLocal = 'ДД.ММ.ГГГГ'
    $kp.Range("G4:G$lastKpRow").NumberFormatLocal = '# ##0,00'
    $kp.Columns.Item('A').ColumnWidth = 31
    $kp.Columns.Item('B').ColumnWidth = 42
    $kp.Columns.Item('C').ColumnWidth = 11
    $kp.Columns.Item('D').ColumnWidth = 56
    $kp.Columns.Item('E').ColumnWidth = 15
    $kp.Columns.Item('F').ColumnWidth = 15
    $kp.Columns.Item('G').ColumnWidth = 23
    $kp.Rows("4:$lastKpRow").EntireRow.AutoFit() | Out-Null
    $kp.Range("A3:G$lastKpRow").AutoFilter() | Out-Null
    $kp.Activate()
    $excel.ActiveWindow.FreezePanes = $false
    $excel.ActiveWindow.SplitRow = 3
    $excel.ActiveWindow.FreezePanes = $true
    $kp.Tab.Color = 0x00983F7B

    # Импорт КП завершается без изменения существующей диаграммы: её привязки
    # к работам СР требуют отдельного смыслового сопоставления этапов.
    $workbook.Save()
    Write-Output "Импорт выполнен: $(@($preview.stages).Count) этапов КП внесены в $master"
    return

    # На этапе импорта нет достоверного соответствия между работами СР и этапами нового КП.
    # Снимаем старую демонстрационную привязку, чтобы не показывать ложный срок.
    $sr.Range('C2:C7').ClearContents()
    $sr.Range('C2:C7').Interior.Color = 0x009BEBFF
    for ($srRow = 2; $srRow -le 7; $srRow++) {
        $sr.Cells.Item($srRow, 16).Formula = "=IF(C$srRow=`"`",`"Не задан этап КП`",IF(L$srRow=`"Нет данных`",`"Неопределённость: нет даты ИД`",IF(O$srRow=`"`",`"Невозможно прогнозировать`",IF(O$srRow>VLOOKUP(C$srRow,'01_КП'!`$C`$4:`$F`$$lastKpRow,4,FALSE),`"Критично: позднее КП`",`"Норма`"))))"
    }

    # 05_График: строки КП с реальными датами и все текущие сроки СР на одной диаграмме.
    $graph.Range('A1:F200').ClearContents()
    $graphHeaders = @('Тип', 'Наименование', 'Начало', 'Длительность', 'КП — срок', 'СР — срок')
    for ($column = 1; $column -le $graphHeaders.Count; $column++) { $graph.Cells.Item(1, $column).Value = $graphHeaders[$column - 1] }
    $graphRow = 2
    foreach ($entry in $graphRows) {
        $graph.Cells.Item($graphRow, 1).Value = 'КП'
        $graph.Cells.Item($graphRow, 2).Formula = "='01_КП'!D$($entry.SourceRow)"
        $graph.Cells.Item($graphRow, 3).Formula = "='01_КП'!E$($entry.SourceRow)"
        $graph.Cells.Item($graphRow, 4).Formula = "='01_КП'!F$($entry.SourceRow)-C$graphRow+1"
        $graph.Cells.Item($graphRow, 5).Formula = "=IF(A$graphRow=`"КП`",D$graphRow,NA())"
        $graph.Cells.Item($graphRow, 6).Formula = '=NA()'
        $graphRow++
    }
    for ($srRow = 2; $srRow -le 7; $srRow++) {
        $graph.Cells.Item($graphRow, 1).Value = 'СР'
        $graph.Cells.Item($graphRow, 2).Formula = "='03_СР_проекта'!B$srRow"
        $graph.Cells.Item($graphRow, 3).Formula = "='03_СР_проекта'!N$srRow"
        $graph.Cells.Item($graphRow, 4).Formula = "=IF(OR(C$graphRow=`"`",'03_СР_проекта'!O$srRow=`"`"),NA(),'03_СР_проекта'!O$srRow-C$graphRow+1)"
        $graph.Cells.Item($graphRow, 5).Formula = '=NA()'
        $graph.Cells.Item($graphRow, 6).Formula = "=IF(A$graphRow=`"СР`",D$graphRow,NA())"
        $graphRow++
    }
    $lastGraphRow = $graphRow - 1
    $graph.Range("A1:F$lastGraphRow").Font.Name = 'Aptos'
    $graph.Range('A1:F1').Interior.Color = 0x00983F7B
    $graph.Range('A1:F1').Font.Color = 0x00FFFFFF
    $graph.Range('A1:F1').Font.Bold = $true
    $graph.Range("A1:F$lastGraphRow").WrapText = $true
    $graph.Columns.Item('A').ColumnWidth = 10
    $graph.Columns.Item('B').ColumnWidth = 52
    $graph.Columns.Item('C').ColumnWidth = 15
    $graph.Columns.Item('D').ColumnWidth = 14
    $graph.Columns.Item('E').ColumnWidth = 14
    $graph.Columns.Item('F').ColumnWidth = 14
    $graph.Range("C2:C$lastGraphRow").NumberFormat = 'dd.mm.yyyy'
    $graph.Range("A1:F$lastGraphRow").Borders.LineStyle = 1
    $graph.Range("A1:F$lastGraphRow").Borders.Color = 0x00E5E7EB
    $graph.Rows("2:$lastGraphRow").EntireRow.AutoFit() | Out-Null

    if ($graph.ChartObjects().Count -gt 0) {
        $chartObject = $graph.ChartObjects().Item(1)
        $chart = $chartObject.Chart
        $sheetRef = "'05_График'!"
        for ($seriesNo = 1; $seriesNo -le 3; $seriesNo++) {
            $series = $chart.SeriesCollection($seriesNo)
            $series.XValues = "=$sheetRef`$B`$2:`$B`$$lastGraphRow"
        }
        $chart.SeriesCollection(1).Values = "=$sheetRef`$C`$2:`$C`$$lastGraphRow"
        $chart.SeriesCollection(2).Values = "=$sheetRef`$E`$2:`$E`$$lastGraphRow"
        $chart.SeriesCollection(3).Values = "=$sheetRef`$F`$2:`$F`$$lastGraphRow"
        $chartObject.Height = [Math]::Max(590, 150 + 21 * ($lastGraphRow - 1))
        $chartObject.Width = 1280
    }

    $dashboardSheet.Range('A1').Value = 'GIP Navigator — этап 0'
    $dashboardSheet.Range('A2').Value = "Проект: $($preview.project)"
    $dashboardSheet.Range('A3').Value = "Импортировано КП: $(@($preview.stages).Count) этапов. Для срокового контроля назначьте этап КП в жёлтых ячейках листа 03_СР_проекта."
    $dashboardSheet.Range('A3').WrapText = $true

    $excel.CalculateFullRebuild()
    $workbook.Save()
    Write-Output "Импорт выполнен: $(@($preview.stages).Count) этапов КП внесены в $master"
}
finally {
    if ($workbook) {
        $workbook.Close($false)
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
    }
    if ($excel) {
        $excel.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
