[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repo = Split-Path -Parent $PSScriptRoot
$template = Join-Path $repo 'templates\GIP_Navigator_TEMPLATE.xlsx'
$builder = Join-Path $repo 'tools\build_template.py'
$creator = Join-Path $repo 'tools\New-ProjectFromTemplate.ps1'

& python $builder --output $template
if ($LASTEXITCODE -ne 0) { throw 'Не удалось собрать мастер-шаблон.' }
if (-not (Test-Path -LiteralPath $template -PathType Leaf)) { throw 'Мастер-шаблон не создан.' }

$requiredSheets = @('00_Главная','01_КП','02_Библиотека_СР','03_СР_проекта','04_Согласование_плана','05_График','06_Контроль','99_Служебный')
$testDir = Join-Path ([IO.Path]::GetTempPath()) ("gip-navigator-template-" + [guid]::NewGuid())
New-Item -ItemType Directory -Path $testDir | Out-Null
try {
    $first = Join-Path $testDir 'Project_A.xlsx'
    $second = Join-Path $testDir 'Project_B.xlsx'
    & $creator -ProjectFilePath $first
    & $creator -ProjectFilePath $second
    if ($LASTEXITCODE -ne 0) { throw 'Не удалось создать книги проектов.' }

    $excel = $null; $templateBook = $null; $firstBook = $null; $secondBook = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false; $excel.DisplayAlerts = $false
        $templateBook = $excel.Workbooks.Open((Resolve-Path -LiteralPath $template).Path)
        $sheetNames = @($templateBook.Worksheets | ForEach-Object Name)
        foreach ($sheet in $requiredSheets) { if ($sheetNames -notcontains $sheet) { throw "Отсутствует лист: $sheet" } }
        if ($templateBook.Worksheets.Item('01_КП').Cells.Item(5,3).Value2) { throw 'В мастер-шаблоне есть этап КП.' }
        if ($templateBook.Worksheets.Item('03_СР_проекта').Cells.Item(5,2).Value2) { throw 'В мастер-шаблоне есть демонстрационная СР.' }
        if ($templateBook.Worksheets.Item('02_Библиотека_СР').Range('B5').Text -ne 'SR-001' -or $templateBook.Worksheets.Item('02_Библиотека_СР').Range('B85').Text -ne 'SR-081') { throw 'Библиотека СР не содержит 81 типовую работу.' }
        if ($templateBook.Worksheets.Item('01_КП').Range('H4').Text -ne 'Дата фактического актирования' -or $templateBook.Worksheets.Item('01_КП').Range('J5').Formula -notmatch 'G5') { throw 'В 01_КП не подготовлены поля H:K для актирования.' }
        $templateBook.Close($false); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($templateBook); $templateBook = $null
        $firstBook = $excel.Workbooks.Open((Resolve-Path -LiteralPath $first).Path)
        $firstId = [string]$firstBook.Worksheets.Item('99_Служебный').Range('B4').Value2
        $kp75 = $firstBook.Worksheets.Item('01_КП')
        for ($row = 5; $row -le 79; $row++) {
            $index = $row - 4
            $start = [datetime]::new(2026, 1, 1).AddDays(($index - 1) * 2)
            $kp75.Cells.Item($row, 1).Value2 = 'Договор-ТЕСТ'
            $kp75.Cells.Item($row, 2).Value2 = 'Тестовый проект 75'
            $kp75.Cells.Item($row, 3).Value2 = "Э-$index"
            $kp75.Cells.Item($row, 4).Value2 = "Этап $index"
            $kp75.Cells.Item($row, 5).Value = $start
            $kp75.Cells.Item($row, 6).Value = $start.AddDays(1)
            $kp75.Cells.Item($row, 7).Value2 = [double]($index * 1000)
        }
        $sr75 = $firstBook.Worksheets.Item('03_СР_проекта')
        for ($index = 1; $index -le 50; $index++) { $sr75.Cells.Item($index + 4, 2).Value2 = ('SR-{0:D3}' -f $index) }
        $sr75.Cells.Item(55, 2).Value2 = 'CUSTOM-001'
        $sr75.Cells.Item(55, 3).Value2 = 'Собственная СР'
        $sr75.Cells.Item(55, 5).Value2 = [double]5
        $excel.CalculateFullRebuild()
        if ([int]$firstBook.Worksheets.Item('00_Главная').Range('E5').Value2 -ne 75) { throw 'Сводка не обработала 75 этапов КП.' }
        if ([decimal]$firstBook.Worksheets.Item('00_Главная').Range('I5').Value2 -ne 2850000) { throw 'Сумма 75 этапов рассчитана неверно.' }
        if ($firstBook.Worksheets.Item('03_СР_проекта').Columns('M').Hidden -ne $true) { throw 'Техническая проверка сети СР должна быть скрыта.' }
        if ($firstBook.Worksheets.Item('04_Согласование_плана').Range('D5').Validation.Type -ne 3) { throw 'Нет списка FS/SS для связей СР.' }
        if ($firstBook.Worksheets.Item('05_График').ChartObjects().Count -ne 1) { throw 'Нет диаграммы Ганта КП.' }
        if ([string]::IsNullOrWhiteSpace($sr75.Range('C5').Text) -or $sr75.Range('C55').Text -ne 'Собственная СР') { throw 'Выбор библиотечной или собственной СР не работает.' }
        if ($sr75.Range('C304').Formula -notmatch 'VLOOKUP') { throw 'Формула выбора СР не распространена на 300-ю строку.' }
        $firstBook.Save(); $firstBook.Close($false); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($firstBook); $firstBook = $null
        $secondBook = $excel.Workbooks.Open((Resolve-Path -LiteralPath $second).Path)
        $secondId = [string]$secondBook.Worksheets.Item('99_Служебный').Range('B4').Value2
        $kp2 = $secondBook.Worksheets.Item('01_КП')
        for ($row = 5; $row -le 6; $row++) {
            $index = $row - 4
            $kp2.Cells.Item($row, 1).Value2 = 'Договор-ТЕСТ'
            $kp2.Cells.Item($row, 2).Value2 = 'Тестовый проект 2'
            $kp2.Cells.Item($row, 3).Value2 = "T-$index"
            $kp2.Cells.Item($row, 4).Value2 = "Этап $index"
            $kp2.Cells.Item($row, 5).Value = [datetime]::new(2026, 1, $index)
            $kp2.Cells.Item($row, 6).Value = [datetime]::new(2026, 1, $index + 1)
            $kp2.Cells.Item($row, 7).Value2 = [double]1000
        }
        $excel.CalculateFullRebuild()
        if ([int]$secondBook.Worksheets.Item('00_Главная').Range('E5').Value2 -ne 2) { throw 'Сводка не обработала 2 этапа КП.' }
        $graph2 = $secondBook.Worksheets.Item('05_График')
        if ($graph2.Range('B5').Text -ne 'Этап 1' -or $graph2.Range('B6').Text -ne 'Этап 2' -or $graph2.Range('B7').Text -ne '') { throw 'График содержит неактивные этапы КП.' }
        if ([string]::IsNullOrWhiteSpace($firstId) -or $firstId -eq $secondId) { throw 'Для отдельных книг не созданы разные workbook_id.' }
        $secondBook.Save()
    }
    finally {
        if ($firstBook) { $firstBook.Close($false); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($firstBook) }
        if ($secondBook) { $secondBook.Close($false); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($secondBook) }
        if ($templateBook) { $templateBook.Close($false); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($templateBook) }
        if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [IO.Compression.ZipFile]::OpenRead($template)
    try {
        $xml = foreach ($entry in $archive.Entries | Where-Object { $_.FullName -match 'sharedStrings|worksheets|core.xml' }) {
            $reader = [IO.StreamReader]::new($entry.Open())
            try { $reader.ReadToEnd() } finally { $reader.Dispose() }
        }
        if (($xml -join "`n") -match 'Стрижов Луг|Лугового месторождения') { throw 'В мастер-шаблоне обнаружены данные примера проекта.' }
    }
    finally { $archive.Dispose() }
    Write-Output 'PASS: мастер-шаблон пуст, содержит обязательные листы и создаёт независимые книги проекта.'
}
finally {
    if (Test-Path -LiteralPath $testDir -PathType Container) { Remove-Item -LiteralPath $testDir -Recurse -Force }
}
