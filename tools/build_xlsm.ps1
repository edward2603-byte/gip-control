param(
    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Off
$repoRoot = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($OutputPath)) { $OutputPath = Join-Path $repoRoot 'templates\GIP_Navigator_TEMPLATE.xlsm' }
$OutputPath = [IO.Path]::GetFullPath($OutputPath)
# The XLSX is only a temporary construction layer.  The deliverable used by a
# project manager is the single macro-enabled XLSM file, so never overwrite a
# workbook in templates while building it.
$basePath = Join-Path ([IO.Path]::GetTempPath()) ("GIP_Navigator_base_{0}.xlsx" -f [guid]::NewGuid().ToString('N'))
$modulePath = Join-Path $repoRoot 'vba\GIPNavigator.bas'
$workbookClassPath = Join-Path $repoRoot 'vba\ThisWorkbook.cls'

& py (Join-Path $repoRoot 'tools\build_template.py') --output $basePath
if ($LASTEXITCODE -ne 0) { throw 'Не удалось собрать базовую книгу XLSX.' }

$excel = $null
$book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Open($basePath)
    $component = $book.VBProject.VBComponents.Add(1)
    $component.Name = 'GIPNavigator'
    $moduleText = [IO.File]::ReadAllText($modulePath, [Text.UTF8Encoding]::new($false))
    $component.CodeModule.AddFromString($moduleText)
    $workbookClassText = [IO.File]::ReadAllText($workbookClassPath, [Text.UTF8Encoding]::new($false))
    # Workbook code-name is localized in Russian Excel (for example, "ЭтаКнига").
    $book.VBProject.VBComponents.Item($book.CodeName).CodeModule.AddFromString($workbookClassText)

    $homeSheet = $book.Worksheets.Item('00_Главная')
    $homeSheet.Range('A15:P17').Value2 = '1. Скопируйте этот файл в папку проекта и переименуйте его.' + [Environment]::NewLine + '2. Откройте книгу в Microsoft Excel, нажмите «Импортировать КП» и укажите дату подписания наряд-заказа на листе 01.' + [Environment]::NewLine + '3. После импорта выберите СР, задайте связи FS/SS на листе 04 и при необходимости нажмите «Пересчитать СР».'
    $book.SaveAs($OutputPath, 52)

    $actions = @(
        @('GIP_Action_Import', 'Импортировать КП', 'ImportKp', 24, 190, 170, 34, 15111910),
        @('GIP_Action_SR', 'Перейти к СР', 'GoToSR', 208, 190, 150, 34, 14087411),
        @('GIP_Action_AddSR', 'Добавить экземпляр СР', 'AddSRInstance', 372, 190, 185, 34, 10125999),
        @('GIP_Action_Recalc', 'Пересчитать СР', 'RecalculateSRNetwork', 571, 190, 165, 34, 5304534),
        @('GIP_Action_Graph', 'Открыть график', 'GoToGraph', 750, 190, 150, 34, 10125999),
        @('GIP_Action_Control', 'Открыть контроль', 'GoToControl', 914, 190, 160, 34, 5304534),
        @('GIP_Action_Approve', 'Зафиксировать согласование', 'ApproveBaseline', 1088, 190, 220, 34, 10125999)
    )
    foreach ($a in $actions) {
        $shape = $homeSheet.Shapes.AddShape(5, $a[3], $a[4], $a[5], $a[6])
        $shape.Name = $a[0]
        $shape.TextFrame2.TextRange.Text = $a[1]
        $shape.TextFrame2.TextRange.Font.Size = 11
        $shape.TextFrame2.TextRange.Font.Bold = $true
        $shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 16777215
        $shape.Fill.ForeColor.RGB = $a[7]
        $shape.Line.Visible = $false
        # Do not qualify the macro with the template filename: the user renames
        # the copied workbook before opening it.
        $shape.OnAction = $a[2]
    }
    $excel.Run("'$($book.Name)'!RefreshGantt")
    $librarySheet = $book.Worksheets.Item('02_Библиотека_СР')
    $librarySheet.Range('A5:A85').NumberFormat = ';;;'
    for ($row = 5; $row -le 85; $row++) {
        $cell = $librarySheet.Cells.Item($row, 1)
        $cell.Value2 = $true
        $checkBox = $librarySheet.CheckBoxes().Add($cell.Left + 22, $cell.Top + 5, 16, 16)
        $checkBox.Name = "GIP_SR_Select_$row"
        $checkBox.Caption = ''
        $checkBox.LinkedCell = "'02_Библиотека_СР'!`$A`$$row"
        $checkBox.Value = 1
        $checkBox.OnAction = 'SyncSelectedSR'
        $checkBox.Placement = 1
    }
    $book.Save()
    Write-Output $OutputPath
}
finally {
    if ($book) { $book.Close($false) }
    if ($excel) { $excel.Quit() }
    foreach ($item in @($component, $book, $excel)) { if ($null -ne $item -and [Runtime.InteropServices.Marshal]::IsComObject($item)) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($item) } }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if (Test-Path -LiteralPath $basePath) { Remove-Item -LiteralPath $basePath -Force }
}
