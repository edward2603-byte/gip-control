[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$ProjectFilePath,

    [string]$TemplatePath,

    [switch]$ReplaceExisting
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($TemplatePath)) {
    $TemplatePath = Join-Path $PSScriptRoot '..\templates\GIP_Navigator_TEMPLATE.xlsx'
}

if (-not (Test-Path -LiteralPath $TemplatePath -PathType Leaf)) {
    throw "Шаблон не найден: $TemplatePath"
}

$template = (Resolve-Path -LiteralPath $TemplatePath).Path
$target = [IO.Path]::GetFullPath($ProjectFilePath)
if ([IO.Path]::GetFullPath($template) -eq $target) {
    throw 'Нельзя создавать проектный файл поверх мастер-шаблона.'
}
if ([IO.Path]::GetExtension($target) -ne '.xlsx') {
    throw 'Файл проекта должен иметь расширение .xlsx.'
}

$targetDirectory = Split-Path -Parent $target
if (-not (Test-Path -LiteralPath $targetDirectory -PathType Container)) {
    throw "Папка проекта не найдена: $targetDirectory"
}
if (Test-Path -LiteralPath $target -PathType Leaf) {
    if (-not $ReplaceExisting) {
        throw "Файл уже существует: $target. Для явной замены укажите -ReplaceExisting."
    }
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $backup = "$target.backup_$stamp.xlsx"
    Copy-Item -LiteralPath $target -Destination $backup -ErrorAction Stop
}

Copy-Item -LiteralPath $template -Destination $target -Force

$excel = $null
$book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Open($target)
    $service = $book.Worksheets.Item('99_Служебный')
    $service.Range('B4').Value2 = [guid]::NewGuid().ToString()
    $service.Range('B5').Value = Get-Date
    $book.Save()
}
finally {
    # Save is called explicitly above. Closing without a second Save avoids a
    # duplicate COM save pass that can leave a damaged workbook on some Excel builds.
    if ($book) { $book.Close($false); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($book) }
    if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

Write-Output "Создана отдельная книга проекта: $target"
