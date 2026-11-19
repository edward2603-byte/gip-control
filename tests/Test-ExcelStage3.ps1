[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$testDir = Join-Path ([IO.Path]::GetTempPath()) ('gip-navigator-stage3-' + [guid]::NewGuid().ToString('N'))
$bookPath = Join-Path $testDir 'Stage3.xlsm'

function Assert-True([bool]$Condition, [string]$Message) { if(-not $Condition){throw $Message} }
function Release-ComObject($Object) { if($null -ne $Object -and [Runtime.InteropServices.Marshal]::IsComObject($Object)){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($Object)} }
function Count-NonBlank($Range) {
    $count=0;$values=$Range.Value2
    if($values -is [array]) {
        foreach($row in $values){foreach($v in $row){if($null -ne $v -and "$v".Trim().Length -gt 0){$count++}}}
    } elseif($null -ne $values -and "$values".Trim().Length -gt 0) {$count++}
    return $count
}
function Find-CodeRow($Sheet,[string]$Code) { for($r=5;$r -le 304;$r++){if($Sheet.Cells.Item($r,2).Text -eq $Code){return $r}}; return 0 }

New-Item -ItemType Directory -Path $testDir | Out-Null
$excel=$null;$book=$null
try {
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $repo 'tools\build_xlsm.ps1') -OutputPath $bookPath
    if($LASTEXITCODE -ne 0 -or -not(Test-Path -LiteralPath $bookPath)){throw 'Could not build the Stage 3 XLSM.'}
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AutomationSecurity=1
    $book=$excel.Workbooks.Open($bookPath)
    $library=$book.Worksheets.Item(3);$kp=$book.Worksheets.Item(2);$sr=$book.Worksheets.Item(4)
    Assert-True ($library.CheckBoxes().Count -eq 81) 'Library must have 81 checkboxes.'
    $selected=0;for($r=5;$r -le 85;$r++){if([bool]$library.Cells.Item($r,1).Value2){$selected++}}
    Assert-True ($selected -eq 81) 'All 81 standard SR rows are not selected by default.'
    Assert-True ((Count-NonBlank $sr.Range('A5:A304')) -eq 81) 'Selected SR rows were not synchronized on opening.'

    $excel.EnableEvents=$false
    $kp.Range('C5').Value2='1';$kp.Range('J5').Value=[datetime]'2027-03-01';$kp.Range('K5').Value=[datetime]'2027-06-01'
    $excel.EnableEvents=$true
    $excel.Run("'$($book.Name)'!RecalculateSRNetwork", $true)
    $chart=$library.ChartObjects('GIP_Library_Gantt').Chart
    $beforeChartRows=Count-NonBlank $library.Range('M5:M304')
    Assert-True ($beforeChartRows -gt 0 -and $chart.SeriesCollection().Count -eq 2) 'Library Gantt was not built from calculated SR dates.'

    # A row with facts cannot be silently removed: selection is restored.
    $protectedCode=[string]$library.Range('B15').Value2;$protectedRow=Find-CodeRow $sr $protectedCode
    $sr.Cells.Item($protectedRow,7).Value=[datetime]'2027-03-05'
    $library.Range('A15').Value2=$false
    $excel.Run("'$($book.Name)'!SyncSelectedSR", $true)
    Assert-True ([bool]$library.Range('A15').Value2 -and (Find-CodeRow $sr $protectedCode) -gt 0) 'A selected SR with facts was removed without conflict protection.'

    # Manual duration and an existing FS link are also user data and must not
    # disappear when the corresponding checkboxes are cleared.
    $manualCode=[string]$library.Range('B16').Value2;$manualRow=Find-CodeRow $sr $manualCode
    $sr.Cells.Item($manualRow,5).Value2=42
    $library.Range('A16').Value2=$false
    $excel.Run("'$($book.Name)'!SyncSelectedSR", $true)
    Assert-True ([bool]$library.Range('A16').Value2 -and [int]$sr.Cells.Item($manualRow,5).Value2 -eq 42) 'A manual SR duration was not preserved.'
    $linkedCode=[string]$library.Range('B17').Value2;$linkedRow=Find-CodeRow $sr $linkedCode
    $predecessorCode=[string]$library.Range('B18').Value2;$predecessorRow=Find-CodeRow $sr $predecessorCode
    $links=$book.Worksheets.Item(5)
    $excel.EnableEvents=$false
    $links.Cells.Item(5,2).Value2=$sr.Cells.Item($linkedRow,1).Value2
    $links.Cells.Item(5,3).Value2=$sr.Cells.Item($predecessorRow,1).Value2
    $links.Cells.Item(5,4).Value2='FS'
    $library.Range('A17').Value2=$false
    $excel.EnableEvents=$true
    $excel.Run("'$($book.Name)'!SyncSelectedSR", $true)
    Assert-True ([bool]$library.Range('A17').Value2 -and "$($links.Cells.Item(5,2).Value2)".Trim().Length -gt 0) 'An SR with an existing dependency was removed without conflict protection.'

    for($r=5;$r -le 14;$r++){$library.Cells.Item($r,1).Value2=$false}
    $excel.Run("'$($book.Name)'!SyncSelectedSR", $true)
    $afterRemoval=Count-NonBlank $sr.Range('A5:A304');$afterChartRows=Count-NonBlank $library.Range('M5:M304')
    Assert-True ($afterRemoval -eq 71 -and $afterChartRows -lt $beforeChartRows) 'Removing ten checkboxes did not remove SRs from the active project and library Gantt.'
    $library.Range('A5').Value2=$true
    $excel.Run("'$($book.Name)'!SyncSelectedSR", $true)
    Assert-True ((Count-NonBlank $sr.Range('A5:A304')) -eq 72 -and (Find-CodeRow $sr ([string]$library.Range('B5').Value2)) -gt 0) 'Restoring a checkbox did not restore exactly one SR.'
    $code=[string]$library.Range('B5').Value2
    $excel.Run("'$($book.Name)'!AddSRInstanceByCode", $code, $true)
    $instances=0;for($r=5;$r -le 304;$r++){if($sr.Cells.Item($r,2).Text -eq $code){$instances++}}
    Assert-True ($instances -eq 2) 'Second instance of an SR code was not retained separately.'
    $unknown=0;for($r=5;$r -le 304;$r++){if("$($sr.Cells.Item($r,1).Value2)".Trim().Length -gt 0 -and "$($sr.Cells.Item($r,5).Value2)".Trim().Length -eq 0){if("$($sr.Cells.Item($r,13).Value2)".Trim().Length -gt 0){$unknown++}}}
    Assert-True ($unknown -gt 0) 'Unknown library duration was converted to a fictitious value instead of being diagnosed.'
    $book.Save();$book.Close($true);Release-ComObject $book;$book=$excel.Workbooks.Open($bookPath)
    $library=$book.Worksheets.Item(3);$sr=$book.Worksheets.Item(4)
    $selectedAfterOpen=0;for($r=5;$r -le 85;$r++){if([bool]$library.Cells.Item($r,1).Value2){$selectedAfterOpen++}}
    Assert-True ($selectedAfterOpen -eq 72 -and (Count-NonBlank $sr.Range('A5:A304')) -eq 73) 'Checkbox choices or separate SR instance were not preserved after reopen.'
    Write-Output 'PASS: Stage 3 library selection, dynamic Gantt, protected facts, duplicate instance, unknown duration, and reopen passed in Microsoft Excel.'
}
catch {Write-Output ('STAGE3_ERROR_LINE='+$_.InvocationInfo.ScriptLineNumber);Write-Output ('STAGE3_ERROR='+$_.Exception.Message);throw}
finally {if($excel){$excel.EnableEvents=$true};if($book){$book.Close($false);Release-ComObject $book};if($excel){$excel.Quit();Release-ComObject $excel};[GC]::Collect();[GC]::WaitForPendingFinalizers();if(Test-Path -LiteralPath $testDir){Remove-Item -LiteralPath $testDir -Recurse -Force}}
