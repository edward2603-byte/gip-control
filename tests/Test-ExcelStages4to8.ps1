[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$testDir = Join-Path ([IO.Path]::GetTempPath()) ('gip-navigator-stage4to8-' + [guid]::NewGuid().ToString('N'))
$bookPath = Join-Path $testDir 'Stages4to8.xlsm'

function Assert-True([bool]$Condition, [string]$Message) { if(-not $Condition){throw $Message} }
function Release-ComObject($Object) { if($null -ne $Object -and [Runtime.InteropServices.Marshal]::IsComObject($Object)){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($Object)} }

New-Item -ItemType Directory -Path $testDir | Out-Null
$excel=$null;$book=$null
try {
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $repo 'tools\build_xlsm.ps1') -OutputPath $bookPath
    if($LASTEXITCODE -ne 0 -or -not(Test-Path -LiteralPath $bookPath)){throw 'Could not build the XLSM for stages 4-8.'}
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AutomationSecurity=1
    $book=$excel.Workbooks.Open($bookPath)
    $kp=$book.Worksheets.Item('01_КП');$sr=$book.Worksheets.Item('03_СР_проекта');$agreement=$book.Worksheets.Item('04_Согласование_плана');$graph=$book.Worksheets.Item('05_График');$control=$book.Worksheets.Item('06_Контроль');$acts=$book.Worksheets.Item('07_Актирование')
    $excel.EnableEvents=$false
    $kp.Range('B3').Value=[datetime]'2027-03-01'
    $kp.Range('C5').Value2='1';$kp.Range('D5').Value2='Этап 1';$kp.Range('E5').Value=[datetime]'2027-03-10';$kp.Range('F5').Value=[datetime]'2027-03-20';$kp.Range('G5').Value2=1000;$kp.Range('J5').Value=[datetime]'2027-03-10';$kp.Range('K5').Value=[datetime]'2027-03-20'
    $kp.Range('C6').Value2='2';$kp.Range('D6').Value2='Этап 2';$kp.Range('E6').Value=[datetime]'2027-03-25';$kp.Range('F6').Value=[datetime]'2027-03-30';$kp.Range('G6').Value2=2000;$kp.Range('J6').Value=[datetime]'2027-03-25';$kp.Range('K6').Value=[datetime]'2027-03-30'
    $sr.Range('I5').Value=[datetime]'2027-03-15'
    $agreement.Range('I5').Value2=$sr.Range('A5').Value2;$agreement.Range('K5').Value2='1';$agreement.Range('M5').Value2='До начала этапа'
    $agreement.Range('I6').Value2=$sr.Range('A5').Value2;$agreement.Range('K6').Value2='2';$agreement.Range('M6').Value2='До окончания этапа'
    $excel.EnableEvents=$true
    $excel.Run("'$($book.Name)'!RecalculateSRNetwork", $true)
    Write-Output 'CHECK: obligation start delay'
    Assert-True ([double]$kp.Range('N5').Value2 -eq ([datetime]'2027-03-16').ToOADate()) 'An SR did not delay the KP stage start to the next calendar day.'
    Assert-True ([double]$kp.Range('O5').Value2 -eq ([datetime]'2027-03-26').ToOADate()) 'The delayed KP stage did not keep its minimum duration.'
    Assert-True ($kp.Range('Q5').Text -eq 'Критично' -and $agreement.Range('N5').Text -eq 'Критично') 'Critical SR-to-KP start dependency was not diagnosed.'
    Assert-True ($agreement.Range('N6').Text -in @('Нормально','Предупреждение')) 'One SR could not be assigned to a second KP stage.'

    # A required SR without a calculable end must be insufficient, never green.
    $unknownRow=0;for($r=5;$r -le 304;$r++){if("$($sr.Cells.Item($r,5).Value2)".Trim().Length -eq 0 -and "$($sr.Cells.Item($r,1).Value2)".Trim().Length -gt 0){$unknownRow=$r;break}}
    Assert-True ($unknownRow -gt 0) 'No unknown-duration SR was available for the test.'
    $agreement.Range('I7').Value2=$sr.Cells.Item($unknownRow,1).Value2;$agreement.Range('K7').Value2='2';$agreement.Range('M7').Value2='До окончания этапа'
    $excel.Run("'$($book.Name)'!RecalculateSRNetwork", $true)
    Write-Output 'CHECK: unknown required SR'
    Assert-True ($kp.Range('Q6').Text -eq 'Недостаточно данных' -and $agreement.Range('N7').Text -eq 'Недостаточно данных') 'Unknown required SR was incorrectly treated as normal.'

    $excel.EnableEvents=$false
    $agreement.Range('I7').Value2='';$agreement.Range('K7').Value2='';$agreement.Range('M7').Value2=''
    $excel.EnableEvents=$true
    $excel.Run("'$($book.Name)'!RecalculateSRNetwork", $true)
    Write-Output ('CHECK: cleared unknown;Q5=' + $kp.Range('Q5').Value2 + ';Q6=' + $kp.Range('Q6').Value2 + ';RULE5=' + $agreement.Range('N5').Value2 + ';RULE6=' + $agreement.Range('N6').Value2 + ';RULE7=' + $agreement.Range('N7').Value2 + ';C5=' + $control.Range('C5').Value2 + ';C6=' + $control.Range('C6').Value2)
    $agreement.Range('R3').Value2='1.0';$agreement.Range('R4').Value=[datetime]'2027-03-02';$agreement.Range('R5').Value2='Тестовый руководитель'
    $excel.Run("'$($book.Name)'!ApproveBaseline", $true)
    Write-Output ('CHECK: approval snapshot;Q10=' + $agreement.Range('Q10').Value2 + ';R6=' + $agreement.Range('R6').Value2 + ';DIAG=' + $book.Worksheets.Item(8).Range('B20').Value2)
    Assert-True ($agreement.Range('Q10').Text -eq '1' -and $agreement.Range('S10').Text -eq '1') 'Approved baseline snapshot was not stored.'

    $excel.EnableEvents=$false
    $acts.Range('A5').Value2='ACT-001';$acts.Range('B5').Value2='1';$acts.Range('D5').Value=[datetime]'2027-03-17';$acts.Range('E5').Value2='1';$acts.Range('F5').Value2=400;$acts.Range('G5').Value2='Подписан'
    $acts.Range('A6').Value2='ACT-002';$acts.Range('B6').Value2='1';$acts.Range('D6').Value=[datetime]'2027-03-18';$acts.Range('E6').Value2='2';$acts.Range('F6').Value2=600;$acts.Range('G6').Value2='Подписан'
    $acts.Range('A7').Value2='ACT-003';$acts.Range('B7').Value2='1';$acts.Range('D7').Value=[datetime]'2027-03-19';$acts.Range('E7').Value2='3';$acts.Range('F7').Value2=999;$acts.Range('G7').Value2='Черновик'
    $excel.EnableEvents=$true
    $excel.Run("'$($book.Name)'!RecalculateActs")
    Write-Output 'CHECK: partial acts'
    Assert-True ([double]$kp.Range('S5').Value2 -eq 1000 -and [double]$kp.Range('T5').Value2 -eq 0) 'Signed partial acts were not summed correctly or draft was included.'
    $acts.Range('E6').Value2='1'
    $excel.Run("'$($book.Name)'!RecalculateActs")
    Assert-True ($acts.Range('H6').Text -match 'Дубликат') 'Duplicate signed act was not diagnosed.'
    $acts.Range('E6').Value2='2'
    $excel.Run("'$($book.Name)'!RecalculateSRNetwork", $true)
    Write-Output 'CHECK: graph and reopen'
    $chart=$graph.ChartObjects(1).Chart
    Assert-True ($chart.SeriesCollection().Count -eq 6) 'Unified Gantt did not contain plan, forecast, and actual series.'
    $criticalSignals=0;for($r=5;$r -le 304;$r++){if($control.Cells.Item($r,3).Text -eq 'Критично'){$criticalSignals++}}
    Assert-True ($criticalSignals -gt 0) 'Control sheet does not include a critical KP signal.'
    $book.Save();$book.Close($true);Release-ComObject $book;$book=$excel.Workbooks.Open($bookPath)
    $kp=$book.Worksheets.Item('01_КП');$agreement=$book.Worksheets.Item('04_Согласование_плана');$acts=$book.Worksheets.Item('07_Актирование')
    Assert-True ($agreement.Range('Q10').Text -eq '1' -and [double]$kp.Range('S5').Value2 -eq 1000 -and $acts.Range('A5').Text -eq 'ACT-001') 'Approval snapshot or act register did not persist after reopen.'
    Write-Output 'PASS: Stages 4-8 dependencies, forecast, control, approval snapshot, acts, and unified Gantt passed in Microsoft Excel.'
}
catch {Write-Output ('STAGES4TO8_ERROR_LINE='+$_.InvocationInfo.ScriptLineNumber);Write-Output ('STAGES4TO8_ERROR='+$_.Exception.Message);throw}
finally {if($excel){$excel.EnableEvents=$true};if($book){$book.Close($false);Release-ComObject $book};if($excel){$excel.Quit();Release-ComObject $excel};[GC]::Collect();[GC]::WaitForPendingFinalizers();if(Test-Path -LiteralPath $testDir){Remove-Item -LiteralPath $testDir -Recurse -Force}}
