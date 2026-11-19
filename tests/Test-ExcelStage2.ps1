[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$testDir = Join-Path ([IO.Path]::GetTempPath()) ('gip-navigator-stage2-' + [guid]::NewGuid().ToString('N'))
$bookPath = Join-Path $testDir 'Stage2.xlsm'

function Assert-True([bool]$Condition, [string]$Message) { if (-not $Condition) { throw $Message } }
function Release-ComObject($Object) { if ($null -ne $Object -and [Runtime.InteropServices.Marshal]::IsComObject($Object)) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Object) } }

New-Item -ItemType Directory -Path $testDir | Out-Null
$excel = $null; $book = $null
try {
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $repo 'tools\build_xlsm.ps1') -OutputPath $bookPath
    if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $bookPath)) { throw 'Could not build the Stage 2 XLSM.' }
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.AutomationSecurity = 1
    $book = $excel.Workbooks.Open($bookPath)
    $kp = $book.Worksheets.Item(2); $sr = $book.Worksheets.Item(4); $links = $book.Worksheets.Item(5); $control = $book.Worksheets.Item(7)
    $excel.EnableEvents = $false
    $kp.Range('C5').Value2 = '1'; $kp.Range('J5').Value = [datetime]'2027-03-01'; $kp.Range('K5').Value = [datetime]'2027-03-10'
    $sr.Range('A5:O304').ClearContents(); $links.Range('A5:G304').ClearContents()
    $records = @(
        @('A','SR-001',10,[datetime]'2027-03-01'), @('B','SR-002',1,$null), @('C','SR-003',1,$null),
        @('D','SR-004',1,$null), @('E','SR-005',0,$null), @('F','SR-006',1,$null),
        @('G','SR-007',1,$null), @('H','SR-008',1,$null), @('P','SR-009',1,$null)
    )
    for ($i=0; $i -lt $records.Count; $i++) {
        $row=5+$i; $sr.Cells.Item($row,1).Value2=$records[$i][0]; $sr.Cells.Item($row,2).Value2=$records[$i][1]; $sr.Cells.Item($row,5).Value2=[string]$records[$i][2]
        if ($null -ne $records[$i][3]) { $sr.Cells.Item($row,15).Value=$records[$i][3] }
    }
    # C=FS after A, D=SS after A plus 2, E=milestone after A, F=two predecessors.
    $linkData=@(@('L1','C','A','FS',0),@('L2','D','A','SS',2),@('L3','E','A','FS',0),@('L4','F','A','FS',0),@('L5','F','B','SS',15),@('L6','G','H','FS',0),@('L7','H','G','FS',0),@('L8','P','MISSING','FS',0))
    for ($i=0; $i -lt $linkData.Count; $i++) { $row=5+$i; for($c=0;$c -lt 5;$c++){ $links.Cells.Item($row,$c+1).Value2=[string]$linkData[$i][$c] } }
    # Actual/forecast test cases: A late actual start with a late expected end;
    # B completed; D in progress without a forecast.
    $sr.Range('G5').Value=[datetime]'2027-03-05'; $sr.Range('I5').Value=[datetime]'2027-03-15'
    $sr.Range('G6').Value=[datetime]'2027-03-01'; $sr.Range('H6').Value=[datetime]'2027-03-04'
    $sr.Range('G8').Value=[datetime]'2027-03-03'
    $excel.EnableEvents = $true
    $excel.Run("'$($book.Name)'!RecalculateSRNetwork", $true)
    $mar1=([datetime]'2027-03-01').ToOADate(); $mar3=([datetime]'2027-03-03').ToOADate(); $mar10=([datetime]'2027-03-10').ToOADate(); $mar11=([datetime]'2027-03-11').ToOADate(); $mar15=([datetime]'2027-03-15').ToOADate(); $mar16=([datetime]'2027-03-16').ToOADate()
    Assert-True ([double]$sr.Range('D5').Value2 -eq $mar1 -and [double]$sr.Range('F5').Value2 -eq $mar10) 'Independent SR A is wrong.'
    Assert-True ([double]$sr.Range('D6').Value2 -eq $mar1) 'Parallel independent SR B inherited another date.'
    Assert-True ([double]$sr.Range('D7').Value2 -eq $mar11) 'FS schedule is wrong.'
    Assert-True ([double]$sr.Range('D8').Value2 -eq $mar3) 'SS schedule is wrong.'
    Assert-True ([double]$sr.Range('D9').Value2 -eq $mar11 -and [double]$sr.Range('F9').Value2 -eq $mar11) 'Zero-duration milestone is wrong.'
    Assert-True ([double]$sr.Range('D10').Value2 -eq $mar16) 'Multiple predecessors did not use the latest date.'
    Assert-True ([double]$sr.Range('J5').Value2 -eq $mar15 -and [double]$sr.Range('J6').Value2 -eq ([datetime]'2027-03-04').ToOADate()) 'Actual and expected forecasts are wrong.'
    Assert-True ($sr.Range('J8').Text.Length -eq 0 -and $sr.Range('K8').Text.Length -gt 0) 'Missing forecast was invented or status omitted.'
    Assert-True ($sr.Range('M11').Text.Length -gt 0 -and $sr.Range('M12').Text.Length -gt 0 -and $sr.Range('M13').Text.Length -gt 0) 'Cycle or missing predecessor was not diagnosed.'
    $controlRows=0; for($row=5;$row -le 304;$row++){if($control.Cells.Item($row,1).Text.Length -gt 0){$controlRows++}}
    Assert-True ($controlRows -ge 3) 'SR diagnostics were not displayed on Control.'
    Assert-True ($sr.Columns.Item(1).Hidden -and $sr.Columns.Item(13).Hidden -and $sr.Columns.Item(15).Hidden) 'Technical SR columns are visible.'
    $book.Save(); $book.Close($true); Release-ComObject $book; $book=$excel.Workbooks.Open($bookPath)
    Assert-True ([double]$book.Worksheets.Item(4).Range('J5').Value2 -eq $mar15) 'SR forecast did not persist after reopen.'
    Write-Output 'PASS: Stage 2 SR network, facts, forecasts, diagnostics, milestone, and reopen passed in Microsoft Excel.'
}
catch { Write-Output ('STAGE2_ERROR_LINE=' + $_.InvocationInfo.ScriptLineNumber); Write-Output ('STAGE2_ERROR=' + $_.Exception.Message); throw }
finally { if($excel){$excel.EnableEvents=$true};if($book){$book.Close($false);Release-ComObject $book};if($excel){$excel.Quit();Release-ComObject $excel};[GC]::Collect();[GC]::WaitForPendingFinalizers();if(Test-Path -LiteralPath $testDir){Remove-Item -LiteralPath $testDir -Recurse -Force} }
