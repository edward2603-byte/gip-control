[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repo = Split-Path -Parent $PSScriptRoot
$testDir = Join-Path ([IO.Path]::GetTempPath()) ("gip-navigator-stage0-" + [guid]::NewGuid().ToString('N'))
$template = Join-Path $testDir 'GIP_Navigator_TEMPLATE.xlsm'
$project = Join-Path $testDir 'Stage0_Project.xlsx.xlsm'
$source = Join-Path $testDir 'source.xlsx'

function Assert-True([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw $Message }
}

function Release-ComObject($Object) {
    if ($null -ne $Object -and [Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Object)
    }
}

New-Item -ItemType Directory -Path $testDir | Out-Null
$excel = $null
$book = $null
$sourceBook = $null
try {
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $repo 'tools\build_xlsm.ps1') -OutputPath $template
    if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $template)) { throw 'Could not build the Stage 0 XLSM.' }
    Copy-Item -LiteralPath $template -Destination $project

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    Write-Output 'PHASE=open-project'
    $book = $excel.Workbooks.Open($project)
    Write-Output 'PHASE=create-source'
    $sourceBook = $excel.Workbooks.Add()
    $sourceSheet = $sourceBook.Worksheets.Item(1)
    $headerCodes = @(
        @(1053,1086,1084,1077,1088,32,1076,1086,1075,1086,1074,1086,1088,1072,32,47,32,1053,1086,1084,1077,1088,32,1053,1072,1088,1103,1076,45,1047,1072,1082,1072,1079,1072),
        @(1053,1072,1080,1084,1077,1085,1086,1074,1072,1085,1080,1077,32,1087,1088,1086,1077,1082,1090,1072),
        @(8470,32,1069,1090,1072,1087,1072),
        @(1053,1072,1080,1084,1077,1085,1086,1074,1072,1085,1080,1077,32,1101,1090,1072,1087,1072),
        @(1053,1072,1095,1072,1083,1086,32,1088,1072,1073,1086,1090),
        @(1054,1082,1086,1085,1095,1072,1085,1080,1077,32,1088,1072,1073,1086,1090),
        @(1057,1090,1086,1080,1084,1086,1089,1090,1100,32,1101,1090,1072,1087,1072,44,32,1088,1091,1073,46,32,1073,1077,1079,32,1053,1044,1057)
    )
    for ($c = 1; $c -le 7; $c++) { $sourceSheet.Cells.Item(1, $c).Value2 = -join ($headerCodes[$c - 1] | ForEach-Object { [char]$_ }) }
    Write-Output 'PHASE=source-headers'
    $sourceSheet.Cells.Item(2, 1).Value2 = 'TEST-01'; $sourceSheet.Cells.Item(2, 2).Value2 = 'Test project'; $sourceSheet.Cells.Item(2, 3).Value2 = '1'; $sourceSheet.Cells.Item(2, 4).Value2 = 'First stage'
    $sourceSheet.Cells.Item(2, 5).Value = '01.03.2027'; $sourceSheet.Cells.Item(2, 6).Value = '05.03.2027'; $sourceSheet.Cells.Item(2, 7).Value2 = '1000'
    $sourceSheet.Cells.Item(3, 1).Value2 = 'TEST-01'; $sourceSheet.Cells.Item(3, 2).Value2 = 'Test project'; $sourceSheet.Cells.Item(3, 3).Value2 = '2'; $sourceSheet.Cells.Item(3, 4).Value2 = 'Second stage'
    $sourceSheet.Cells.Item(3, 5).Value = '06.03.2027'; $sourceSheet.Cells.Item(3, 6).Value = '10.03.2027'; $sourceSheet.Cells.Item(3, 7).Value2 = '2000'
    Write-Output 'PHASE=source-values'
    $sourceBook.SaveAs($source, 51)
    Write-Output 'PHASE=source-saved'
    for ($c = 1; $c -le 7; $c++) { Write-Output ("SOURCE_HEADER_$c=" + [string]$sourceSheet.Cells.Item(1, $c).Value2) }
    $sourceBook.Close($false)
    Release-ComObject $sourceSheet
    Release-ComObject $sourceBook
    $sourceBook = $null

    Write-Output 'PHASE=verify-controls'
    $library = $book.Worksheets.Item(3)
    Assert-True ($library.CheckBoxes().Count -eq 81) 'Library checkbox count is not 81.'
    $selectedCount = 0
    for ($row = 5; $row -le 85; $row++) { if ([bool]$library.Cells.Item($row, 1).Value2) { $selectedCount++ } }
    Assert-True ($selectedCount -eq 81) 'Standard SR rows are not selected by default.'
    $homeSheet = $book.Worksheets.Item(1)
    Assert-True ($homeSheet.Shapes.Item('GIP_Action_Import').OnAction -eq 'ImportKp') 'Import button has no macro binding.'
    $graphAction = [string]$homeSheet.Shapes.Item('GIP_Action_Graph').OnAction
    Assert-True ($graphAction -eq 'GoToGraph') 'Graph button has no macro binding.'
    $excel.Run("'$($book.Name)'!$graphAction")
    Assert-True ($book.ActiveSheet.Index -eq 6) 'Graph button action did not activate the graph sheet.'
    $srAction = [string]$homeSheet.Shapes.Item('GIP_Action_SR').OnAction
    $excel.Run("'$($book.Name)'!$srAction")
    Assert-True ($book.ActiveSheet.Index -eq 4) 'SR button action did not activate the SR sheet.'

    # First import, repeat import with an entered actual, then a failed import.
    Write-Output 'PHASE=first-import'
    $excel.Run("'$($book.Name)'!ImportKpFromPath", $source, $true, $true)
    $kp = $book.Worksheets.Item(2)
    Write-Output ('IMPORT_DIAG=' + $book.Worksheets.Item(8).Range('B20').Text)
    Assert-True ($kp.Range('C5').Text -eq '1' -and $kp.Range('C6').Text -eq '2') 'First KP import did not write two stages.'
    $kp.Range('H5').Value = [datetime]'2027-03-06'
    $kp.Range('I5').Value2 = 1000
    $excel.Run("'$($book.Name)'!ImportKpFromPath", $source, $true, $true)
    Assert-True ([double]$kp.Range('H5').Value2 -eq ([datetime]'2027-03-06').ToOADate() -and [decimal]$kp.Range('I5').Value2 -eq 1000) 'Repeat import did not retain actual data.'
    $book.Worksheets.Item(8).Range('B20').ClearContents()
    $excel.Run("'$($book.Name)'!ImportKpFromPath", (Join-Path $testDir 'invalid.txt'), $true, $true)
    Assert-True ($book.Worksheets.Item(8).Range('B20').Text -match 'ImportKpFromPath') 'Invalid import was not recorded for automation diagnostics.'
    Assert-True ([bool]$excel.EnableEvents) 'Application.EnableEvents remained False after invalid import.'

    # Selection sync and removal of a row with only technical calculated values.
    Write-Output 'PHASE=sync-sr'
    # Default selection of all 81 rows was verified above.  Use a four-row
    # subset for the network regression so the COM test remains fast enough
    # for automated execution while still exercising add/remove behavior.
    $library.Range('A5:A85').Value2 = $false
    $library.Range('A5:A8').Value2 = $true
    $excel.Run("'$($book.Name)'!SyncSelectedSR", $true)
    $sr = $book.Worksheets.Item(4)
    Assert-True ((@($sr.Range('B5:B304').Value2 | Where-Object { $_ -ne $null -and "$_".Length -gt 0 }).Count) -eq 4) 'Selected SR rows did not synchronize to the project.'
    $library.Range('A5').Value2 = $false
    $excel.Run("'$($book.Name)'!SyncSelectedSR", $true)
    Assert-True ((@($sr.Range('B5:B304').Value2 | Where-Object { $_ -eq 'SR-001' }).Count) -eq 0) 'An unchecked SR with only technical values was not removed.'
    $library.Range('A5').Value2 = $true
    $excel.Run("'$($book.Name)'!SyncSelectedSR", $true)
    $excel.Run("'$($book.Name)'!AddSRInstanceByCode", 'SR-001', $true)
    Assert-True ((@($sr.Range('B5:B304').Value2 | Where-Object { $_ -eq 'SR-001' }).Count) -eq 2) 'Second SR instance was not added.'

    # Use an isolated four-work network to test independent, FS and SS starts.
    $sr.Range('A5:O304').ClearContents()
    $links = $book.Worksheets.Item(5)
    $links.Range('A5:G304').ClearContents()
    foreach ($record in @(@('A','SR-001',10,[datetime]'2027-03-01'), @('B','SR-002',1,[datetime]'2027-03-01'), @('C','SR-003',1,$null), @('D','SR-004',1,$null))) {
        $row = 5 + [array]::IndexOf(@('A','B','C','D'), $record[0])
        $sr.Cells.Item($row, 1).Value2 = $record[0]
        $sr.Cells.Item($row, 2).Value2 = $record[1]
        $sr.Cells.Item($row, 5).Value2 = [string]$record[2]
        if ($null -ne $record[3]) { $sr.Cells.Item($row, 15).Value = $record[3].ToString('dd.MM.yyyy') }
    }
    $links.Cells.Item(5, 2).Value2 = 'C'; $links.Cells.Item(5, 3).Value2 = 'A'; $links.Cells.Item(5, 4).Value2 = 'FS'; $links.Cells.Item(5, 5).Value2 = '0'
    $links.Cells.Item(6, 2).Value2 = 'D'; $links.Cells.Item(6, 3).Value2 = 'A'; $links.Cells.Item(6, 4).Value2 = 'SS'; $links.Cells.Item(6, 5).Value2 = '2'
    Write-Output 'PHASE=network'
    $excel.Run("'$($book.Name)'!RecalculateSRNetwork", $true)
    $mar1 = ([datetime]'2027-03-01').ToOADate(); $mar3 = ([datetime]'2027-03-03').ToOADate(); $mar10 = ([datetime]'2027-03-10').ToOADate(); $mar11 = ([datetime]'2027-03-11').ToOADate()
    Write-Output ('NETWORK_VALUES=' + $sr.Range('D5').Value2 + ',' + $sr.Range('F5').Value2 + ',' + $sr.Range('D6').Value2 + ',' + $sr.Range('D7').Value2 + ',' + $sr.Range('D8').Value2 + ',' + $sr.Range('O5').Value2)
    Assert-True ([double]$sr.Range('D5').Value2 -eq $mar1 -and [double]$sr.Range('F5').Value2 -eq $mar10) 'SR A calculation is incorrect.'
    Assert-True ([double]$sr.Range('D6').Value2 -eq $mar1) 'Independent SR B inherited the preceding SR date.'
    Assert-True ([double]$sr.Range('D7').Value2 -eq $mar11) 'FS calculation is incorrect.'
    Assert-True ([double]$sr.Range('D8').Value2 -eq $mar3) 'SS calculation is incorrect.'
    Assert-True ($sr.Range('O5').Text -eq '01.03.2027') 'Start constraint was overwritten by calculation.'

    Write-Output 'PHASE=gantt'
    $excel.Run("'$($book.Name)'!RefreshGantt", $true)
    $graph = $book.Worksheets.Item(6)
    Assert-True ($graph.ChartObjects().Count -eq 1 -and $graph.ChartObjects(1).Chart.SeriesCollection().Count -eq 3) 'Gantt chart was not formed.'
    $beforeRemoval = 0
    for ($row = 5; $row -le 379; $row++) { if ($graph.Cells.Item($row, 1).Text.Length -gt 0) { $beforeRemoval++ } }
    $sr.Range('A6:O6').ClearContents()
    $excel.Run("'$($book.Name)'!RefreshGantt", $true)
    $afterRemoval = 0
    for ($row = 5; $row -le 379; $row++) { if ($graph.Cells.Item($row, 1).Text.Length -gt 0) { $afterRemoval++ } }
    Write-Output ("GANTT_COUNTS=" + $beforeRemoval + ',' + $afterRemoval + ';REMOVED_ID=' + $sr.Range('A6').Text)
    Assert-True ($afterRemoval -eq ($beforeRemoval - 1)) 'Gantt chart retained a deleted SR.'

    Write-Output 'PHASE=reopen'
    $book.Save()
    $book.Close($true)
    Release-ComObject $book
    $book = $excel.Workbooks.Open($project)
    Assert-True ($book.Worksheets.Item(2).Range('C5').Text -eq '1') 'KP data disappeared after save and reopen.'
    Assert-True ($book.Worksheets.Item(5).Range('B5').Text -eq 'C') 'SR link disappeared after reopen.'
    Write-Output 'PASS: Stage 0 XLSM, import, rollback, SR, FS/SS, Gantt, and reopen passed in Microsoft Excel.'
}
catch {
    Write-Output ("STAGE0_ERROR_LINE=" + $_.InvocationInfo.ScriptLineNumber)
    Write-Output ("STAGE0_ERROR=" + $_.Exception.Message)
    throw
}
finally {
    if ($book) { $book.Close($false); Release-ComObject $book }
    if ($sourceBook) { $sourceBook.Close($false); Release-ComObject $sourceBook }
    if ($excel) { $excel.Quit(); Release-ComObject $excel }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if (Test-Path -LiteralPath $testDir) { Remove-Item -LiteralPath $testDir -Recurse -Force }
}
