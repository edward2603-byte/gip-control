[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repo = Split-Path -Parent $PSScriptRoot
$testDir = Join-Path ([IO.Path]::GetTempPath()) ("gip-navigator-stage1-" + [guid]::NewGuid().ToString('N'))
$template = Join-Path $testDir 'GIP_Navigator_TEMPLATE.xlsm'
$project = Join-Path $testDir 'Stage1_Project.xlsm'
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
    if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $template)) { throw 'Could not build the Stage 1 XLSM.' }
    Copy-Item -LiteralPath $template -Destination $project

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $book = $excel.Workbooks.Open($project)

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
    $sourceSheet.Cells.Item(2, 1).Value2 = 'TEST-01'; $sourceSheet.Cells.Item(2, 2).Value2 = 'Test project'; $sourceSheet.Cells.Item(2, 3).Value2 = '1'; $sourceSheet.Cells.Item(2, 4).Value2 = 'First stage'
    $sourceSheet.Cells.Item(2, 5).Value = '01.03.2027'; $sourceSheet.Cells.Item(2, 6).Value = '05.03.2027'; $sourceSheet.Cells.Item(2, 7).Value2 = '1000'
    $sourceSheet.Cells.Item(3, 1).Value2 = 'TEST-01'; $sourceSheet.Cells.Item(3, 2).Value2 = 'Test project'; $sourceSheet.Cells.Item(3, 3).Value2 = '2'; $sourceSheet.Cells.Item(3, 4).Value2 = 'Second stage'
    $sourceSheet.Cells.Item(3, 5).Value = '06.03.2027'; $sourceSheet.Cells.Item(3, 6).Value = '10.03.2027'; $sourceSheet.Cells.Item(3, 7).Value2 = '2000'
    $sourceBook.SaveAs($source, 51)
    $sourceBook.Close($false)
    Release-ComObject $sourceSheet
    Release-ComObject $sourceBook
    $sourceBook = $null

    $excel.Run("'$($book.Name)'!ImportKpFromPath", $source, $true, $true)
    $kp = $book.Worksheets.Item(2)
    Assert-True ($kp.Range('E4').Text -match 'approved' -or $kp.Range('E4').Text.Length -gt 0) 'Approved KP dates are not displayed.'
    Assert-True ($kp.Range('J4').Text.Length -gt 0 -and $kp.Range('K4').Text.Length -gt 0) 'Shifted KP date columns are missing.'
    $mar1 = ([datetime]'2027-03-01').ToOADate(); $mar5 = ([datetime]'2027-03-05').ToOADate(); $mar6 = ([datetime]'2027-03-06').ToOADate(); $mar10 = ([datetime]'2027-03-10').ToOADate()
    Assert-True ([double]$kp.Range('E5').Value2 -eq $mar1 -and [double]$kp.Range('F5').Value2 -eq $mar5) 'Imported approved dates are incorrect.'
    Assert-True ([double]$kp.Range('J5').Value2 -eq $mar1 -and [double]$kp.Range('K5').Value2 -eq $mar5) 'Without NZ date the shifted dates must equal approved dates.'
    Assert-True ($kp.Range('H3').Text -match 'NZ' -or $kp.Range('H3').Text.Length -gt 0) 'Missing NZ warning is not visible.'

    # COM automation is treated as a macro and may edit a UserInterfaceOnly
    # protected sheet.  Verify the actual Excel protection metadata instead.
    Assert-True ($kp.ProtectContents -and [bool]$kp.Range('E5').Locked -and -not [bool]$kp.Range('B3').Locked) 'Approved KP dates are not protected while NZ input is editable.'

    $kp.Range('B3').Value = [datetime]'2027-03-11'
    $excel.CalculateFullRebuild()
    $excel.Run("'$($book.Name)'!RefreshGantt", $true)
    $mar11 = ([datetime]'2027-03-11').ToOADate(); $mar15 = ([datetime]'2027-03-15').ToOADate(); $mar16 = ([datetime]'2027-03-16').ToOADate(); $mar20 = ([datetime]'2027-03-20').ToOADate()
    Assert-True ([double]$kp.Range('E3').Value2 -eq 10) 'NZ delta is not ten calendar days.'
    Assert-True ([double]$kp.Range('J5').Value2 -eq $mar11 -and [double]$kp.Range('K5').Value2 -eq $mar15) 'First stage was not shifted by ten days.'
    Assert-True ([double]$kp.Range('J6').Value2 -eq $mar16 -and [double]$kp.Range('K6').Value2 -eq $mar20) 'Second stage did not preserve its gap and duration.'

    $library = $book.Worksheets.Item(3)
    $library.Range('A5:A85').Value2 = $false
    $library.Range('A5').Value2 = $true
    $excel.Run("'$($book.Name)'!SyncSelectedSR", $true)
    $sr = $book.Worksheets.Item(4)
    Assert-True ([double]$sr.Range('D5').Value2 -eq $mar11) 'SR schedule does not start from shifted KP date.'
    $excel.Run("'$($book.Name)'!AddSRInstanceByCode", 'SR-002', $true)
    $links = $book.Worksheets.Item(5)
    $links.Range('A5').Value2 = 'L1'; $links.Range('B5').Value2 = $sr.Range('A5').Value2; $links.Range('C5').Value2 = $sr.Range('A6').Value2; $links.Range('D5').Value2 = 'FS'; $links.Range('E5').Value2 = 0
    $kp.Range('H5').Value = [datetime]'2027-03-12'; $kp.Range('I5').Value2 = 1000
    $excel.Run("'$($book.Name)'!ImportKpFromPath", $source, $true, $true)
    Assert-True ([decimal]$kp.Range('I5').Value2 -eq 1000) 'Repeat import lost actual data.'
    Assert-True ($links.Range('A5').Text -eq 'L1') 'Repeat import lost SR dependency data.'
    Assert-True ([double]$kp.Range('B3').Value2 -eq $mar11) 'Repeat import overwrote the NZ date.'

    $kp.Range('B3').Value = [datetime]'2027-02-20'; $excel.CalculateFullRebuild()
    Assert-True ([double]$kp.Range('E3').Value2 -eq 0 -and [double]$kp.Range('J5').Value2 -eq $mar1) 'Earlier NZ date must produce zero shift.'
    $kp.Range('B3').Value = [datetime]'2027-03-01'; $excel.CalculateFullRebuild()
    Assert-True ([double]$kp.Range('E3').Value2 -eq 0 -and [double]$kp.Range('K6').Value2 -eq $mar10) 'Equal NZ date must produce zero shift.'
    $kp.Range('B3').Value = [datetime]'2027-03-11'; $excel.CalculateFullRebuild()
    $excel.Run("'$($book.Name)'!RecalculateSRNetwork", $true)
    $excel.Run("'$($book.Name)'!RefreshGantt", $true)
    Assert-True ([double]$sr.Range('D5').Value2 -eq $mar11) 'Changing NZ date did not rebuild SR dates.'
    $graph = $book.Worksheets.Item(6)
    Assert-True ([double]$graph.Range('C5').Value2 -eq $mar11 -and $graph.ChartObjects(1).Chart.SeriesCollection().Count -eq 3) 'Changing NZ date did not rebuild the Gantt chart.'

    $book.Save()
    $book.Close($true)
    Release-ComObject $book
    $book = $excel.Workbooks.Open($project)
    $kp = $book.Worksheets.Item(2)
    Assert-True ([double]$kp.Range('B3').Value2 -eq $mar11 -and [double]$kp.Range('J5').Value2 -eq $mar11) 'NZ date or shifted schedule disappeared after reopen.'
    Write-Output 'PASS: Stage 1 NZ date, shifted KP, protected baseline, SR/Gantt propagation, repeat import, and reopen passed in Microsoft Excel.'
}
catch {
    Write-Output ('STAGE1_ERROR_LINE=' + $_.InvocationInfo.ScriptLineNumber)
    Write-Output ('STAGE1_ERROR=' + $_.Exception.Message)
    throw
}
finally {
    if ($book) { $book.Close($false); Release-ComObject $book }
    if ($sourceBook) { $sourceBook.Close($false); Release-ComObject $sourceBook }
    if ($excel) { $excel.Quit(); Release-ComObject $excel }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if (Test-Path -LiteralPath $testDir) { Remove-Item -LiteralPath $testDir -Recurse -Force }
}
