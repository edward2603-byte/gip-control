Option Explicit

Private Const FIRST_DATA_ROW As Long = 5
Private Const LAST_DATA_ROW As Long = 79
Private Const MIN_STAGES As Long = 2
Private Const MAX_STAGES As Long = 75

Public Sub ImportKp()
    Dim picker As FileDialog
    Set picker = Application.FileDialog(msoFileDialogFilePicker)
    With picker
        .Title = "Выберите исходный календарный план"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Календарный план Excel", "*.xlsx;*.xls"
        If .Show <> -1 Then Exit Sub
        ImportKpFromPath .SelectedItems(1), False
    End With
End Sub

' Public entry point for a controlled Excel test.  The normal user command is ImportKp.
Public Sub ImportKpFromPath(ByVal sourcePath As String, Optional ByVal autoConfirm As Boolean = False, Optional ByVal silent As Boolean = False)
    Dim target As Worksheet, sourceBook As Workbook, source As Worksheet
    Dim headerRow As Long, lastRow As Long, r As Long, stageCount As Long
    Dim stageData(1 To MAX_STAGES, 1 To 7) As Variant
    Dim warnings As String, projectName As String, contractNumber As String, projectKey As String
    Dim seenStages As Object, amount As Currency, amountPresent As Boolean, total As Currency
    Dim previousScreenUpdating As Boolean, previousEnableEvents As Boolean
    Dim backupPath As String

    On Error GoTo Fail
    If Len(Dir$(sourcePath)) = 0 Then Err.Raise vbObjectError + 100, , "Файл КП не найден."
    If LCase$(Right$(sourcePath, 4)) <> ".xls" And LCase$(Right$(sourcePath, 5)) <> ".xlsx" Then _
        Err.Raise vbObjectError + 101, , "Поддерживаются только файлы .xls и .xlsx."

    Set target = ThisWorkbook.Worksheets("01_КП")
    If Application.WorksheetFunction.CountA(target.Range("C" & FIRST_DATA_ROW & ":C" & LAST_DATA_ROW)) > 0 Then
        Err.Raise vbObjectError + 102, , "Повторный импорт остановлен: в 01_КП уже есть данные. Это защищает ранее введённые сведения."
    End If

    previousScreenUpdating = Application.ScreenUpdating
    previousEnableEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set sourceBook = Workbooks.Open(Filename:=sourcePath, ReadOnly:=True, UpdateLinks:=0)
    Set source = sourceBook.Worksheets(1)
    headerRow = FindHeaderRow(source)
    If headerRow = 0 Then Err.Raise vbObjectError + 103, , "Шапка КП не соответствует семи обязательным колонкам в заданном порядке."
    lastRow = source.Cells(source.Rows.Count, 3).End(xlUp).Row
    If lastRow <= headerRow Then Err.Raise vbObjectError + 104, , "В источнике нет этапов КП."

    Set seenStages = CreateObject("Scripting.Dictionary")
    For r = headerRow + 1 To lastRow
        If RowIsBlank(source, r) Then GoTo NextRow
        If Len(Trim$(CStr(source.Cells(r, 3).Value))) = 0 Then
            warnings = warnings & vbCrLf & "• Строка " & r & ": служебная строка не импортирована."
            GoTo NextRow
        End If
        If Len(Trim$(CStr(source.Cells(r, 4).Value))) = 0 Then Err.Raise vbObjectError + 105, , "Строка " & r & ": не заполнено наименование этапа."
        If Len(Trim$(CStr(source.Cells(r, 5).Value))) = 0 And Len(Trim$(CStr(source.Cells(r, 6).Value))) = 0 And Len(Trim$(CStr(source.Cells(r, 7).Value))) = 0 Then
            warnings = warnings & vbCrLf & "• Строка " & r & ": групповой заголовок не импортирован как этап."
            GoTo NextRow
        End If
        If Not IsDate(source.Cells(r, 5).Value) Or Not IsDate(source.Cells(r, 6).Value) Then Err.Raise vbObjectError + 106, , "Строка " & r & ": начало и окончание работ должны быть датами."
        If CDate(source.Cells(r, 5).Value) > CDate(source.Cells(r, 6).Value) Then Err.Raise vbObjectError + 107, , "Строка " & r & ": начало позже окончания."
        If Len(Trim$(CStr(source.Cells(r, 1).Value))) = 0 Or Len(Trim$(CStr(source.Cells(r, 2).Value))) = 0 Then Err.Raise vbObjectError + 108, , "Строка " & r & ": не заполнен реквизит проекта."

        projectKey = Trim$(CStr(source.Cells(r, 1).Value)) & ChrW$(31) & Trim$(CStr(source.Cells(r, 2).Value))
        If Len(contractNumber) = 0 Then
            contractNumber = Trim$(CStr(source.Cells(r, 1).Value))
            projectName = Trim$(CStr(source.Cells(r, 2).Value))
        ElseIf projectKey <> contractNumber & ChrW$(31) & projectName Then
            Err.Raise vbObjectError + 109, , "В источнике смешаны разные проекты или реквизиты."
        End If
        If seenStages.Exists(Trim$(CStr(source.Cells(r, 3).Value))) Then Err.Raise vbObjectError + 110, , "Повторяющийся номер этапа: " & Trim$(CStr(source.Cells(r, 3).Value)) & "."
        seenStages.Add Trim$(CStr(source.Cells(r, 3).Value)), True

        stageCount = stageCount + 1
        If stageCount > MAX_STAGES Then Err.Raise vbObjectError + 111, , "В КП больше " & MAX_STAGES & " этапов."
        stageData(stageCount, 1) = Trim$(CStr(source.Cells(r, 1).Value))
        stageData(stageCount, 2) = Trim$(CStr(source.Cells(r, 2).Value))
        stageData(stageCount, 3) = Trim$(CStr(source.Cells(r, 3).Value))
        stageData(stageCount, 4) = Trim$(CStr(source.Cells(r, 4).Value))
        stageData(stageCount, 5) = CDate(source.Cells(r, 5).Value)
        stageData(stageCount, 6) = CDate(source.Cells(r, 6).Value)
        ParseAmount source.Cells(r, 7).Value, amountPresent, amount
        If amountPresent Then
            stageData(stageCount, 7) = amount
            total = total + amount
        Else
            warnings = warnings & vbCrLf & "• Строка " & r & ": стоимость не указана; импортируется пустое значение."
        End If
NextRow:
    Next r

    If stageCount < MIN_STAGES Then Err.Raise vbObjectError + 112, , "Количество этапов должно быть от " & MIN_STAGES & " до " & MAX_STAGES & ". Получено: " & stageCount & "."
    If Not autoConfirm Then
        If MsgBox("Проект: " & projectName & vbCrLf & "Этапов к импорту: " & stageCount & vbCrLf & "Стоимость без НДС: " & Format$(total, "#,##0.00") & vbCrLf & _
                  IIf(Len(warnings) > 0, "Есть предупреждения." & vbCrLf, "") & vbCrLf & "Нажмите ОК для импорта или Отмена для выхода.", vbOKCancel + vbQuestion, "Предпросмотр импорта КП") <> vbOK Then GoTo SafeExit
    End If

    backupPath = ThisWorkbook.Path & Application.PathSeparator & "Резервная_копия_до_импорта_" & Format$(Now, "yyyymmdd_hhnnss") & ".xlsm"
    ThisWorkbook.SaveCopyAs backupPath
    target.Range("A" & FIRST_DATA_ROW & ":G" & LAST_DATA_ROW).ClearContents
    target.Range("A" & FIRST_DATA_ROW).Resize(stageCount, 7).Value = stageData
    ThisWorkbook.Worksheets("99_Служебный").Range("B7").Value = Now
    ThisWorkbook.Worksheets("99_Служебный").Range("B7").NumberFormat = "dd.mm.yyyy hh:mm"
    ThisWorkbook.Save
    If Not silent Then MsgBox "КП импортирован: " & stageCount & " этапов." & vbCrLf & "Резервная копия: " & backupPath & IIf(Len(warnings) > 0, vbCrLf & "Предупреждения:" & warnings, ""), vbInformation, "GIP Navigator"

SafeExit:
    On Error Resume Next
    If Not sourceBook Is Nothing Then sourceBook.Close SaveChanges:=False
    Application.ScreenUpdating = previousScreenUpdating
    Application.EnableEvents = previousEnableEvents
    Exit Sub
Fail:
    If Not silent Then MsgBox "Импорт не выполнен. " & Err.Description, vbExclamation, "GIP Navigator"
    Resume SafeExit
End Sub

Private Function FindHeaderRow(ByVal source As Worksheet) As Long
    Dim r As Long, c As Long, extraLastColumn As Long
    For r = 1 To 30
        If HeaderMatches(source.Cells(r, 1).Value, 1) And HeaderMatches(source.Cells(r, 2).Value, 2) And HeaderMatches(source.Cells(r, 3).Value, 3) And _
           HeaderMatches(source.Cells(r, 4).Value, 4) And HeaderMatches(source.Cells(r, 5).Value, 5) And HeaderMatches(source.Cells(r, 6).Value, 6) And HeaderMatches(source.Cells(r, 7).Value, 7) Then
            extraLastColumn = source.Cells(r, source.Columns.Count).End(xlToLeft).Column
            For c = 8 To extraLastColumn
                If Len(Trim$(CStr(source.Cells(r, c).Value))) > 0 Then Err.Raise vbObjectError + 113, , "Шапка КП содержит данные после седьмой колонки."
            Next c
            FindHeaderRow = r
            Exit Function
        End If
    Next r
End Function

Private Function HeaderMatches(ByVal value As Variant, ByVal position As Long) As Boolean
    Dim normalized As String
    normalized = NormalizeText(CStr(value))
    Select Case position
        Case 1: HeaderMatches = (normalized = "номер договора/номер наряд-заказа" Or normalized = "номер договора")
        Case 2: HeaderMatches = (normalized = "наименование проекта")
        Case 3: HeaderMatches = (normalized = "№ этапа" Or normalized = "этап")
        Case 4: HeaderMatches = (normalized = "наименование этапа" Or normalized = "наименование этапа по договору")
        Case 5: HeaderMatches = (normalized = "начало работ")
        Case 6: HeaderMatches = (normalized = "окончание работ")
        Case 7: HeaderMatches = (normalized = "стоимость этапа, руб без ндс" Or normalized = "стоимость этапа, руб. без ндс")
    End Select
End Function

Private Function NormalizeText(ByVal value As String) As String
    value = Replace(value, vbCr, " ")
    value = Replace(value, vbLf, " ")
    value = Replace(value, ChrW$(160), " ")
    NormalizeText = LCase$(Application.WorksheetFunction.Trim(value))
End Function

Private Function RowIsBlank(ByVal source As Worksheet, ByVal rowNumber As Long) As Boolean
    Dim c As Long
    RowIsBlank = True
    For c = 1 To 7
        If Len(Trim$(CStr(source.Cells(rowNumber, c).Value))) > 0 Then RowIsBlank = False: Exit Function
    Next c
End Function

Private Sub ParseAmount(ByVal value As Variant, ByRef present As Boolean, ByRef amount As Currency)
    Dim raw As String
    present = False
    If IsEmpty(value) Or Len(Trim$(CStr(value))) = 0 Then Exit Sub
    If IsNumeric(value) Then amount = CCur(value): present = True: Exit Sub
    raw = Replace(CStr(value), ChrW$(160), "")
    raw = Replace(raw, " ", "")
    raw = Replace(LCase$(raw), "руб.", "")
    raw = Replace(LCase$(raw), "руб", "")
    raw = Replace(raw, "₽", "")
    If InStr(raw, ",") > 0 And InStr(raw, ".") > 0 Then raw = Replace(raw, ".", "")
    If Application.International(xlDecimalSeparator) = "," Then raw = Replace(raw, ".", ",") Else raw = Replace(raw, ",", ".")
    If Not IsNumeric(raw) Then Err.Raise vbObjectError + 114, , "Некорректная стоимость этапа: " & CStr(value)
    amount = CCur(raw)
    present = True
End Sub

Public Sub GoToSR(): ThisWorkbook.Worksheets("03_СР_проекта").Activate: End Sub
Public Sub GoToGraph(): ThisWorkbook.Worksheets("05_График").Activate: End Sub
Public Sub GoToControl(): ThisWorkbook.Worksheets("06_Контроль").Activate: End Sub
Public Sub GoToActs(): ThisWorkbook.Worksheets("07_Акты").Activate: End Sub
