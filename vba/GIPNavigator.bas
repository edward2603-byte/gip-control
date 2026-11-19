Option Explicit

Private Const FIRST_DATA_ROW As Long = 5
Private Const LAST_DATA_ROW As Long = 79
Private Const MIN_STAGES As Long = 2
Private Const MAX_STAGES As Long = 75
Private Const SR_FIRST_ROW As Long = 5
Private Const SR_LAST_ROW As Long = 304
Private Const SR_NETWORK_COLUMN As Long = 13
Private Const SR_SOURCE_COLUMN As Long = 14

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

' Public entry point for an automated Excel test. The usual command is ImportKp.
Public Sub ImportKpFromPath(ByVal sourcePath As String, Optional ByVal autoConfirm As Boolean = False, Optional ByVal silent As Boolean = False)
    Dim target As Worksheet, sourceBook As Workbook, source As Worksheet
    Dim headerRow As Long, lastRow As Long, r As Long, stageCount As Long, i As Long
    Dim stageData(1 To MAX_STAGES, 1 To 7) As Variant, importData() As Variant
    Dim warnings As String, projectName As String, contractNumber As String, projectKey As String, stageKey As String
    Dim seenStages As Object, actualByStage As Object, actualValue As Variant
    Dim amount As Currency, amountPresent As Boolean, total As Currency
    Dim previousScreenUpdating As Boolean, previousEnableEvents As Boolean
    Dim backupPath As String, existingCount As Long, preservedActuals As Long, previewText As String
    Dim oldProjectStart As Variant, newProjectStart As Date

    On Error GoTo Fail
    If Len(Dir$(sourcePath)) = 0 Then Err.Raise vbObjectError + 100, , "Файл КП не найден."
    If LCase$(Right$(sourcePath, 4)) <> ".xls" And LCase$(Right$(sourcePath, 5)) <> ".xlsx" Then _
        Err.Raise vbObjectError + 101, , "Поддерживаются только файлы .xls и .xlsx."
    Set target = ThisWorkbook.Worksheets("01_КП")
    existingCount = Application.WorksheetFunction.CountA(target.Range("C" & FIRST_DATA_ROW & ":C" & LAST_DATA_ROW))
    oldProjectStart = ProjectStartDate()
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
            contractNumber = Trim$(CStr(source.Cells(r, 1).Value)): projectName = Trim$(CStr(source.Cells(r, 2).Value))
        ElseIf projectKey <> contractNumber & ChrW$(31) & projectName Then
            Err.Raise vbObjectError + 109, , "В источнике смешаны разные проекты или реквизиты."
        End If
        stageKey = StableStageId(source.Cells(r, 3).Value)
        If seenStages.Exists(stageKey) Then Err.Raise vbObjectError + 110, , "Повторяющийся номер этапа: " & Trim$(CStr(source.Cells(r, 3).Value)) & "."
        seenStages.Add stageKey, True
        stageCount = stageCount + 1
        If stageCount > MAX_STAGES Then Err.Raise vbObjectError + 111, , "В КП больше " & MAX_STAGES & " этапов."
        stageData(stageCount, 1) = Trim$(CStr(source.Cells(r, 1).Value))
        stageData(stageCount, 2) = Trim$(CStr(source.Cells(r, 2).Value))
        stageData(stageCount, 3) = Trim$(CStr(source.Cells(r, 3).Value))
        stageData(stageCount, 4) = Trim$(CStr(source.Cells(r, 4).Value))
        stageData(stageCount, 5) = CDate(source.Cells(r, 5).Value)
        stageData(stageCount, 6) = CDate(source.Cells(r, 6).Value)
        If stageCount = 1 Or CDate(stageData(stageCount, 5)) < newProjectStart Then newProjectStart = CDate(stageData(stageCount, 5))
        ParseAmount source.Cells(r, 7).Value, amountPresent, amount
        If amountPresent Then total = total + amount: stageData(stageCount, 7) = amount Else warnings = warnings & vbCrLf & "• Строка " & r & ": стоимость не указана; импортируется пустое значение."
NextRow:
    Next r
    If stageCount < MIN_STAGES Then Err.Raise vbObjectError + 112, , "Количество этапов должно быть от " & MIN_STAGES & " до " & MAX_STAGES & ". Получено: " & stageCount & "."
    Set actualByStage = ExistingActualsByStage(target, seenStages)
    preservedActuals = actualByStage.Count
    If Not autoConfirm Then
        previewText = "Проект: " & projectName & vbCrLf & "Этапов к импорту: " & stageCount & vbCrLf & "Стоимость без НДС: " & Format$(total, "#,##0.00")
        If existingCount > 0 Then previewText = previewText & vbCrLf & vbCrLf & "Повторный импорт: факты H:I будут сохранены по номеру этапа (" & preservedActuals & ")." & vbCrLf & "Начало проекта: " & DisplayDate(oldProjectStart) & " → " & Format$(newProjectStart, "dd.mm.yyyy") & "."
        If Len(warnings) > 0 Then previewText = previewText & vbCrLf & "Есть предупреждения."
        If MsgBox(previewText & vbCrLf & vbCrLf & "Нажмите ОК для импорта или Отмена для выхода.", vbOKCancel + vbQuestion, "Предпросмотр импорта КП") <> vbOK Then GoTo SafeExit
    End If
    backupPath = ThisWorkbook.Path & Application.PathSeparator & "Резервная_копия_до_импорта_" & Format$(Now, "yyyymmdd_hhnnss") & ".xlsm"
    If Not silent Then ThisWorkbook.SaveCopyAs backupPath
    target.Range("A" & FIRST_DATA_ROW & ":G" & LAST_DATA_ROW).ClearContents
    target.Range("H" & FIRST_DATA_ROW & ":I" & LAST_DATA_ROW).ClearContents
    ReDim importData(1 To stageCount, 1 To 7)
    For r = 1 To stageCount
        For i = 1 To 7: importData(r, i) = stageData(r, i): Next i
    Next r
    target.Range("A" & FIRST_DATA_ROW).Resize(stageCount, 7).Value = importData
    For i = 1 To stageCount
        If actualByStage.Exists(StableStageId(stageData(i, 3))) Then
            actualValue = actualByStage(StableStageId(stageData(i, 3)))
            target.Cells(FIRST_DATA_ROW + i - 1, 8).Value = actualValue(0)
            target.Cells(FIRST_DATA_ROW + i - 1, 9).Value = actualValue(1)
        End If
    Next i
    ThisWorkbook.Worksheets("99_Служебный").Range("B6").Value = Now
    ThisWorkbook.Worksheets("99_Служебный").Range("B6").NumberFormat = "dd.mm.yyyy hh:mm"
    RecalculateSRNetwork True
    RefreshGantt
    ' Persist calculated Gantt source values even when the user's Excel is in
    ' manual calculation mode.
    Application.CalculateFullRebuild
    ' Automated verification is non-persistent; the user command always saves.
    If Not silent Then ThisWorkbook.Save
    If Not silent Then MsgBox "КП импортирован: " & stageCount & " этапов." & vbCrLf & "Резервная копия: " & backupPath, vbInformation, "GIP Navigator"
SafeExit:
    On Error Resume Next
    If Not sourceBook Is Nothing Then sourceBook.Close SaveChanges:=False
    Application.ScreenUpdating = previousScreenUpdating: Application.EnableEvents = previousEnableEvents
    Exit Sub
Fail:
    If silent Then
        On Error Resume Next
        If Not sourceBook Is Nothing Then sourceBook.Close SaveChanges:=False
        Application.ScreenUpdating = previousScreenUpdating: Application.EnableEvents = previousEnableEvents
        Exit Sub
    End If
    If Not silent Then MsgBox "Импорт не выполнен. " & Err.Description, vbExclamation, "GIP Navigator"
    Resume SafeExit
End Sub

Private Function ExistingActualsByStage(ByVal target As Worksheet, ByVal incomingStages As Object) As Object
    Dim result As Object, r As Long, stageId As String, hasFact As Boolean
    Set result = CreateObject("Scripting.Dictionary")
    For r = FIRST_DATA_ROW To LAST_DATA_ROW
        stageId = StableStageId(target.Cells(r, 3).Value)
        hasFact = Len(Trim$(CStr(target.Cells(r, 8).Value))) > 0 Or Len(Trim$(CStr(target.Cells(r, 9).Value))) > 0
        If Len(stageId) > 0 And hasFact Then
            If Not incomingStages.Exists(stageId) Then Err.Raise vbObjectError + 115, , "Конфликт повторного импорта: этап «" & stageId & "» с фактическим актированием отсутствует в новом КП."
            If result.Exists(stageId) Then Err.Raise vbObjectError + 116, , "Конфликт повторного импорта: повторяется номер этапа «" & stageId & "»."
            result.Add stageId, Array(target.Cells(r, 8).Value, target.Cells(r, 9).Value)
        End If
    Next r
    Set ExistingActualsByStage = result
End Function

' Called by the checkboxes on 02_Библиотека_СР. Only GIP_LIBRARY rows are managed.
Public Sub SyncSelectedSR(Optional ByVal silent As Boolean = False)
    Const LIBRARY_FIRST_ROW As Long = 5, LIBRARY_LAST_ROW As Long = 85
    Dim library As Worksheet, projectSR As Worksheet, r As Long, projectRow As Long, freeRow As Long
    Dim code As String, protectedCount As Long, oldScreenUpdating As Boolean, oldEnableEvents As Boolean
    On Error GoTo Fail
    Set library = ThisWorkbook.Worksheets("02_Библиотека_СР"): Set projectSR = ThisWorkbook.Worksheets("03_СР_проекта")
    oldScreenUpdating = Application.ScreenUpdating: oldEnableEvents = Application.EnableEvents
    Application.ScreenUpdating = False: Application.EnableEvents = False
    For r = LIBRARY_FIRST_ROW To LIBRARY_LAST_ROW
        code = Trim$(CStr(library.Cells(r, 2).Value))
        If Len(code) = 0 Then GoTo NextLibraryRow
        projectRow = FindManagedSRRow(projectSR, code)
        If IsSelected(library.Cells(r, 1).Value) Then
            If projectRow = 0 Then
                freeRow = FindFreeSRRow(projectSR)
                If freeRow = 0 Then Err.Raise vbObjectError + 120, , "В 03_СР_проекта нет свободной строки для СР " & code & "."
                projectSR.Cells(freeRow, 1).Value = NextInstanceId(projectSR, code)
                projectSR.Cells(freeRow, 2).Value = code
                projectSR.Cells(freeRow, 4).Value = ProjectStartDate()
                projectSR.Cells(freeRow, SR_SOURCE_COLUMN).Value = "GIP_LIBRARY"
            End If
        ElseIf projectRow > 0 Then
            If SRRowHasProjectData(projectSR, projectRow) Then
                library.Cells(r, 1).Value = True: protectedCount = protectedCount + 1
            Else
                projectSR.Range(projectSR.Cells(projectRow, 1), projectSR.Cells(projectRow, SR_SOURCE_COLUMN)).ClearContents
            End If
        End If
NextLibraryRow:
    Next r
    RecalculateSRNetwork True
    If protectedCount > 0 And Not silent Then MsgBox "СР с введёнными проектными данными оставлены в плане: " & protectedCount & ".", vbInformation, "GIP Navigator"
SafeExit:
    On Error Resume Next
    Application.ScreenUpdating = oldScreenUpdating: Application.EnableEvents = oldEnableEvents
    Exit Sub
Fail:
    If Not silent Then MsgBox "Выбор СР не выполнен. " & Err.Description, vbExclamation, "GIP Navigator"
    Resume SafeExit
End Sub

Public Sub AddSRInstance()
    Dim projectSR As Worksheet, code As String, rowNumber As Long
    code = Trim$(InputBox("Введите код типовой СР (например, SR-001). Повторный код создаст отдельный экземпляр.", "Добавить экземпляр СР"))
    If Len(code) = 0 Then Exit Sub
    If Application.WorksheetFunction.CountIf(ThisWorkbook.Worksheets("02_Библиотека_СР").Range("B5:B85"), code) = 0 Then MsgBox "Код отсутствует в библиотеке СР: " & code, vbExclamation, "GIP Navigator": Exit Sub
    Set projectSR = ThisWorkbook.Worksheets("03_СР_проекта"): rowNumber = FindFreeSRRow(projectSR)
    If rowNumber = 0 Then MsgBox "В 03_СР_проекта нет свободной строки.", vbExclamation, "GIP Navigator": Exit Sub
    Application.EnableEvents = False
    projectSR.Cells(rowNumber, 1).Value = NextInstanceId(projectSR, code)
    projectSR.Cells(rowNumber, 2).Value = code
    projectSR.Cells(rowNumber, 4).Value = ProjectStartDate()
    projectSR.Cells(rowNumber, SR_SOURCE_COLUMN).Value = "GIP_LIBRARY"
    RecalculateSRNetwork True
    Application.EnableEvents = True: projectSR.Activate: projectSR.Cells(rowNumber, 1).Select
End Sub

' Inclusive dates. FS means the next date after predecessor finish; SS means
' predecessor start. A lag is counted in the calendar of the successor.
Public Sub RecalculateSRNetwork(Optional ByVal silent As Boolean = False)
    Dim sr As Worksheet, links As Worksheet, rows As Object, indegree As Object, successors As Object, predecessors As Object
    Dim queue As Collection, order As Collection, itemId As Variant, r As Long, linkRow As Long
    Dim successorId As Variant, predecessorId As String, relation As String
    Dim projectStart As Variant, duration As Variant, ownStart As Variant, calculatedStart As Date, calculatedEnd As Date
    Dim candidates As Collection, linkData As Variant, candidate As Variant, maxDate As Date, hasCandidate As Boolean, oldEnableEvents As Boolean
    On Error GoTo Fail
    Set sr = ThisWorkbook.Worksheets("03_СР_проекта"): Set links = ThisWorkbook.Worksheets("04_Согласование_плана")
    projectStart = ProjectStartDate(): oldEnableEvents = Application.EnableEvents: Application.EnableEvents = False
    Set rows = CreateObject("Scripting.Dictionary"): Set indegree = CreateObject("Scripting.Dictionary")
    Set successors = CreateObject("Scripting.Dictionary"): Set predecessors = CreateObject("Scripting.Dictionary")
    For r = SR_FIRST_ROW To SR_LAST_ROW
        If Len(Trim$(CStr(sr.Cells(r, 1).Value))) > 0 Then
            itemId = Trim$(CStr(sr.Cells(r, 1).Value)): sr.Cells(r, SR_NETWORK_COLUMN).ClearContents: sr.Cells(r, 6).ClearContents
            If rows.Exists(itemId) Then
                AddSRIssue sr, r, "Повторяется ID экземпляра СР."
            Else
                rows.Add itemId, r: indegree.Add itemId, 0
                Set successors(itemId) = New Collection: Set predecessors(itemId) = New Collection
            End If
        End If
    Next r
    If rows.Count = 0 Then GoTo SafeExit
    If Not IsDate(projectStart) Then
        For Each itemId In rows.Keys: AddSRIssue sr, CLng(rows(itemId)), "Нет дат КП: невозможно определить начало проекта.": Next itemId
        GoTo SafeExit
    End If
    For linkRow = SR_FIRST_ROW To SR_LAST_ROW
        successorId = Trim$(CStr(links.Cells(linkRow, 2).Value)): predecessorId = Trim$(CStr(links.Cells(linkRow, 3).Value))
        If Len(successorId) = 0 And Len(predecessorId) = 0 Then GoTo NextLink
        relation = UCase$(Trim$(CStr(links.Cells(linkRow, 4).Value)))
        If Len(successorId) = 0 Or Len(predecessorId) = 0 Then
            links.Cells(linkRow, 6).Value = "Ошибка: заполните оба ID": GoTo NextLink
        End If
        If Not rows.Exists(successorId) Or Not rows.Exists(predecessorId) Then
            links.Cells(linkRow, 6).Value = "Ошибка: СР не найдена"
            If rows.Exists(successorId) Then AddSRIssue sr, CLng(rows(successorId)), "Связь с невыбранной СР: " & predecessorId
            GoTo NextLink
        End If
        If successorId = predecessorId Or (relation <> "FS" And relation <> "SS") Or (Len(Trim$(CStr(links.Cells(linkRow, 5).Value))) > 0 And Not IsNumeric(links.Cells(linkRow, 5).Value)) Then
            links.Cells(linkRow, 6).Value = "Ошибка: проверьте тип/лаг": AddSRIssue sr, CLng(rows(successorId)), "Некорректная связь.": GoTo NextLink
        End If
        links.Cells(linkRow, 6).Value = "Готово"
        predecessors(successorId).Add Array(predecessorId, relation, CLng(Val(links.Cells(linkRow, 5).Value)))
        successors(predecessorId).Add successorId: indegree(successorId) = CLng(indegree(successorId)) + 1
NextLink:
    Next linkRow
    Set queue = New Collection: Set order = New Collection
    For Each itemId In rows.Keys: If CLng(indegree(itemId)) = 0 Then queue.Add itemId
    Next itemId
    Do While queue.Count > 0
        itemId = queue(1): queue.Remove 1: order.Add itemId
        For Each successorId In successors(itemId)
            indegree(successorId) = CLng(indegree(successorId)) - 1
            If CLng(indegree(successorId)) = 0 Then queue.Add successorId
        Next successorId
    Loop
    If order.Count <> rows.Count Then
        For Each itemId In rows.Keys: If CLng(indegree(itemId)) > 0 Then AddSRIssue sr, CLng(rows(itemId)), "Обнаружен цикл зависимостей СР."
        Next itemId
    End If
    Dim calculated As Object: Set calculated = CreateObject("Scripting.Dictionary")
    For Each itemId In order
        r = CLng(rows(itemId)): duration = sr.Cells(r, 5).Value
        If Not IsNumeric(duration) Or CLng(duration) < 0 Then AddSRIssue sr, r, "Не задана длительность СР.": GoTo NextItem
        ownStart = sr.Cells(r, 4).Value
        If IsDate(ownStart) And CDate(ownStart) < CDate(projectStart) Then AddSRIssue sr, r, "Плановое начало раньше начала проекта.": GoTo NextItem
        Set candidates = New Collection
        If IsDate(ownStart) Then candidates.Add CDate(ownStart)
        If predecessors(itemId).Count = 0 And Not IsDate(ownStart) Then candidates.Add CDate(projectStart)
        For Each linkData In predecessors(itemId)
            If calculated.Exists(CStr(linkData(0))) Then
                If CStr(linkData(1)) = "FS" Then candidate = DateAdd("d", CLng(linkData(2)) + 1, CDate(calculated(CStr(linkData(0)))(1))) Else candidate = DateAdd("d", CLng(linkData(2)), CDate(calculated(CStr(linkData(0)))(0)))
                candidates.Add candidate
            Else
                AddSRIssue sr, r, "Нет срока предшествующей СР: " & CStr(linkData(0))
            End If
        Next linkData
        For Each candidate In candidates: If Not hasCandidate Or CDate(candidate) > maxDate Then maxDate = CDate(candidate): hasCandidate = True
        Next candidate
        If Not hasCandidate Then AddSRIssue sr, r, "Не определено плановое начало СР.": GoTo NextItem
        calculatedStart = maxDate: calculatedEnd = DateAdd("d", CLng(duration) - 1, calculatedStart)
        sr.Cells(r, 4).Value = calculatedStart: sr.Cells(r, 6).Value = calculatedEnd
        calculated.Add CStr(itemId), Array(calculatedStart, calculatedEnd)
NextItem:
    Next itemId
SafeExit:
    RefreshGantt
    Application.EnableEvents = oldEnableEvents
    Exit Sub
Fail:
    If Not silent Then MsgBox "Расчёт сети СР не выполнен. " & Err.Description, vbExclamation, "GIP Navigator"
    Resume SafeExit
End Sub

Private Function ProjectStartDate() As Variant
    Dim values As Variant, r As Long, found As Boolean, result As Date
    values = ThisWorkbook.Worksheets("01_КП").Range("E5:E79").Value
    For r = 1 To UBound(values, 1)
        If IsDate(values(r, 1)) Then If Not found Or CDate(values(r, 1)) < result Then result = CDate(values(r, 1)): found = True
    Next r
    If found Then ProjectStartDate = result
End Function


Private Sub AddSRIssue(ByVal sr As Worksheet, ByVal rowNumber As Long, ByVal issueText As String)
    If Len(sr.Cells(rowNumber, SR_NETWORK_COLUMN).Value) > 0 Then sr.Cells(rowNumber, SR_NETWORK_COLUMN).Value = sr.Cells(rowNumber, SR_NETWORK_COLUMN).Value & " "
    sr.Cells(rowNumber, SR_NETWORK_COLUMN).Value = sr.Cells(rowNumber, SR_NETWORK_COLUMN).Value & issueText
End Sub
Private Function IsSelected(ByVal value As Variant) As Boolean: If VarType(value) = vbBoolean Then IsSelected = CBool(value) Else IsSelected = (LCase$(Trim$(CStr(value))) = "true" Or Trim$(CStr(value)) = "-1")
End Function
Private Function FindManagedSRRow(ByVal projectSR As Worksheet, ByVal code As String) As Long
    Dim r As Long
    For r = SR_FIRST_ROW To SR_LAST_ROW
        If CStr(projectSR.Cells(r, SR_SOURCE_COLUMN).Value) = "GIP_LIBRARY" And Trim$(CStr(projectSR.Cells(r, 2).Value)) = code Then FindManagedSRRow = r: Exit Function
    Next r
End Function
Private Function FindFreeSRRow(ByVal projectSR As Worksheet) As Long
    Dim r As Long
    For r = SR_FIRST_ROW To SR_LAST_ROW: If Len(Trim$(CStr(projectSR.Cells(r, 2).Value))) = 0 Then FindFreeSRRow = r: Exit Function
    Next r
End Function
Private Function NextInstanceId(ByVal projectSR As Worksheet, ByVal code As String) As String
    Dim r As Long, index As Long
    For r = SR_FIRST_ROW To SR_LAST_ROW: If Trim$(CStr(projectSR.Cells(r, 2).Value)) = code Then index = index + 1
    Next r
    NextInstanceId = "GIP-" & code & "-" & Format$(index + 1, "000")
End Function
Private Function SRRowHasProjectData(ByVal projectSR As Worksheet, ByVal rowNumber As Long) As Boolean
    Dim columnNumber As Variant
    For Each columnNumber In Array(4, 5, 7, 8, 9, 10, 11, 12)
        If Len(Trim$(CStr(projectSR.Cells(rowNumber, CLng(columnNumber)).Value))) > 0 Then SRRowHasProjectData = True: Exit Function
    Next columnNumber
End Function
Private Function FindHeaderRow(ByVal source As Worksheet) As Long
    Dim r As Long, c As Long, extraLastColumn As Long
    For r = 1 To 30
        If HeaderMatches(source.Cells(r, 1).Value, 1) And HeaderMatches(source.Cells(r, 2).Value, 2) And HeaderMatches(source.Cells(r, 3).Value, 3) And HeaderMatches(source.Cells(r, 4).Value, 4) And HeaderMatches(source.Cells(r, 5).Value, 5) And HeaderMatches(source.Cells(r, 6).Value, 6) And HeaderMatches(source.Cells(r, 7).Value, 7) Then
            extraLastColumn = source.Cells(r, source.Columns.Count).End(xlToLeft).Column
            For c = 8 To extraLastColumn: If Len(Trim$(CStr(source.Cells(r, c).Value))) > 0 Then Err.Raise vbObjectError + 113, , "Шапка КП содержит данные после седьмой колонки."
            Next c
            FindHeaderRow = r: Exit Function
        End If
    Next r
End Function
Private Function HeaderMatches(ByVal value As Variant, ByVal position As Long) As Boolean
    Dim normalized As String: normalized = NormalizeText(CStr(value))
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
    value = Replace(value, vbCr, " "): value = Replace(value, vbLf, " "): value = Replace(value, ChrW$(160), " ")
    NormalizeText = LCase$(Application.WorksheetFunction.Trim(value))
End Function
Private Function RowIsBlank(ByVal source As Worksheet, ByVal rowNumber As Long) As Boolean
    Dim c As Long: RowIsBlank = True
    For c = 1 To 7: If Len(Trim$(CStr(source.Cells(rowNumber, c).Value))) > 0 Then RowIsBlank = False: Exit Function
    Next c
End Function
Private Sub ParseAmount(ByVal value As Variant, ByRef present As Boolean, ByRef amount As Currency)
    Dim raw As String: present = False
    If IsEmpty(value) Or Len(Trim$(CStr(value))) = 0 Then Exit Sub
    If IsNumeric(value) Then amount = CCur(value): present = True: Exit Sub
    raw = Replace(CStr(value), ChrW$(160), ""): raw = Replace(raw, " ", ""): raw = Replace(LCase$(raw), "руб.", ""): raw = Replace(LCase$(raw), "руб", ""): raw = Replace(raw, "₽", "")
    If InStr(raw, ",") > 0 And InStr(raw, ".") > 0 Then raw = Replace(raw, ".", "")
    If Application.International(xlDecimalSeparator) = "," Then raw = Replace(raw, ".", ",") Else raw = Replace(raw, ",", ".")
    If Not IsNumeric(raw) Then Err.Raise vbObjectError + 114, , "Некорректная стоимость этапа: " & CStr(value)
    amount = CCur(raw): present = True
End Sub
Private Function DisplayDate(ByVal value As Variant) As String: If IsDate(value) Then DisplayDate = Format$(CDate(value), "dd.mm.yyyy") Else DisplayDate = "не задано"
End Function
Private Function StableStageId(ByVal value As Variant) As String
    StableStageId = Replace(Trim$(CStr(value)), ",", ".")
End Function

Public Sub RefreshGantt()
    Dim graph As Worksheet, kp As Worksheet, sr As Worksheet, chartObject As ChartObject, chart As Chart
    Dim sourceRow As Long, outputRow As Long, lastRow As Long, startDate As Variant, endDate As Variant
    On Error GoTo SafeExit
    Set graph = ThisWorkbook.Worksheets("05_График")
    Set kp = ThisWorkbook.Worksheets("01_КП")
    Set sr = ThisWorkbook.Worksheets("03_СР_проекта")
    graph.Range("A5:F379").ClearContents
    graph.Rows("5:379").Hidden = False
    outputRow = 5
    For sourceRow = FIRST_DATA_ROW To LAST_DATA_ROW
        If Len(Trim$(CStr(kp.Cells(sourceRow, 3).Value))) > 0 And IsDate(kp.Cells(sourceRow, 5).Value) And IsDate(kp.Cells(sourceRow, 6).Value) Then
            graph.Cells(outputRow, 1).Value = "КП"
            graph.Cells(outputRow, 2).Value = kp.Cells(sourceRow, 4).Value
            graph.Cells(outputRow, 3).Value = CDate(kp.Cells(sourceRow, 5).Value)
            graph.Cells(outputRow, 4).Value = DateDiff("d", CDate(kp.Cells(sourceRow, 5).Value), CDate(kp.Cells(sourceRow, 6).Value)) + 1
            graph.Cells(outputRow, 5).Value = graph.Cells(outputRow, 4).Value
            outputRow = outputRow + 1
        End If
    Next sourceRow
    For sourceRow = SR_FIRST_ROW To SR_LAST_ROW
        startDate = sr.Cells(sourceRow, 4).Value: endDate = sr.Cells(sourceRow, 6).Value
        If Len(Trim$(CStr(sr.Cells(sourceRow, 1).Value))) > 0 And IsDate(startDate) And IsDate(endDate) Then
            graph.Cells(outputRow, 1).Value = "СР"
            graph.Cells(outputRow, 2).Value = sr.Cells(sourceRow, 3).Value
            graph.Cells(outputRow, 3).Value = CDate(startDate)
            graph.Cells(outputRow, 4).Value = DateDiff("d", CDate(startDate), CDate(endDate)) + 1
            graph.Cells(outputRow, 6).Value = graph.Cells(outputRow, 4).Value
            outputRow = outputRow + 1
        End If
    Next sourceRow
    If outputRow <= 379 Then graph.Rows(outputRow & ":379").Hidden = True
    Set chartObject = graph.ChartObjects(1)
    Set chart = chartObject.Chart
    chartObject.Left = graph.Range("H4").Left: chartObject.Top = graph.Range("H4").Top
    chartObject.Width = 1050: chartObject.Height = 600
    Do While chart.SeriesCollection.Count > 0
        chart.SeriesCollection(1).Delete
    Loop
    If outputRow = 5 Then GoTo SafeExit
    lastRow = outputRow - 1
    chart.ChartType = xlBarStacked
    With chart.SeriesCollection.NewSeries
        .Name = "Начало"
        .XValues = graph.Range("B" & FIRST_DATA_ROW & ":B" & lastRow)
        .Values = graph.Range("C" & FIRST_DATA_ROW & ":C" & lastRow)
        .Format.Fill.Visible = msoFalse
        .Format.Line.Visible = msoFalse
    End With
    With chart.SeriesCollection.NewSeries
        .Name = "КП — договорные сроки"
        .XValues = graph.Range("B5:B" & lastRow)
        .Values = graph.Range("E5:E" & lastRow)
        .Format.Fill.ForeColor.RGB = RGB(115, 64, 160)
        .Format.Line.ForeColor.RGB = RGB(33, 99, 191)
    End With
    With chart.SeriesCollection.NewSeries
        .Name = "СР — сроки"
        .XValues = graph.Range("B5:B" & lastRow)
        .Values = graph.Range("F5:F" & lastRow)
        .Format.Fill.ForeColor.RGB = RGB(245, 166, 35)
        .Format.Line.ForeColor.RGB = RGB(192, 115, 0)
    End With
    chart.HasTitle = True
    chart.ChartTitle.Text = "Диаграмма Ганта: КП и СР"
    chart.HasLegend = True
    chart.Legend.Position = xlLegendPositionBottom
    chart.Axes(xlCategory).ReversePlotOrder = True
    chart.Axes(xlValue).TickLabels.NumberFormat = "dd.mm.yy"
SafeExit:
End Sub
Public Sub GoToSR(): ThisWorkbook.Worksheets("03_СР_проекта").Activate: End Sub
Public Sub GoToGraph(): ThisWorkbook.Worksheets("05_График").Activate: End Sub
Public Sub GoToControl(): ThisWorkbook.Worksheets("06_Контроль").Activate: End Sub
