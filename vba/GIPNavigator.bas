Option Explicit

Private Const FIRST_DATA_ROW As Long = 5
Private Const LAST_DATA_ROW As Long = 79
Private Const MIN_STAGES As Long = 2
Private Const MAX_STAGES As Long = 75
Private Const SR_FIRST_ROW As Long = 5
Private Const SR_LAST_ROW As Long = 304
Private Const SR_NETWORK_COLUMN As Long = 13
Private Const SR_SOURCE_COLUMN As Long = 14
Private Const SR_CONSTRAINT_START_COLUMN As Long = 15

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
    Dim previousScreenUpdating As Boolean, previousEnableEvents As Boolean, applicationStateCaptured As Boolean
    Dim backupPath As String, existingCount As Long, preservedActuals As Long, previewText As String
    Dim oldProjectStart As Variant, newProjectStart As Date
    Dim beforeImport As Variant, writeStarted As Boolean, failureNumber As Long, failureDescription As String, targetWasProtected As Boolean

    On Error GoTo Fail
    previousScreenUpdating = Application.ScreenUpdating
    previousEnableEvents = Application.EnableEvents
    applicationStateCaptured = True
    If Len(Dir$(sourcePath)) = 0 Then Err.Raise vbObjectError + 100, , "Файл КП не найден."
    If LCase$(Right$(sourcePath, 4)) <> ".xls" And LCase$(Right$(sourcePath, 5)) <> ".xlsx" Then _
        Err.Raise vbObjectError + 101, , "Поддерживаются только файлы .xls и .xlsx."
    Set target = ThisWorkbook.Worksheets("01_КП")
    targetWasProtected = target.ProtectContents
    existingCount = Application.WorksheetFunction.CountA(target.Range("C" & FIRST_DATA_ROW & ":C" & LAST_DATA_ROW))
    oldProjectStart = ProjectStartDate()
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
    ThisWorkbook.SaveCopyAs backupPath
    beforeImport = target.Range("A" & FIRST_DATA_ROW & ":I" & LAST_DATA_ROW).Formula
    If targetWasProtected Then target.Unprotect
    writeStarted = True
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
    ' The shift is an Excel formula.  Calculate it before building the SR
    ' network and graph, so both always use the same shifted project start.
    target.Range("J3:M" & LAST_DATA_ROW).Calculate
    ThisWorkbook.Worksheets("99_Служебный").Range("B6").Value = Now
    ThisWorkbook.Worksheets("99_Служебный").Range("B6").NumberFormat = "dd.mm.yyyy hh:mm"
    RecalculateSRNetwork True
    RefreshGantt silent
    ' Persist calculated Gantt source values even when the user's Excel is in
    ' manual calculation mode.
    Application.CalculateFullRebuild
    If targetWasProtected Then target.Protect UserInterfaceOnly:=True, AllowFiltering:=True
    ' Automated verification is non-persistent; the user command always saves.
    If Not silent Then ThisWorkbook.Save
    If Not silent Then MsgBox "КП импортирован: " & stageCount & " этапов." & vbCrLf & "Резервная копия: " & backupPath, vbInformation, "GIP Navigator"
SafeExit:
    If targetWasProtected And Not target Is Nothing Then
        If Not target.ProtectContents Then target.Protect UserInterfaceOnly:=True, AllowFiltering:=True
    End If
    CloseImportSource sourceBook
    If applicationStateCaptured Then
        Application.ScreenUpdating = previousScreenUpdating
        Application.EnableEvents = previousEnableEvents
    End If
    Exit Sub
Fail:
    failureNumber = Err.Number
    failureDescription = Err.Description
    If writeStarted Then
        target.Range("A" & FIRST_DATA_ROW & ":I" & LAST_DATA_ROW).Formula = beforeImport
    End If
    If targetWasProtected And Not target Is Nothing Then
        If Not target.ProtectContents Then target.Protect UserInterfaceOnly:=True, AllowFiltering:=True
    End If
    CloseImportSource sourceBook
    If applicationStateCaptured Then
        Application.ScreenUpdating = previousScreenUpdating
        Application.EnableEvents = previousEnableEvents
    End If
    If silent Then
        RecordAutomationError "ImportKpFromPath", failureNumber, failureDescription
        Exit Sub
    End If
    MsgBox "Импорт не выполнен. Данные книги сохранены без изменений. " & failureDescription, vbExclamation, "GIP Navigator"
End Sub

Private Sub CloseImportSource(ByRef sourceBook As Workbook)
    If sourceBook Is Nothing Then Exit Sub
    On Error GoTo CloseFailed
    sourceBook.Close SaveChanges:=False
    Exit Sub
CloseFailed:
    Debug.Print "GIP Navigator: не удалось закрыть источник импорта: " & Err.Description
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
    Dim code As String, protectedCount As Long, oldScreenUpdating As Boolean, oldEnableEvents As Boolean, applicationStateCaptured As Boolean
    Dim failureNumber As Long, failureDescription As String
    On Error GoTo Fail
    oldScreenUpdating = Application.ScreenUpdating: oldEnableEvents = Application.EnableEvents
    applicationStateCaptured = True
    Set library = ThisWorkbook.Worksheets("02_Библиотека_СР"): Set projectSR = ThisWorkbook.Worksheets("03_СР_проекта")
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
    If applicationStateCaptured Then
        Application.ScreenUpdating = oldScreenUpdating
        Application.EnableEvents = oldEnableEvents
    End If
    Exit Sub
Fail:
    failureNumber = Err.Number: failureDescription = Err.Description
    If applicationStateCaptured Then
        Application.ScreenUpdating = oldScreenUpdating
        Application.EnableEvents = oldEnableEvents
    End If
    If silent Then
        RecordAutomationError "SyncSelectedSR", failureNumber, failureDescription
        Exit Sub
    End If
    MsgBox "Выбор СР не выполнен. " & failureDescription, vbExclamation, "GIP Navigator"
End Sub

Public Sub AddSRInstance()
    Dim code As String
    code = Trim$(InputBox("Введите код типовой СР (например, SR-001). Повторный код создаст отдельный экземпляр.", "Добавить экземпляр СР"))
    If Len(code) = 0 Then Exit Sub
    AddSRInstanceByCode code, False
End Sub

' Separate entry point keeps the button workflow simple and makes the same
' validation available to the reproducible Excel test.
Public Sub AddSRInstanceByCode(ByVal code As String, Optional ByVal silent As Boolean = False)
    Dim projectSR As Worksheet, rowNumber As Long, oldEnableEvents As Boolean, applicationStateCaptured As Boolean
    Dim failureNumber As Long, failureDescription As String
    On Error GoTo Fail
    oldEnableEvents = Application.EnableEvents: applicationStateCaptured = True
    code = Trim$(code)
    If Application.WorksheetFunction.CountIf(ThisWorkbook.Worksheets("02_Библиотека_СР").Range("B5:B85"), code) = 0 Then Err.Raise vbObjectError + 121, , "Код отсутствует в библиотеке СР: " & code
    Set projectSR = ThisWorkbook.Worksheets("03_СР_проекта"): rowNumber = FindFreeSRRow(projectSR)
    If rowNumber = 0 Then Err.Raise vbObjectError + 122, , "В 03_СР_проекта нет свободной строки."
    Application.EnableEvents = False
    projectSR.Cells(rowNumber, 1).Value = NextInstanceId(projectSR, code)
    projectSR.Cells(rowNumber, 2).Value = code
    projectSR.Cells(rowNumber, SR_SOURCE_COLUMN).Value = "GIP_LIBRARY"
    RecalculateSRNetwork True
    Application.EnableEvents = oldEnableEvents
    If Not silent Then projectSR.Activate: projectSR.Cells(rowNumber, 1).Select
    Exit Sub
Fail:
    failureNumber = Err.Number: failureDescription = Err.Description
    If applicationStateCaptured Then Application.EnableEvents = oldEnableEvents
    If silent Then
        RecordAutomationError "AddSRInstanceByCode", failureNumber, failureDescription
        Exit Sub
    End If
    MsgBox "Экземпляр СР не добавлен. " & failureDescription, vbExclamation, "GIP Navigator"
End Sub

' Inclusive dates. FS means the next date after predecessor finish; SS means
' predecessor start. A lag is counted in the calendar of the successor.
Public Sub RecalculateSRNetwork(Optional ByVal silent As Boolean = False)
    Dim sr As Worksheet, links As Worksheet, rows As Object, indegree As Object, successors As Object, predecessors As Object
    Dim queue As Collection, order As Collection, itemId As Variant, r As Long, linkRow As Long
    Dim successorId As Variant, predecessorId As String, relation As String
    Dim projectStart As Variant, duration As Variant, ownStart As Variant, calculatedStart As Date, calculatedEnd As Date
    Dim candidates As Collection, linkData As Variant, candidate As Variant, maxDate As Date, hasCandidate As Boolean, oldEnableEvents As Boolean, applicationStateCaptured As Boolean
    Dim failureNumber As Long, failureDescription As String
    On Error GoTo Fail
    oldEnableEvents = Application.EnableEvents: applicationStateCaptured = True
    Set sr = ThisWorkbook.Worksheets("03_СР_проекта"): Set links = ThisWorkbook.Worksheets("04_Согласование_плана")
    projectStart = ProjectStartDate(): Application.EnableEvents = False
    Set rows = CreateObject("Scripting.Dictionary"): Set indegree = CreateObject("Scripting.Dictionary")
    Set successors = CreateObject("Scripting.Dictionary"): Set predecessors = CreateObject("Scripting.Dictionary")
    For r = SR_FIRST_ROW To SR_LAST_ROW
        If Len(Trim$(CStr(sr.Cells(r, 1).Value))) > 0 Then
            itemId = Trim$(CStr(sr.Cells(r, 1).Value))
            sr.Cells(r, SR_NETWORK_COLUMN).ClearContents
            sr.Cells(r, 4).ClearContents: sr.Cells(r, 6).ClearContents
            sr.Cells(r, 10).ClearContents: sr.Cells(r, 11).ClearContents
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
        ' Each SR needs its own maximum candidate.  Carrying this state from a
        ' preceding row makes independent work inherit the previous deadline.
        hasCandidate = False
        maxDate = 0
        r = CLng(rows(itemId)): duration = sr.Cells(r, 5).Value
        ' VBA does not short-circuit Or.  Test a blank/unknown duration before
        ' converting it, so a library value "уточнить" is diagnosed, never
        ' converted to a zero-day task or raised as a runtime error.
        If Not IsNumeric(duration) Then
            AddSRIssue sr, r, "Не задана длительность СР."
            GoTo NextItem
        End If
        If CLng(duration) < 0 Then
            AddSRIssue sr, r, "Не задана длительность СР."
            GoTo NextItem
        End If
        ownStart = sr.Cells(r, SR_CONSTRAINT_START_COLUMN).Value
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
        calculatedStart = maxDate
        ' A zero-duration SR is a milestone: it starts and finishes on one date.
        If CLng(duration) = 0 Then
            calculatedEnd = calculatedStart
        Else
            calculatedEnd = DateAdd("d", CLng(duration) - 1, calculatedStart)
        End If
        sr.Cells(r, 4).Value = calculatedStart: sr.Cells(r, 6).Value = calculatedEnd
        calculated.Add CStr(itemId), Array(calculatedStart, calculatedEnd)
        UpdateSRForecast sr, r, calculatedStart, calculatedEnd
NextItem:
    Next itemId
SafeExit:
    RecalculateKpObligations sr
    RefreshControl sr
    RefreshLibraryGantt silent
    RefreshGantt silent
    If applicationStateCaptured Then Application.EnableEvents = oldEnableEvents
    Exit Sub
Fail:
    failureNumber = Err.Number: failureDescription = Err.Description
    If applicationStateCaptured Then Application.EnableEvents = oldEnableEvents
    If silent Then
        RecordAutomationError "RecalculateSRNetwork", failureNumber, failureDescription
        Exit Sub
    End If
    MsgBox "Расчёт сети СР не выполнен. " & failureDescription, vbExclamation, "GIP Navigator"
End Sub

' Calculates the only project-level forecast used by the graph and control.
' This version deliberately has no КП -> СР links, so an СР -> КП rule cannot
' create a circular network with the existing SR -> SR dependencies.
Public Sub RecalculateKpObligations(Optional ByVal suppliedSR As Worksheet)
    Dim kp As Worksheet, rules As Worksheet, sr As Worksheet
    Dim r As Long, stageRow As Long, srRow As Long, ruleStage As String, ruleId As String, ruleKind As String
    Dim srEnd As Variant, deadline As Date, stageStart As Date, stageEnd As Date, candidateStart As Date, durationDays As Long
    Dim message As String, hasRule As Boolean, reserveDays As Long
    On Error GoTo Fail
    Set kp = ThisWorkbook.Worksheets("01_КП")
    Set rules = ThisWorkbook.Worksheets("04_Согласование_плана")
    If suppliedSR Is Nothing Then Set sr = ThisWorkbook.Worksheets("03_СР_проекта") Else Set sr = suppliedSR
    Application.EnableEvents = False
    kp.Range("N5:R79").ClearContents
    kp.Range("V5:V79").ClearContents
    rules.Range("N5:O304").ClearContents
    For r = FIRST_DATA_ROW To LAST_DATA_ROW
        If Len(Trim$(CStr(kp.Cells(r, 3).Value))) > 0 And IsDate(kp.Cells(r, 10).Value) And IsDate(kp.Cells(r, 11).Value) Then
            kp.Cells(r, 14).Value = CDate(kp.Cells(r, 10).Value)
            kp.Cells(r, 15).Value = CDate(kp.Cells(r, 11).Value)
            kp.Cells(r, 16).Value = 0
            kp.Cells(r, 17).Value = "Нормально"
        End If
    Next r
    For r = SR_FIRST_ROW To SR_LAST_ROW
        ruleId = Trim$(CStr(rules.Cells(r, 9).Value))
        ruleStage = StableStageId(rules.Cells(r, 11).Value)
        ruleKind = Trim$(CStr(rules.Cells(r, 13).Value))
        If Len(ruleId) = 0 And Len(ruleStage) = 0 Then GoTo NextRule
        If Len(ruleId) = 0 Or Len(ruleStage) = 0 Or Len(ruleKind) = 0 Then
            rules.Cells(r, 14).Value = "Недостаточно данных": rules.Cells(r, 15).Value = "Заполните СР, этап КП и условие."
            GoTo NextRule
        End If
        stageRow = FindKpStageRow(kp, ruleStage): srRow = FindSRInstanceRow(sr, ruleId)
        If stageRow = 0 Or srRow = 0 Then
            rules.Cells(r, 14).Value = "Недостаточно данных": rules.Cells(r, 15).Value = "Этап КП или выбранный экземпляр СР не найден."
            If stageRow > 0 Then MarkKpInsufficient kp, stageRow, "Обязательная СР не найдена: " & ruleId
            GoTo NextRule
        End If
        If Not IsDate(kp.Cells(stageRow, 14).Value) Or Not IsDate(kp.Cells(stageRow, 15).Value) Then
            rules.Cells(r, 14).Value = "Недостаточно данных": rules.Cells(r, 15).Value = "Не определены смещённые сроки этапа КП."
            GoTo NextRule
        End If
        srEnd = EffectiveSREnd(sr, srRow)
        If Not IsDate(srEnd) Then
            rules.Cells(r, 14).Value = "Недостаточно данных": rules.Cells(r, 15).Value = "Не определён срок обязательной СР."
            MarkKpInsufficient kp, stageRow, "Не определён срок СР: " & sr.Cells(srRow, 2).Value & " — " & sr.Cells(srRow, 3).Value
            GoTo NextRule
        End If
        stageStart = CDate(kp.Cells(stageRow, 14).Value): stageEnd = CDate(kp.Cells(stageRow, 15).Value)
        message = BuildSRReason(sr, rules, srRow, CDate(srEnd))
        If ruleKind = "До начала этапа" Then
            deadline = stageStart
            If CDate(srEnd) >= deadline Then
                durationDays = DateDiff("d", stageStart, stageEnd) + 1
                candidateStart = DateAdd("d", 1, CDate(srEnd))
                If candidateStart > stageStart Then
                    stageStart = candidateStart: stageEnd = DateAdd("d", durationDays - 1, stageStart)
                    kp.Cells(stageRow, 14).Value = stageStart: kp.Cells(stageRow, 15).Value = stageEnd
                End If
                MarkKpCritical kp, stageRow, message
                rules.Cells(r, 14).Value = "Критично": rules.Cells(r, 15).Value = message
            Else
                reserveDays = DateDiff("d", CDate(srEnd), deadline) - 1
                MarkKpReserve kp, stageRow, reserveDays, message
                rules.Cells(r, 14).Value = IIf(reserveDays <= 7, "Предупреждение", "Нормально"): rules.Cells(r, 15).Value = message
            End If
        ElseIf ruleKind = "До окончания этапа" Then
            deadline = stageEnd
            If CDate(srEnd) > deadline Then
                kp.Cells(stageRow, 15).Value = CDate(srEnd)
                MarkKpCritical kp, stageRow, message
                rules.Cells(r, 14).Value = "Критично": rules.Cells(r, 15).Value = message
            Else
                reserveDays = DateDiff("d", CDate(srEnd), deadline)
                MarkKpReserve kp, stageRow, reserveDays, message
                rules.Cells(r, 14).Value = IIf(reserveDays <= 7, "Предупреждение", "Нормально"): rules.Cells(r, 15).Value = message
            End If
        Else
            rules.Cells(r, 14).Value = "Недостаточно данных": rules.Cells(r, 15).Value = "Неизвестное условие обязательной связи."
            MarkKpInsufficient kp, stageRow, "Неизвестное условие для СР: " & ruleId
        End If
NextRule:
    Next r
    RecalculateActs
    For r = FIRST_DATA_ROW To LAST_DATA_ROW
        If Len(Trim$(CStr(kp.Cells(r, 3).Value))) > 0 And IsDate(kp.Cells(r, 11).Value) And IsDate(kp.Cells(r, 15).Value) Then
            kp.Cells(r, 16).Value = DateDiff("d", CDate(kp.Cells(r, 11).Value), CDate(kp.Cells(r, 15).Value))
            If CStr(kp.Cells(r, 17).Value) = "Критично" Then kp.Cells(r, 18).Value = kp.Cells(r, 20).Value Else kp.Cells(r, 18).Value = 0
        End If
    Next r
    Exit Sub
Fail:
    RecordAutomationError "RecalculateKpObligations", Err.Number, Err.Description
End Sub

Private Function FindKpStageRow(ByVal kp As Worksheet, ByVal stageId As String) As Long
    Dim r As Long
    For r = FIRST_DATA_ROW To LAST_DATA_ROW
        If StableStageId(kp.Cells(r, 3).Value) = stageId Then FindKpStageRow = r: Exit Function
    Next r
End Function

Private Function FindSRInstanceRow(ByVal sr As Worksheet, ByVal instanceId As String) As Long
    Dim r As Long
    For r = SR_FIRST_ROW To SR_LAST_ROW
        If Trim$(CStr(sr.Cells(r, 1).Value)) = instanceId Then FindSRInstanceRow = r: Exit Function
    Next r
End Function

Private Function EffectiveSREnd(ByVal sr As Worksheet, ByVal rowNumber As Long) As Variant
    If IsDate(sr.Cells(rowNumber, 8).Value) Then EffectiveSREnd = sr.Cells(rowNumber, 8).Value: Exit Function
    If IsDate(sr.Cells(rowNumber, 10).Value) Then EffectiveSREnd = sr.Cells(rowNumber, 10).Value: Exit Function
    If IsDate(sr.Cells(rowNumber, 6).Value) Then EffectiveSREnd = sr.Cells(rowNumber, 6).Value
End Function

Private Function BuildSRReason(ByVal sr As Worksheet, ByVal rules As Worksheet, ByVal srRow As Long, ByVal finishDate As Date) As String
    Dim r As Long, predecessor As Long, itemId As String, chain As String
    itemId = Trim$(CStr(sr.Cells(srRow, 1).Value))
    For r = SR_FIRST_ROW To SR_LAST_ROW
        If Trim$(CStr(rules.Cells(r, 2).Value)) = itemId Then
            predecessor = FindSRInstanceRow(sr, Trim$(CStr(rules.Cells(r, 3).Value)))
            If predecessor > 0 Then chain = chain & IIf(Len(chain) > 0, "; ", "") & sr.Cells(predecessor, 2).Value
        End If
    Next r
    BuildSRReason = sr.Cells(srRow, 2).Value & " — " & sr.Cells(srRow, 3).Value & "; окончание " & Format$(finishDate, "dd.mm.yyyy")
    If Len(chain) > 0 Then BuildSRReason = BuildSRReason & "; предшественники: " & chain
End Function

Private Sub MarkKpCritical(ByVal kp As Worksheet, ByVal stageRow As Long, ByVal message As String)
    kp.Cells(stageRow, 17).Value = "Критично": kp.Cells(stageRow, 22).Value = message
End Sub
Private Sub MarkKpInsufficient(ByVal kp As Worksheet, ByVal stageRow As Long, ByVal message As String)
    If CStr(kp.Cells(stageRow, 17).Value) <> "Критично" Then kp.Cells(stageRow, 17).Value = "Недостаточно данных"
    kp.Cells(stageRow, 22).Value = message
End Sub
Private Sub MarkKpReserve(ByVal kp As Worksheet, ByVal stageRow As Long, ByVal reserveDays As Long, ByVal message As String)
    If CStr(kp.Cells(stageRow, 17).Value) = "Критично" Or CStr(kp.Cells(stageRow, 17).Value) = "Недостаточно данных" Then Exit Sub
    If reserveDays <= 7 Then
        kp.Cells(stageRow, 17).Value = "Предупреждение": kp.Cells(stageRow, 22).Value = message
    Else
        kp.Cells(stageRow, 17).Value = "Нормально"
    End If
End Sub

Public Sub RecalculateActs()
    Dim acts As Worksheet, kp As Worksheet, totals As Object, latest As Object, seen As Object
    Dim r As Long, stageId As String, rowNumber As Long, amount As Currency, key As String
    Dim oldEnableEvents As Boolean, eventsCaptured As Boolean
    On Error GoTo Fail
    oldEnableEvents = Application.EnableEvents: eventsCaptured = True: Application.EnableEvents = False
    Set acts = ThisWorkbook.Worksheets("07_Актирование"): Set kp = ThisWorkbook.Worksheets("01_КП")
    Set totals = CreateObject("Scripting.Dictionary"): Set latest = CreateObject("Scripting.Dictionary"): Set seen = CreateObject("Scripting.Dictionary")
    acts.Range("H5:H304").ClearContents
    For r = SR_FIRST_ROW To SR_LAST_ROW
        stageId = StableStageId(acts.Cells(r, 2).Value)
        If Len(stageId) = 0 Then GoTo NextAct
        If FindKpStageRow(kp, stageId) = 0 Then acts.Cells(r, 8).Value = "Этап КП не найден": GoTo NextAct
        If Trim$(CStr(acts.Cells(r, 7).Value)) = "Подписан" Then
            If Not IsDate(acts.Cells(r, 4).Value) Or Not IsNumeric(acts.Cells(r, 6).Value) Then
                acts.Cells(r, 8).Value = "Для подписанного акта нужны дата и сумма": GoTo NextAct
            End If
            If CCur(acts.Cells(r, 6).Value) <= 0 Then acts.Cells(r, 8).Value = "Для подписанного акта нужна положительная сумма": GoTo NextAct
            key = stageId & "|" & Trim$(CStr(acts.Cells(r, 5).Value))
            If Len(Trim$(CStr(acts.Cells(r, 5).Value))) > 0 And seen.Exists(key) Then acts.Cells(r, 8).Value = "Дубликат номера акта": GoTo NextAct
            If Len(Trim$(CStr(acts.Cells(r, 5).Value))) > 0 Then seen.Add key, True
            amount = CCur(acts.Cells(r, 6).Value)
            If Not totals.Exists(stageId) Then totals.Add stageId, amount: latest.Add stageId, CDate(acts.Cells(r, 4).Value) Else totals(stageId) = CCur(totals(stageId)) + amount: If CDate(acts.Cells(r, 4).Value) > CDate(latest(stageId)) Then latest(stageId) = CDate(acts.Cells(r, 4).Value)
        End If
NextAct:
    Next r
    For r = FIRST_DATA_ROW To LAST_DATA_ROW
        stageId = StableStageId(kp.Cells(r, 3).Value)
        If Len(stageId) > 0 And totals.Exists(stageId) Then
            kp.Cells(r, 19).Value = totals(stageId)
            kp.Cells(r, 8).Value = latest(stageId): kp.Cells(r, 9).Value = totals(stageId)
            If IsNumeric(kp.Cells(r, 7).Value) And CCur(totals(stageId)) > CCur(kp.Cells(r, 7).Value) Then MarkActsOverrun acts, stageId
        Else
            kp.Cells(r, 19).ClearContents
        End If
    Next r
    If eventsCaptured Then Application.EnableEvents = oldEnableEvents
    Exit Sub
Fail:
    If eventsCaptured Then Application.EnableEvents = oldEnableEvents
    RecordAutomationError "RecalculateActs", Err.Number, Err.Description
End Sub

Private Sub MarkActsOverrun(ByVal acts As Worksheet, ByVal stageId As String)
    Dim r As Long
    For r = SR_FIRST_ROW To SR_LAST_ROW
        If StableStageId(acts.Cells(r, 2).Value) = stageId And Trim$(CStr(acts.Cells(r, 7).Value)) = "Подписан" Then acts.Cells(r, 8).Value = "Сумма актов превышает стоимость этапа"
    Next r
End Sub

Private Function ProjectStartDate() As Variant
    Dim values As Variant, r As Long, found As Boolean, result As Date
    ' Project planning always starts from the shifted contractual schedule.
    ' Columns E:F retain the immutable dates imported from the approved KP.
    values = ThisWorkbook.Worksheets("01_КП").Range("J5:J79").Value
    For r = 1 To UBound(values, 1)
        If IsDate(values(r, 1)) Then If Not found Or CDate(values(r, 1)) < result Then result = CDate(values(r, 1)): found = True
    Next r
    If found Then ProjectStartDate = result
End Function


Private Sub AddSRIssue(ByVal sr As Worksheet, ByVal rowNumber As Long, ByVal issueText As String)
    If Len(sr.Cells(rowNumber, SR_NETWORK_COLUMN).Value) > 0 Then sr.Cells(rowNumber, SR_NETWORK_COLUMN).Value = sr.Cells(rowNumber, SR_NETWORK_COLUMN).Value & " "
    sr.Cells(rowNumber, SR_NETWORK_COLUMN).Value = sr.Cells(rowNumber, SR_NETWORK_COLUMN).Value & issueText
    sr.Cells(rowNumber, 11).Value = "Ошибка данных"
End Sub

Private Sub UpdateSRForecast(ByVal sr As Worksheet, ByVal rowNumber As Long, ByVal calculatedStart As Date, ByVal calculatedEnd As Date)
    Dim actualStart As Variant, actualEnd As Variant, expectedEnd As Variant
    actualStart = sr.Cells(rowNumber, 7).Value
    actualEnd = sr.Cells(rowNumber, 8).Value
    expectedEnd = sr.Cells(rowNumber, 9).Value
    sr.Cells(rowNumber, 10).ClearContents

    If IsDate(actualEnd) And Not IsDate(actualStart) Then AddSRIssue sr, rowNumber, "Указано фактическое окончание без фактического начала."
    If IsDate(actualStart) And IsDate(actualEnd) And CDate(actualEnd) < CDate(actualStart) Then AddSRIssue sr, rowNumber, "Фактическое окончание раньше фактического начала."
    If IsDate(actualStart) And IsDate(expectedEnd) And CDate(expectedEnd) < CDate(actualStart) Then AddSRIssue sr, rowNumber, "Ожидаемое окончание раньше фактического начала."
    If Len(Trim$(CStr(sr.Cells(rowNumber, SR_NETWORK_COLUMN).Value))) > 0 Then
        sr.Cells(rowNumber, 11).Value = "Ошибка данных"
        Exit Sub
    End If

    If IsDate(actualEnd) Then
        sr.Cells(rowNumber, 10).Value = CDate(actualEnd)
        sr.Cells(rowNumber, 11).Value = "Завершена"
    ElseIf IsDate(actualStart) Then
        If IsDate(expectedEnd) Then
            sr.Cells(rowNumber, 10).Value = CDate(expectedEnd)
            If CDate(actualStart) > calculatedStart Or CDate(expectedEnd) > calculatedEnd Then
                sr.Cells(rowNumber, 11).Value = "В работе — отставание от плана"
            Else
                sr.Cells(rowNumber, 11).Value = "В работе"
            End If
        Else
            sr.Cells(rowNumber, 11).Value = "В работе — прогноз не определён"
        End If
    ElseIf IsDate(expectedEnd) Then
        sr.Cells(rowNumber, 10).Value = CDate(expectedEnd)
        sr.Cells(rowNumber, 11).Value = "Не начата — прогноз задан"
    Else
        sr.Cells(rowNumber, 11).Value = "Не начата"
    End If
End Sub

Public Sub RefreshControl(ByVal sr As Worksheet)
    Dim control As Worksheet, kp As Worksheet, r As Long, outputRow As Long, issueText As String, statusText As String, stateText As String
    Set control = ThisWorkbook.Worksheets("06_Контроль"): Set kp = ThisWorkbook.Worksheets("01_КП")
    control.Range("A5:H304").ClearContents: outputRow = 5
    For r = FIRST_DATA_ROW To LAST_DATA_ROW
        stateText = Trim$(CStr(kp.Cells(r, 17).Value))
        If Len(Trim$(CStr(kp.Cells(r, 3).Value))) > 0 And (stateText = "Критично" Or stateText = "Предупреждение" Or stateText = "Недостаточно данных") Then
            control.Cells(outputRow, 1).Value = "КП: обязательные СР"
            control.Cells(outputRow, 2).Value = kp.Cells(r, 3).Value
            control.Cells(outputRow, 3).Value = stateText
            control.Cells(outputRow, 4).Value = kp.Cells(r, 22).Value
            control.Cells(outputRow, 5).Value = kp.Cells(r, 11).Value
            control.Cells(outputRow, 6).Value = kp.Cells(r, 15).Value
            control.Cells(outputRow, 7).Value = kp.Cells(r, 16).Value
            If stateText = "Критично" Then control.Cells(outputRow, 8).Value = "Проверьте критическую СР, согласуйте перенос этапа и при необходимости уведомьте Заказчика."
            If stateText = "Предупреждение" Then control.Cells(outputRow, 8).Value = "Контролируйте резерв сроков и уточните ожидаемое окончание СР."
            If stateText = "Недостаточно данных" Then control.Cells(outputRow, 8).Value = "Уточните длительность, прогноз или обязательную связь СР."
            outputRow = outputRow + 1
        End If
    Next r
    For r = SR_FIRST_ROW To SR_LAST_ROW
        If Len(Trim$(CStr(sr.Cells(r, 1).Value))) > 0 Then
            issueText = Trim$(CStr(sr.Cells(r, SR_NETWORK_COLUMN).Value)): statusText = Trim$(CStr(sr.Cells(r, 11).Value))
            If Len(issueText) > 0 Or InStr(1, statusText, "прогноз не определён", vbTextCompare) > 0 Then
                control.Cells(outputRow, 1).Value = "СР: данные"
                control.Cells(outputRow, 3).Value = "Недостаточно данных"
                control.Cells(outputRow, 4).Value = sr.Cells(r, 2).Value & " — " & sr.Cells(r, 3).Value
                control.Cells(outputRow, 8).Value = IIf(Len(issueText) > 0, issueText, "Введите ожидаемое окончание или фактическое окончание работы.")
                outputRow = outputRow + 1
            End If
        End If
    Next r
End Sub

Public Sub ApproveBaseline(Optional ByVal silent As Boolean = False)
    Dim agreement As Worksheet, kp As Worksheet, control As Worksheet
    Dim revision As String, responsible As String, approvalDate As Variant, r As Long, outputRow As Long
    On Error GoTo Fail
    Set agreement = ThisWorkbook.Worksheets("04_Согласование_плана")
    Set kp = ThisWorkbook.Worksheets("01_КП"): Set control = ThisWorkbook.Worksheets("06_Контроль")
    revision = Trim$(CStr(agreement.Range("R3").Value)): approvalDate = agreement.Range("R4").Value: responsible = Trim$(CStr(agreement.Range("R5").Value))
    If Not IsDate(kp.Range("B3").Value) Then Err.Raise vbObjectError + 160, , "Не введена дата подписания наряд-заказа."
    If Len(revision) = 0 Or Not IsDate(approvalDate) Or Len(responsible) = 0 Then Err.Raise vbObjectError + 161, , "Заполните версию, дату согласования и ответственного на листе 04."
    For r = SR_FIRST_ROW To SR_LAST_ROW
        If Len(Trim$(CStr(agreement.Cells(r, 9).Value))) > 0 And Len(Trim$(CStr(agreement.Cells(r, 11).Value))) > 0 And Len(Trim$(CStr(agreement.Cells(r, 13).Value))) > 0 Then
            If Trim$(CStr(agreement.Cells(r, 14).Value)) = "Недостаточно данных" Then Err.Raise vbObjectError + 162, , "Есть незаполненные обязательные данные по обязательным СР. Согласование невозможно."
        End If
    Next r
    If Application.WorksheetFunction.CountIf(ThisWorkbook.Worksheets("03_СР_проекта").Range("M5:M304"), "*цикл*") > 0 Then Err.Raise vbObjectError + 163, , "В сети СР обнаружен цикл. Согласование невозможно."
    outputRow = agreement.Cells(agreement.Rows.Count, 17).End(xlUp).Row + 1
    If outputRow < 10 Then outputRow = 10
    For r = FIRST_DATA_ROW To LAST_DATA_ROW
        If Len(Trim$(CStr(kp.Cells(r, 3).Value))) > 0 Then
            agreement.Cells(outputRow, 17).Value = revision
            agreement.Cells(outputRow, 18).Value = CDate(approvalDate)
            agreement.Cells(outputRow, 19).Value = kp.Cells(r, 3).Value
            agreement.Cells(outputRow, 20).Value = kp.Cells(r, 4).Value
            agreement.Cells(outputRow, 21).Value = kp.Cells(r, 14).Value
            agreement.Cells(outputRow, 22).Value = kp.Cells(r, 15).Value
            outputRow = outputRow + 1
        End If
    Next r
    agreement.Range("R6").Value = "Согласовано: версия " & revision & "; " & Format$(CDate(approvalDate), "dd.mm.yyyy") & "; " & responsible
    ThisWorkbook.Worksheets("99_Служебный").Range("B8").Value = revision
    If Not silent Then ThisWorkbook.Save: MsgBox "Снимок согласованного графика сохранён: версия " & revision & ".", vbInformation, "GIP Navigator"
    Exit Sub
Fail:
    If silent Then RecordAutomationError "ApproveBaseline", Err.Number, Err.Description Else MsgBox "Согласование не зафиксировано. " & Err.Description, vbExclamation, "GIP Navigator"
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
    ' Calculated start/end, formula defaults and technical fields do not make a
    ' selected library row user-owned. Facts, status/comment, a duration
    ' override, an explicit constraint or a live SR-to-SR link do.
    For Each columnNumber In Array(7, 8, 9, 12, SR_CONSTRAINT_START_COLUMN)
        If Len(Trim$(CStr(projectSR.Cells(rowNumber, CLng(columnNumber)).Value))) > 0 Then SRRowHasProjectData = True: Exit Function
    Next columnNumber
    If Len(Trim$(CStr(projectSR.Cells(rowNumber, 5).Value))) > 0 And Not projectSR.Cells(rowNumber, 5).HasFormula Then SRRowHasProjectData = True: Exit Function
    SRRowHasProjectData = HasActiveSRDependency(Trim$(CStr(projectSR.Cells(rowNumber, 1).Value)))
End Function

Private Function HasActiveSRDependency(ByVal instanceId As String) As Boolean
    Dim links As Worksheet, r As Long
    If Len(instanceId) = 0 Then Exit Function
    Set links = ThisWorkbook.Worksheets("04_Согласование_плана")
    For r = SR_FIRST_ROW To SR_LAST_ROW
        If Trim$(CStr(links.Cells(r, 2).Value)) = instanceId Or Trim$(CStr(links.Cells(r, 3).Value)) = instanceId Then
            HasActiveSRDependency = True
            Exit Function
        End If
    Next r
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
    value = Application.WorksheetFunction.Trim(value)
    value = Replace(value, " / ", "/"): value = Replace(value, " /", "/"): value = Replace(value, "/ ", "/")
    NormalizeText = LCase$(value)
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

Public Sub RefreshLibraryGantt(Optional ByVal silent As Boolean = False)
    Dim library As Worksheet, sr As Worksheet, chartObject As ChartObject, chart As Chart
    Dim r As Long, outputRow As Long, lastRow As Long, startDate As Variant, endDate As Variant
    Dim failureNumber As Long, failureDescription As String
    On Error GoTo Fail
    Set library = ThisWorkbook.Worksheets("02_Библиотека_СР")
    Set sr = ThisWorkbook.Worksheets("03_СР_проекта")
    library.Range("M5:P304").ClearContents
    outputRow = 5
    For r = SR_FIRST_ROW To SR_LAST_ROW
        startDate = sr.Cells(r, 4).Value: endDate = sr.Cells(r, 6).Value
        If Len(Trim$(CStr(sr.Cells(r, 1).Value))) > 0 And IsDate(startDate) And IsDate(endDate) Then
            library.Cells(outputRow, 13).Value = sr.Cells(r, 1).Value
            library.Cells(outputRow, 14).Value = sr.Cells(r, 3).Value
            library.Cells(outputRow, 15).Value = CDate(startDate)
            library.Cells(outputRow, 16).Value = DateDiff("d", CDate(startDate), CDate(endDate)) + 1
            outputRow = outputRow + 1
        End If
    Next r
    If library.ChartObjects.Count = 0 Then
        Set chartObject = library.ChartObjects.Add(library.Range("H4").Left, library.Range("H4").Top, 900, 550)
        chartObject.Name = "GIP_Library_Gantt"
    Else
        Set chartObject = library.ChartObjects("GIP_Library_Gantt")
    End If
    chartObject.Left = library.Range("H4").Left: chartObject.Top = library.Range("H4").Top
    chartObject.Width = 900: chartObject.Height = 550
    Set chart = chartObject.Chart
    Do While chart.SeriesCollection.Count > 0
        chart.SeriesCollection(1).Delete
    Loop
    If outputRow = 5 Then GoTo SafeExit
    lastRow = outputRow - 1
    chart.ChartType = xlBarStacked
    ' The Gantt source is intentionally hidden from the user.  Excel otherwise
    ' omits every series when it is configured to plot visible cells only.
    chart.PlotVisibleOnly = False
    With chart.SeriesCollection.NewSeries
        .Name = "Начало"
        .XValues = library.Range("N5:N" & lastRow)
        .Values = library.Range("O5:O" & lastRow)
        .Format.Fill.Visible = msoFalse
        .Format.Line.Visible = msoFalse
    End With
    With chart.SeriesCollection.NewSeries
        .Name = "СР — расчётные сроки"
        .XValues = library.Range("N5:N" & lastRow)
        .Values = library.Range("P5:P" & lastRow)
        .Format.Fill.ForeColor.RGB = RGB(12, 155, 214)
        .Format.Line.ForeColor.RGB = RGB(18, 90, 182)
    End With
    chart.HasTitle = True
    chart.ChartTitle.Text = "Диаграмма Ганта: выбранные СР"
    chart.HasLegend = True
    chart.Legend.Position = xlLegendPositionBottom
    chart.Axes(xlCategory).ReversePlotOrder = True
    chart.Axes(xlValue).TickLabels.NumberFormat = "dd.mm.yy"
SafeExit:
    Exit Sub
Fail:
    failureNumber = Err.Number: failureDescription = Err.Description
    If silent Then
        RecordAutomationError "RefreshLibraryGantt", failureNumber, failureDescription
        Exit Sub
    End If
    MsgBox "Не удалось обновить диаграмму библиотеки СР. " & failureDescription, vbExclamation, "GIP Navigator"
End Sub

Public Sub RefreshGantt(Optional ByVal silent As Boolean = False)
    Dim graph As Worksheet, kp As Worksheet, sr As Worksheet, chartObject As ChartObject, chart As Chart
    Dim sourceRow As Long, outputRow As Long, lastRow As Long, startDate As Variant, endDate As Variant
    Dim failureNumber As Long, failureDescription As String
    On Error GoTo Fail
    Set graph = ThisWorkbook.Worksheets("05_График")
    Set kp = ThisWorkbook.Worksheets("01_КП")
    Set sr = ThisWorkbook.Worksheets("03_СР_проекта")
    graph.Range("A5:I379").ClearContents
    graph.Rows("5:379").Hidden = False
    outputRow = 5
    For sourceRow = FIRST_DATA_ROW To LAST_DATA_ROW
        If Len(Trim$(CStr(kp.Cells(sourceRow, 3).Value))) > 0 And IsDate(kp.Cells(sourceRow, 10).Value) And IsDate(kp.Cells(sourceRow, 11).Value) Then
            graph.Cells(outputRow, 1).Value = "КП — смещённый план"
            graph.Cells(outputRow, 2).Value = kp.Cells(sourceRow, 3).Value & " — " & kp.Cells(sourceRow, 4).Value
            graph.Cells(outputRow, 3).Value = CDate(kp.Cells(sourceRow, 10).Value)
            graph.Cells(outputRow, 4).Value = DateDiff("d", CDate(kp.Cells(sourceRow, 10).Value), CDate(kp.Cells(sourceRow, 11).Value)) + 1
            graph.Cells(outputRow, 5).Value = graph.Cells(outputRow, 4).Value
            outputRow = outputRow + 1
            If IsDate(kp.Cells(sourceRow, 14).Value) And IsDate(kp.Cells(sourceRow, 15).Value) Then
                graph.Cells(outputRow, 1).Value = "КП — прогноз"
                graph.Cells(outputRow, 2).Value = kp.Cells(sourceRow, 3).Value & " — " & kp.Cells(sourceRow, 4).Value
                graph.Cells(outputRow, 3).Value = CDate(kp.Cells(sourceRow, 14).Value)
                graph.Cells(outputRow, 4).Value = DateDiff("d", CDate(kp.Cells(sourceRow, 14).Value), CDate(kp.Cells(sourceRow, 15).Value)) + 1
                graph.Cells(outputRow, 6).Value = graph.Cells(outputRow, 4).Value
                outputRow = outputRow + 1
            End If
        End If
    Next sourceRow
    For sourceRow = SR_FIRST_ROW To SR_LAST_ROW
        startDate = sr.Cells(sourceRow, 4).Value: endDate = sr.Cells(sourceRow, 6).Value
        If Len(Trim$(CStr(sr.Cells(sourceRow, 1).Value))) > 0 And IsDate(startDate) And IsDate(endDate) Then
            graph.Cells(outputRow, 1).Value = "СР — план"
            graph.Cells(outputRow, 2).Value = sr.Cells(sourceRow, 2).Value & " — " & sr.Cells(sourceRow, 3).Value
            graph.Cells(outputRow, 3).Value = CDate(startDate)
            graph.Cells(outputRow, 4).Value = DateDiff("d", CDate(startDate), CDate(endDate)) + 1
            graph.Cells(outputRow, 7).Value = graph.Cells(outputRow, 4).Value
            outputRow = outputRow + 1
            If IsDate(sr.Cells(sourceRow, 10).Value) Then
                graph.Cells(outputRow, 1).Value = "СР — прогноз"
                graph.Cells(outputRow, 2).Value = sr.Cells(sourceRow, 2).Value & " — " & sr.Cells(sourceRow, 3).Value
                If IsDate(sr.Cells(sourceRow, 7).Value) Then graph.Cells(outputRow, 3).Value = sr.Cells(sourceRow, 7).Value Else graph.Cells(outputRow, 3).Value = CDate(startDate)
                graph.Cells(outputRow, 4).Value = DateDiff("d", CDate(graph.Cells(outputRow, 3).Value), CDate(sr.Cells(sourceRow, 10).Value)) + 1
                graph.Cells(outputRow, 8).Value = graph.Cells(outputRow, 4).Value
                outputRow = outputRow + 1
            End If
            If IsDate(sr.Cells(sourceRow, 7).Value) And IsDate(sr.Cells(sourceRow, 8).Value) Then
                graph.Cells(outputRow, 1).Value = "СР — факт"
                graph.Cells(outputRow, 2).Value = sr.Cells(sourceRow, 2).Value & " — " & sr.Cells(sourceRow, 3).Value
                graph.Cells(outputRow, 3).Value = sr.Cells(sourceRow, 7).Value
                graph.Cells(outputRow, 4).Value = DateDiff("d", CDate(sr.Cells(sourceRow, 7).Value), CDate(sr.Cells(sourceRow, 8).Value)) + 1
                graph.Cells(outputRow, 9).Value = graph.Cells(outputRow, 4).Value
                outputRow = outputRow + 1
            End If
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
        .Name = "КП — смещённый план"
        .XValues = graph.Range("B5:B" & lastRow)
        .Values = graph.Range("E5:E" & lastRow)
        .Format.Fill.ForeColor.RGB = RGB(115, 64, 160)
        .Format.Line.ForeColor.RGB = RGB(33, 99, 191)
    End With
    With chart.SeriesCollection.NewSeries
        .Name = "КП — прогноз / риск"
        .XValues = graph.Range("B5:B" & lastRow)
        .Values = graph.Range("F5:F" & lastRow)
        .Format.Fill.ForeColor.RGB = RGB(235, 63, 63)
        .Format.Line.ForeColor.RGB = RGB(190, 25, 25)
    End With
    With chart.SeriesCollection.NewSeries
        .Name = "СР — план"
        .XValues = graph.Range("B5:B" & lastRow)
        .Values = graph.Range("G5:G" & lastRow)
        .Format.Fill.ForeColor.RGB = RGB(12, 155, 214)
        .Format.Line.ForeColor.RGB = RGB(18, 90, 182)
    End With
    With chart.SeriesCollection.NewSeries
        .Name = "СР — прогноз"
        .XValues = graph.Range("B5:B" & lastRow)
        .Values = graph.Range("H5:H" & lastRow)
        .Format.Fill.ForeColor.RGB = RGB(245, 166, 35)
        .Format.Line.ForeColor.RGB = RGB(192, 115, 0)
    End With
    With chart.SeriesCollection.NewSeries
        .Name = "СР — факт"
        .XValues = graph.Range("B5:B" & lastRow)
        .Values = graph.Range("I5:I" & lastRow)
        .Format.Fill.ForeColor.RGB = RGB(0, 176, 80)
        .Format.Line.ForeColor.RGB = RGB(0, 112, 55)
    End With
    chart.HasTitle = True
    chart.ChartTitle.Text = "Диаграмма Ганта: КП и СР"
    chart.HasLegend = True
    chart.Legend.Position = xlLegendPositionBottom
    chart.Axes(xlCategory).ReversePlotOrder = True
    chart.Axes(xlValue).TickLabels.NumberFormat = "dd.mm.yy"
SafeExit:
    Exit Sub
Fail:
    If silent Then
        failureNumber = Err.Number: failureDescription = Err.Description
        RecordAutomationError "RefreshGantt", failureNumber, failureDescription
        Exit Sub
    End If
    MsgBox "Не удалось обновить диаграмму Ганта. " & Err.Description, vbExclamation, "GIP Navigator"
End Sub
Public Sub GoToSR(): ThisWorkbook.Worksheets("03_СР_проекта").Activate: End Sub
Public Sub GoToGraph(): ThisWorkbook.Worksheets("05_График").Activate: End Sub
Public Sub GoToControl(): ThisWorkbook.Worksheets("06_Контроль").Activate: End Sub

Private Sub RecordAutomationError(ByVal operationName As String, ByVal errorNumber As Long, ByVal errorDescription As String)
    Dim service As Worksheet
    On Error GoTo UnableToRecord
    Set service = ThisWorkbook.Worksheets("99_Служебный")
    service.Range("A20").Value = "last_automation_error"
    service.Range("B20").Value = operationName & " (" & CStr(errorNumber) & "): " & errorDescription
    Exit Sub
UnableToRecord:
    Debug.Print "GIP Navigator " & operationName & ": " & errorDescription
End Sub
