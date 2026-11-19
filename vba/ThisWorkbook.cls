Private Sub Workbook_Open()
    On Error GoTo Fail
    ThisWorkbook.ForceFullCalculation = True
    Application.Calculation = xlCalculationAutomatic
    Application.CalculateFullRebuild
    ThisWorkbook.Worksheets("01_КП").Protect UserInterfaceOnly:=True, AllowFiltering:=True
    ' Linked checkbox values persist in the workbook.  Synchronize them on
    ' opening so a new project immediately receives its selected standard SRs.
    SyncSelectedSR True
    Exit Sub
Fail:
    MsgBox "Не удалось подготовить книгу при открытии. " & Err.Description, vbExclamation, "GIP Navigator"
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Dim changedStart As Range, cell As Range
    On Error GoTo Fail
    If Not Application.EnableEvents Then Exit Sub
    If Sh.Name = "01_КП" Then
        ' B3 is the single project-level date of signing the order.  A change
        ' immediately propagates through the shifted KP dates, SR network and Gantt.
        If Not Intersect(Target, Union(Sh.Range("B3"), Sh.Range("A5:I79"))) Is Nothing Then
            RecalculateSRNetwork
            RefreshGantt
        End If
    ElseIf Sh.Name = "03_СР_проекта" Then
        Set changedStart = Intersect(Target, Sh.Range("D5:D304"))
        If Not changedStart Is Nothing Then
            ' D is the displayed calculation. A direct date entry is retained
            ' separately as a "not earlier than" constraint before recalculation.
            For Each cell In changedStart.Cells
                Sh.Cells(cell.Row, 15).Value = cell.Value
            Next cell
        End If
        If Not Intersect(Target, Sh.Range("A5:O304")) Is Nothing Then RecalculateSRNetwork
    ElseIf Sh.Name = "04_Согласование_плана" Then
        If Not Intersect(Target, Sh.Range("A5:O304")) Is Nothing Then RecalculateSRNetwork
    ElseIf Sh.Name = "07_Актирование" Then
        If Not Intersect(Target, Sh.Range("A5:H304")) Is Nothing Then
            RecalculateActs
            RecalculateKpObligations
            RefreshControl ThisWorkbook.Worksheets("03_СР_проекта")
            RefreshGantt
        End If
    End If
    Exit Sub
Fail:
    MsgBox "Не удалось пересчитать связанные работы. " & Err.Description, vbExclamation, "GIP Navigator"
End Sub
