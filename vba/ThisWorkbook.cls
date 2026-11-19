Private Sub Workbook_Open()
    On Error Resume Next
    ThisWorkbook.ForceFullCalculation = True
    Application.Calculation = xlCalculationAutomatic
    Application.CalculateFullRebuild
    RefreshGantt
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo SafeExit
    If Not Application.EnableEvents Then Exit Sub
    If Sh.Name = "01_КП" Then
        If Not Intersect(Target, Sh.Range("A5:I79")) Is Nothing Then
            RecalculateSRNetwork True
            RefreshGantt
        End If
    ElseIf Sh.Name = "03_СР_проекта" Then
        If Not Intersect(Target, Sh.Range("A5:M304")) Is Nothing Then RecalculateSRNetwork True
    ElseIf Sh.Name = "04_Согласование_плана" Then
        If Not Intersect(Target, Sh.Range("A5:G304")) Is Nothing Then RecalculateSRNetwork True
    End If
SafeExit:
End Sub
