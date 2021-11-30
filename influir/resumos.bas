Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Address <> "$A$1" Then Exit Sub
    am = all_unique("L", "BASE_VENDAS")
    With Sheets("BASE_VENDAS")
        For ano_mes = 0 To UBound(am)
            Sheets("BASE_RESUMO").Cells(3, ano_mes + 2).Value = WorksheetFunction.SumIfs(.Range("D:D"), .Range("L:L"), am(ano_mes), .Range("P:P"), Sheets("BASE_RESUMO").Range("A1"))
        Next
    End With
End Sub

Sub set_resumo()
    Dim situacao As Variant, situacoes As String: situacoes = ""
    For Each situacao In all_unique("P", "BASE_VENDAS")
        situacoes = situacao & "," & situacoes
    Next
    With Sheets("BASE_RESUMO").Range("A1").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=situacoes
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    am = all_unique("L", "BASE_VENDAS")
    With Sheets("BASE_RESUMO")
        For ano_mes = 0 To UBound(am)
            .Cells(2, ano_mes + 2).Value = "'" & am(ano_mes)
            .Cells(1, ano_mes + 2).Value = MonthName(Right(am(ano_mes), 2), True)
        Next
    End With
End Sub

