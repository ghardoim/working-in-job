Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Address <> "$A$5" Then Exit Sub
    am = all_unique("L", "BASE_VENDAS")
    With Sheets("BASE_RESUMO")
        For ano_mes = 0 To UBound(am)
            'incluir condição do canal de venda
            For canal = 9 To 14
                .Cells(canal, ano_mes + 2).Value = WorksheetFunction.SumIfs(Sheets("BASE_VENDAS").Range("D:D"), Sheets("BASE_VENDAS").Range("L:L"), am(ano_mes), Sheets("BASE_VENDAS").Range("P:P"), Sheets("BASE_RESUMO").Range("A5"))
                .Cells(canal, UBound(am) + 3).Value = WorksheetFunction.Sum(.Range(.Cells(canal, 2), .Cells(canal, UBound(am) + 2)))
            Next
            .Cells(8, ano_mes + 2).Value = WorksheetFunction.Sum(.Range(.Cells(9, ano_mes + 2), .Cells(14, ano_mes + 2)))
            .Cells(8, ano_mes + 3).Value = WorksheetFunction.Sum(.Range(.Cells(8, 2), .Cells(8, UBound(am) + 2)))
        Next
    End With
End Sub

Sub set_resumo()
    Dim situacao As Variant, situacoes As String: situacoes = ""
    For Each situacao In all_unique("P", "BASE_VENDAS")
        situacoes = situacao & "," & situacoes
    Next
    With Sheets("BASE_RESUMO").Range("A5").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=situacoes
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    am = all_unique("L", "BASE_VENDAS")
    With Sheets("BASE_RESUMO")
        For ano_mes = 0 To UBound(am)
            With .Cells(6, ano_mes + 2)
                .Value = "'" & am(ano_mes)
                .Interior.Color = RGB(180, 250, 120)
                .Borders.LineStyle = xlContinuous
                .Font.Bold = True
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
            End With
            With .Cells(5, ano_mes + 2)
                .Value = MonthName(Right(am(ano_mes), 2), True)
                .Interior.Color = RGB(173, 216, 230)
                .Borders.LineStyle = xlContinuous
                .Font.Bold = True
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
            End With
        Next
        With .Cells(5, ano_mes + 2)
            .Value = "total"
            .Interior.Color = RGB(173, 216, 230)
            .Borders.LineStyle = xlContinuous
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
        End With
    End With
End Sub