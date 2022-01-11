Sub set_privenda()
    Call liga_desliga(False)
    With Sheets("BASE_PRIVENDA")
        .Rows("6:1048576").Delete
        Dim ult_linha As Integer: ult_linha = .Range("A1048576").End(xlUp).Row + 1
        produtos_cor = all_unique("AO", "BASE_VENDAS")
        For Each produto In produtos_cor
            .Cells(ult_linha, 1).Value = produto
            primeira_venda = WorksheetFunction.MinIfs(Sheets("BASE_VENDAS").Range("P:P"), Sheets("BASE_VENDAS").Range("AO:AO"), produto)
            .Cells(ult_linha, 2).Value = primeira_venda
            .Cells(ult_linha, 3).Value = primeira_venda + 15
            .Cells(ult_linha, 4).Value = primeira_venda + 40
            .Cells(ult_linha, 5).Value = DateAdd("m", 2, primeira_venda)
            .Cells(ult_linha, 6).Value = DateAdd("m", 3, primeira_venda)
            Select Case Date
                Case Is <= .Cells(ult_linha, 3).Value
                    .Cells(ult_linha, 7).Value = "Lançamento"
                Case Is <= .Cells(ult_linha, 4).Value
                    .Cells(ult_linha, 7).Value = "Novidade"
                Case Is <= .Cells(ult_linha, 5).Value
                    .Cells(ult_linha, 7).Value = "Regular"
                Case Is <= .Cells(ult_linha, 6).Value
                    .Cells(ult_linha, 7).Value = "Sale"
                Case Else
                    .Cells(ult_linha, 7).Value = "Antigo"
            End Select
            ult_linha = ult_linha + 1
        Next
    End With
    Call liga_desliga(True)
End Sub