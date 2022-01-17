Sub set_giro()
    Call liga_desliga(False)
    var_tamanhos = tamanhos
    Dim svendas As Worksheet: Set svendas = ThisWorkbook.Sheets("BASE_VENDAS")
    Dim sprodutos As Worksheet: Set sprodutos = ThisWorkbook.Sheets("BASE_PRODUTOS")
    Dim linha As Integer: linha = 6

    referencias_produto_cor = all_unique("V", svendas.Name)
    With Sheets("BASE_GIRO")
        For Each produto_cor In referencias_produto_cor
            .Cells(linha, 2) = produto_cor
            .Cells(linha, 3).Value = indice_corresp(produto_cor, "V:V", "I:I", svendas.Name)
            .Cells(linha, 4).Value = indice_corresp(produto_cor, "V:V", "U:U", svendas.Name)

            On Error GoTo nao_achei
            .Cells(linha, 5).Value = indice_corresp(produto_cor, "Q:Q", "I:I", sprodutos.Name)
            .Cells(linha, 7).Value = indice_corresp(produto_cor, "Q:Q", "K:K", sprodutos.Name)

            Dim coluna As Integer: coluna = 8
            For Each tamanho In var_tamanhos
                .Cells(5, coluna).Value = tamanho
                .Cells(linha, coluna).Value = WorksheetFunction.SumIfs(sprodutos.Range("J:J"), sprodutos.Range("Q:Q"), produto_cor, sprodutos.Range("P:P"), tamanho)
                coluna = coluna + 1
            Next
            .Cells(5, coluna).Value = "???"
            .Cells(linha, coluna).Value = WorksheetFunction.CountIfs(svendas.Range("V:V"), produto_cor, svendas.Range("T:T"), "")
            coluna = coluna + 1

            .Cells(linha, coluna).Value = WorksheetFunction.Sum(.Range(.Cells(linha, 8), .Cells(linha, coluna - 1)))
            coluna = coluna + 2

            For Each x_dias In Array(10, 15, 30, 60)
                inicio_x_dias = coluna
                For Each tamanho In var_tamanhos
                    .Cells(5, coluna).Value = tamanho
                    .Cells(linha, coluna).Value = WorksheetFunction.CountIfs(svendas.Range("V:V"), produto_cor, svendas.Range("T:T"), tamanho, svendas.Range("G:G"), "<=" & DateAdd("d", x_dias, .Cells(3, 5).Value))
                    coluna = coluna + 1
                Next
                .Cells(5, coluna).Value = "???"
                .Cells(linha, coluna).Value = WorksheetFunction.CountIfs(svendas.Range("V:V"), produto_cor, svendas.Range("T:T"), "", svendas.Range("G:G"), "<=" & DateAdd("d", x_dias, .Cells(3, 5).Value))
                coluna = coluna + 1

                .Cells(5, coluna).Value = "Vendas " & x_dias & " dias"
                .Cells(linha, coluna).Value = WorksheetFunction.Sum(.Range(.Cells(linha, inicio_x_dias), .Cells(linha, coluna - 1)))
                coluna = coluna + 1
            Next
            .Cells(linha, coluna).Value = CInt(CDate(.Cells(3, 5).Value) - CDate(.Cells(linha, 1).Value))
            .Cells(linha, coluna + 1).Value = WorksheetFunction.MinIfs(Sheets("BASE_VENDAS").Range("G:G"), Sheets("BASE_VENDAS").Range("V:V"), produto_cor)
            .Cells(linha, coluna + 2).Value = WorksheetFunction.MaxIfs(Sheets("BASE_VENDAS").Range("G:G"), Sheets("BASE_VENDAS").Range("V:V"), produto_cor)

nao_achei:
    If Err.Number = 1004 Then: .Cells(linha, 2).Interior.ColorIndex = 3
    On Error GoTo 0

            linha = linha + 1
        Next
    End With
    Call MsgBox("agora o giro das vendas foi atualizado! :D", vbInformation, "Base Atualizada")
    Call liga_desliga(True)
End Sub

Sub drop_giro()
    Sheets("BASE_GIRO").Rows("6:1048576").Delete
End Sub