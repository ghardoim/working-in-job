Sub set_giro()
    Call liga_desliga(False)
    Dim linha As Integer: linha = 6
    referencias_produto_cor = all_unique("V", "BASE_VENDAS")
    With Sheets("BASE_GIRO")
        For Each produto_cor In referencias_produto_cor
            .Cells(linha, 2) = produto_cor
            .Cells(linha, 3).Value = indice_corresp(produto_cor, "V:V", "I:I", "BASE_VENDAS")
            .Cells(linha, 4).Value = indice_corresp(produto_cor, "V:V", "U:U", "BASE_VENDAS")

            On Error GoTo nao_achei
            .Cells(linha, 5).Value = indice_corresp(produto_cor, "Q:Q", "I:I", "BASE_PRODUTOS")
            .Cells(linha, 7).Value = indice_corresp(produto_cor, "Q:Q", "K:K", "BASE_PRODUTOS")

            Dim coluna As Integer: coluna = 8
            Dim total_estoque As Integer: total_estoque = 0
            For Each tamanho In tamanhos
                .Cells(5, coluna).Value = tamanho
                estoque = WorksheetFunction.SumIfs(Sheets("BASE_PRODUTOS").Range("J:J"), Sheets("BASE_PRODUTOS").Range("Q:Q"), produto_cor, Sheets("BASE_PRODUTOS").Range("P:P"), tamanho)
                .Cells(linha, coluna).Value = estoque
                coluna = coluna + 1
                total_estoque = total_estoque + estoque
            Next
            .Cells(linha, coluna).Value = total_estoque
            .Cells(linha, coluna + 1).Value = WorksheetFunction.MinIfs(Sheets("BASE_VENDAS").Range("G:G"), Sheets("BASE_VENDAS").Range("V:V"), produto_cor)
            .Cells(linha, coluna + 2).Value = WorksheetFunction.MaxIfs(Sheets("BASE_VENDAS").Range("G:G"), Sheets("BASE_VENDAS").Range("V:V"), produto_cor)

nao_achei:
    If Err Then: .Cells(linha, 2).Interior.ColorIndex = 3
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