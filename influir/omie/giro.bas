Sub set_giro()
    Dim linha As Integer: linha = 6

    ThisWorkbook.Sheets("BASE_PRODUTOS").Range("A5").AutoFilter Field:=14, Criteria1:="="
    ThisWorkbook.Sheets("BASE_VENDAS").Range("A5").AutoFilter Field:=11, Criteria1:="Autorizado"
    ThisWorkbook.Sheets("BASE_VENDAS").Range("A5").AutoFilter Field:=12, Operator:=xlFilterValues, _
                        Criteria1:=Array("Clientes - Vendas PDV", "Clientes - Vendas Malinha / Whatsapp", _
                        "Clientes - Vendas Farfetch Nacional", "Clientes - Vendas Farfetch Internacional", _
                        "Clientes - Vendas Site Neriage", "Devoluções de Vendas de Mercadoria", _
                        "Devoluções de Compra de Mercadoria de Revenda")
    
    referencias_produto_cor = all_unique("V", "BASE_VENDAS")

    For Each produto_cor In referencias_produto_cor
        Call atualiza_giro(linha, produto_cor)
        linha = linha + 1
    Next
    ThisWorkbook.Sheets("BASE_PRODUTOS").Range("A5").AutoFilter
    ThisWorkbook.Sheets("BASE_VENDAS").Range("A5").AutoFilter
    Call MsgBox("agora o giro das vendas foi atualizado! :D", vbInformation, "Base Atualizada")
End Sub

Sub atualiza_giro(linha As Integer, ByVal produto_cor As String)
    Call liga_desliga(False)

    var_tamanhos = tamanhos
    Dim svendas As Worksheet: Set svendas = ThisWorkbook.Sheets("BASE_VENDAS")
    Dim sprodutos As Worksheet: Set sprodutos = ThisWorkbook.Sheets("BASE_PRODUTOS")

    With ThisWorkbook.Sheets("BASE_GIRO")
        On Error GoTo nao_achei

        data_lancamento = indice_corresp(produto_cor, "A:A", "B:B", "BASE_APOIO")
        If data_lancamento = "" Then data_lancamento = WorksheetFunction.MinIfs(svendas.Range("G:G"), svendas.Range("V:V"), produto_cor)
        .Cells(linha, 1).Value = data_lancamento

        .Cells(linha, 2).Value = produto_cor
        .Cells(linha, 3).Value = indice_corresp(produto_cor, "V:V", "I:I", svendas.Name)
        .Cells(linha, 4).Value = indice_corresp(produto_cor, "V:V", "U:U", svendas.Name)

        .Cells(linha, 5).Value = indice_corresp(produto_cor, "Q:Q", "C:C", sprodutos.Name)
        .Cells(linha, 6).Value = indice_corresp(produto_cor, "Q:Q", "I:I", sprodutos.Name)
        .Cells(linha, 8).Value = indice_corresp(produto_cor, "Q:Q", "K:K", sprodutos.Name)

        estoque_atual = WorksheetFunction.SumIfs(sprodutos.Range("J:J"), sprodutos.Range("Q:Q"), produto_cor)
        estoque_inicial = estoque_atual + WorksheetFunction.CountIfs(svendas.Range("V:V"), produto_cor, svendas.Range("L:L"), "Devoluções*")
        .Cells(linha, 9).Value = estoque_atual
        .Cells(linha, 10).Value = estoque_inicial

        Dim coluna As Integer: coluna = 11
        For Each x_dias In Array(7, 10, 15, 20, 30, 40, 45, 60)
            inicio_x_dias = coluna
            For Each tamanho In var_tamanhos
                .Cells(5, coluna).Value = tamanho
                .Cells(linha, coluna).Value = WorksheetFunction.CountIfs(svendas.Range("V:V"), produto_cor, svendas.Range("T:T"), tamanho, svendas.Range("G:G"), "<=" & DateAdd("d", x_dias, data_lancamento))
                coluna = coluna + 1
            Next
            .Cells(5, coluna).Value = "???"
            .Cells(linha, coluna).Value = WorksheetFunction.CountIfs(svendas.Range("V:V"), produto_cor, svendas.Range("T:T"), "", svendas.Range("G:G"), "<=" & DateAdd("d", x_dias, data_lancamento))
            coluna = coluna + 1

            .Cells(5, coluna).Value = "Vendas " & x_dias & " dias"
            .Cells(linha, coluna).Value = WorksheetFunction.Sum(.Range(.Cells(linha, inicio_x_dias), .Cells(linha, coluna - 1)))
            coluna = coluna + 1
        Next

        coluna_venda_dias = 29
        For Each x_dias In Array(7, 10, 15, 20, 30, 40, 45, 60)
            .Cells(5, coluna).Value = "Giro " & x_dias & " dias"
            .Cells(linha, coluna).Value = .Cells(linha, coluna_venda_dias).Value / estoque_inicial

            coluna_venda_dias = coluna_venda_dias + 19
            coluna = coluna + 1
        Next

        .Cells(linha, coluna + 1).Value = WorksheetFunction.MinIfs(svendas.Range("G:G"), svendas.Range("V:V"), produto_cor)
        .Cells(linha, coluna + 2).Value = WorksheetFunction.MaxIfs(svendas.Range("G:G"), svendas.Range("V:V"), produto_cor)
        .Cells(linha, coluna).Value = CInt(CDate(data_lancamento) - Date)
nao_achei:
    If Err.Number = 1004 Then: .Cells(linha, 2).Interior.ColorIndex = 3
    On Error GoTo 0

    End With
    Call liga_desliga(True)
End Sub

Sub drop_giro()
    Call liga_desliga(False)
    Sheets("BASE_GIRO").Rows("6:1048576").Delete
    Call liga_desliga(True)
End Sub