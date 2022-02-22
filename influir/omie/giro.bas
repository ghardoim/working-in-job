Sub set_giro()
    Dim linha As Integer: linha = 6
    Call liga_desliga(False)

    ThisWorkbook.Sheets("BASE_PRODUTOS").Range("A5").AutoFilter Field:=14, Criteria1:="="
    ThisWorkbook.Sheets("BASE_VENDAS").Range("A5").AutoFilter Field:=11, Criteria1:="Autorizado"
    ThisWorkbook.Sheets("BASE_VENDAS").Range("A5").AutoFilter Field:=10, Operator:=xlFilterValues, _
                        Criteria1:=Array("Venda de Produto pelo PDV", "Pedido de Venda", "Devolução de Venda", "Devolução (Emissão do Cliente)")
    ThisWorkbook.Sheets("BASE_VENDAS").Range("A5").AutoFilter Field:=12, Operator:=xlFilterValues, _
                        Criteria1:=Array("Clientes - Vendas PDV", "Clientes - Vendas Malinha / Whatsapp", _
                        "Clientes - Vendas Farfetch Nacional", "Clientes - Vendas Farfetch Internacional", _
                        "Clientes - Vendas Site Neriage", "Devoluções de Vendas de Mercadoria", _
                        "Devoluções de Compra de Mercadoria de Revenda")

    Set referencias_produto_cor = all_unique("W", "BASE_VENDAS")

    For Each produto_cor In referencias_produto_cor

        var_tamanhos = tamanhos
        Dim svendas As Worksheet: Set svendas = ThisWorkbook.Sheets("BASE_VENDAS")
        Dim sprodutos As Worksheet: Set sprodutos = ThisWorkbook.Sheets("BASE_PRODUTOS")

        With ThisWorkbook.Sheets("BASE_GIRO")
            On Error Resume Next
            svendas.Range("A5").AutoFilter Field:=23, Criteria1:=produto_cor

            data_lancamento = indice_corresp(produto_cor, "A:A", "B:B", "BASE_APOIO")
            If data_lancamento = "" Then data_lancamento = WorksheetFunction.Subtotal(5, svendas.Range("G:G"))
            .Cells(linha, 1).Value = data_lancamento

            .Cells(linha, 2).Value = produto_cor
            .Cells(linha, 3).Value = indice_corresp(produto_cor, "W:W", "I:I", svendas.Name)
            .Cells(linha, 4).Value = indice_corresp(produto_cor, "W:W", "V:V", svendas.Name)

            .Cells(linha, 5).Value = indice_corresp(produto_cor, "Q:Q", "C:C", sprodutos.Name)
            .Cells(linha, 6).Value = indice_corresp(produto_cor, "Q:Q", "I:I", sprodutos.Name)
            .Cells(linha, 8).Value = indice_corresp(produto_cor, "Q:Q", "K:K", sprodutos.Name)

            estoque_atual = WorksheetFunction.SumIfs(sprodutos.Range("J:J"), sprodutos.Range("Q:Q"), produto_cor)
            estoque_inicial = estoque_atual + WorksheetFunction.Subtotal(9, svendas.Range("T:T"))
            .Cells(linha, 9).Value = estoque_atual
            .Cells(linha, 10).Value = estoque_inicial
            
            Dim coluna As Integer: coluna = 11
    
            For Each x_dias In Array(7, 10, 15, 20, 30, 40, 45, 60, 90)
                inicio_x_dias = coluna
                svendas.Range("A5").AutoFilter Field:=7, Criteria1:="<=" & CDbl(DateAdd("d", x_dias, data_lancamento))
                For Each tamanho In var_tamanhos
                    .Cells(5, coluna).Value = tamanho
                    svendas.Range("A5").AutoFilter Field:=21, Criteria1:=tamanho
                    .Cells(linha, coluna).Value = WorksheetFunction.Subtotal(9, svendas.Range("T:T"))
                    coluna = coluna + 1
                Next
                .Cells(5, coluna).Value = "???"
                svendas.Range("A5").AutoFilter Field:=21, Criteria1:="="
                .Cells(linha, coluna).Value = WorksheetFunction.Subtotal(9, svendas.Range("T:T"))
                coluna = coluna + 1

                .Cells(5, coluna).Value = "Vendas " & x_dias & " dias"
                .Cells(linha, coluna).Value = WorksheetFunction.Sum(.Range(.Cells(linha, inicio_x_dias), .Cells(linha, coluna - 1)))
                coluna = coluna + 1
            Next

            coluna_venda_dias = 29
            For Each x_dias In Array(7, 10, 15, 20, 30, 40, 45, 60, 90)
                .Cells(5, coluna).Value = "Giro " & x_dias & " dias"
                .Cells(linha, coluna).Value = .Cells(linha, coluna_venda_dias).Value / estoque_inicial
                coluna_venda_dias = coluna_venda_dias + 19
                coluna = coluna + 1
            Next
            svendas.Range("A5").AutoFilter Field:=21
            svendas.Range("A5").AutoFilter Field:=7
            .Cells(linha, coluna).Value = CInt(Date - CDate(data_lancamento))
            .Cells(linha, coluna + 1).Value = WorksheetFunction.Subtotal(5, svendas.Range("G:G"))
            .Cells(linha, coluna + 2).Value = WorksheetFunction.Subtotal(4, svendas.Range("G:G"))
            .Cells(linha, coluna + 3).Value = WorksheetFunction.Subtotal(9, svendas.Range("T:T"))

            On Error GoTo 0        
        End With
        linha = linha + 1
    Next
    ThisWorkbook.Sheets("BASE_PRODUTOS").Range("A5").AutoFilter
    ThisWorkbook.Sheets("BASE_VENDAS").Range("A5").AutoFilter
    Call liga_desliga(True)

    Call MsgBox("agora o giro das vendas foi atualizado! :D", vbInformation, "Base Atualizada")
End Sub

Sub drop_giro()
    Call liga_desliga(False)
    Sheets("BASE_GIRO").Rows("6:1048576").Delete
    Call liga_desliga(True)
End Sub