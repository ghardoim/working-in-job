Sub set_giro()
    Dim linha As Integer: linha = 6
    Call liga_desliga(False)
    Call drop_giro
    
    referencias_produto_cor = all_unique("AM", "BASE_VENDAS")

    For Each produto_cor In referencias_produto_cor

        var_tamanhos = Array("34", "36", "38", "40", "PP", "P", "M", "G", "ÚNICO")
        
        Dim svendas As Worksheet: Set svendas = ThisWorkbook.Sheets("BASE_VENDAS")
        Dim sprodutos As Worksheet: Set sprodutos = ThisWorkbook.Sheets("BASE_PRODUTOS")

        With ThisWorkbook.Sheets("BASE_GIRO")
            On Error Resume Next
            svendas.Range("A5").AutoFilter Field:=39, Criteria1:=produto_cor

            data_lancamento = indice_corresp(produto_cor, "A:A", "B:B", "BASE_APOIO")
            If data_lancamento = "" Then data_lancamento = WorksheetFunction.Subtotal(5, svendas.Range("P:P"))
            .Cells(linha, 1).Value = data_lancamento

            .Cells(linha, 2).Value = produto_cor
            .Cells(linha, 3).Value = indice_corresp(produto_cor, "AM:AM", "A:A", svendas.Name)
            .Cells(linha, 4).Value = indice_corresp(produto_cor, "AM:AM", "D:D", svendas.Name)

            .Cells(linha, 5).Value = indice_corresp(produto_cor, "R:R", "H:H", sprodutos.Name)
            .Cells(linha, 7).Value = indice_corresp(produto_cor, "R:R", "J:J", sprodutos.Name)

            estoque_atual = WorksheetFunction.SumIfs(sprodutos.Range("G:G"), sprodutos.Range("R:R"), produto_cor)
            estoque_inicial = estoque_atual + WorksheetFunction.Subtotal(9, svendas.Range("C:C"))
            .Cells(linha, 8).Value = estoque_atual
            .Cells(linha, 9).Value = estoque_inicial
            
            Dim coluna As Integer: coluna = 10
    
            For Each x_dias In Array(7, 10, 15, 20, 30, 40, 45, 60, 90)
                inicio_x_dias = coluna
                svendas.Range("A5").AutoFilter Field:=16, Criteria1:="<=" & CDbl(DateAdd("d", x_dias, data_lancamento))
                For Each tamanho In var_tamanhos
                    .Cells(5, coluna).Value = tamanho
                    svendas.Range("A5").AutoFilter Field:=5, Criteria1:=tamanho
                    .Cells(linha, coluna).Value = WorksheetFunction.Subtotal(9, svendas.Range("C:C"))
                    coluna = coluna + 1
                Next
                .Cells(5, coluna).Value = "???"
                svendas.Range("A5").AutoFilter Field:=5, Criteria1:="="
                .Cells(linha, coluna).Value = WorksheetFunction.Subtotal(9, svendas.Range("C:C"))
                coluna = coluna + 1

                .Cells(5, coluna).Value = "Vendas " & x_dias & " dias"
                .Cells(linha, coluna).Value = WorksheetFunction.Sum(.Range(.Cells(linha, inicio_x_dias), .Cells(linha, coluna - 1)))
                coluna = coluna + 1
            Next

            coluna_venda_dias = 20
            For Each x_dias In Array(7, 10, 15, 20, 30, 40, 45, 60, 90)
                .Cells(5, coluna).Value = "Giro " & x_dias & " dias"
                .Cells(linha, coluna).Value = .Cells(linha, coluna_venda_dias).Value / estoque_inicial
                coluna_venda_dias = coluna_venda_dias + 11
                coluna = coluna + 1
            Next
            svendas.Range("A5").AutoFilter Field:=5
            svendas.Range("A5").AutoFilter Field:=16
            .Cells(linha, coluna).Value = CInt(Date - CDate(data_lancamento))
            .Cells(linha, coluna + 1).Value = WorksheetFunction.Subtotal(5, svendas.Range("P:P"))
            .Cells(linha, coluna + 2).Value = WorksheetFunction.Subtotal(4, svendas.Range("P:P"))
            .Cells(linha, coluna + 3).Value = WorksheetFunction.Subtotal(9, svendas.Range("C:C"))
            .Cells(linha, coluna + 4).Value = .Cells(linha, coluna + 3).Value / estoque_inicial

            coluna = coluna + 5
            For Each tamanho In var_tamanhos
                .Cells(5, coluna).Value = "Giro " & tamanho
                svendas.Range("A5").AutoFilter Field:=5, Criteria1:=tamanho
                venda_tamanho = WorksheetFunction.Subtotal(9, svendas.Range("C:C"))
                estoque_tamanho = venda_tamanho + WorksheetFunction.SumIfs(sprodutos.Range("G:G"), sprodutos.Range("R:R"), produto_cor, sprodutos.Range("E:E"), tamanho)
                .Cells(linha, coluna).Value = venda_tamanho / estoque_tamanho
                coluna = coluna + 1
            Next
            svendas.Range("A5").AutoFilter Field:=5

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