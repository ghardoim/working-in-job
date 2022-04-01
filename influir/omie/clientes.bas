Sub set_clientes()
    Call liga_desliga(False)

    Dim bvendas As Worksheet: Set bvendas = ThisWorkbook.Sheets("BASE_VENDAS")
    Dim linha As Integer: linha = bvendas.Range("A1048576").End(xlUp).Row
    bvendas.Range("A5").AutoFilter Field:=11, Criteria1:="Autorizado"
    bvendas.Range("A5").AutoFilter Field:=10, Operator:=xlFilterValues, _
        Criteria1:=Array("Venda de Produto pelo PDV", "Pedido de Venda", "Devolução de Venda", "Devolução (Emissão do Cliente)")
    bvendas.Range("A5").AutoFilter Field:=12, Operator:=xlFilterValues, _
                        Criteria1:=Array("Clientes - Vendas PDV", "Clientes - Vendas Malinha / Whatsapp", _
                        "Clientes - Vendas Farfetch Nacional", "Clientes - Vendas Farfetch Internacional", _
                        "Clientes - Vendas Site Neriage", "Devoluções de Vendas de Mercadoria", _
                        "Devoluções de Compra de Mercadoria de Revenda")

    Set anos_meses = all_unique("X", bvendas.Name)
    With ThisWorkbook.Sheets("BASE_CLIENTES")
        .Rows("6:1048576").Delete
        .Range("A6:F" & linha).Value = bvendas.Range("A6:F" & linha).Value
        .Range("A6:F" & linha).RemoveDuplicates Columns:=2

        For linha = 6 To .Range("A1048576").End(xlUp).Row
            Cliente = .Cells(linha, 2).Value
            bvendas.Range("A5").AutoFilter Field:=2, Criteria1:=Cliente

            primeira_venda = WorksheetFunction.Subtotal(5, bvendas.Range("G:G"))
            .Cells(linha, 7).Value = IIf(primeira_venda <> 0, primeira_venda, "")

            ultima_venda = WorksheetFunction.Subtotal(4, bvendas.Range("G:G"))
            .Cells(linha, 8).Value = IIf(ultima_venda <> 0, ultima_venda, "")
            .Cells(linha, 9).Value = IIf(ultima_venda <> 0, Date - ultima_venda, "")

            .Cells(linha, 10).Value = WorksheetFunction.Subtotal(9, bvendas.Range("R:R"))

            coluna = 11
            For Each am In anos_meses
                .Cells(5, coluna).Value = am
                bvendas.Range("A5").AutoFilter Field:=24, Criteria1:=am

                .Cells(linha, coluna).Value = WorksheetFunction.Subtotal(9, bvendas.Range("R:R"))
                coluna = coluna + 1
            Next
            Set range_anomes = .Range(.Cells(linha, 11), .Cells(linha, anos_meses.Count + coluna - 1))
            n_meses_venda = WorksheetFunction.CountIf(range_anomes, ">0")

            ticket_medio = 0
            If n_meses_venda <> 0 Then ticket_medio = WorksheetFunction.Sum(range_anomes) / n_meses_venda

            .Cells(5, coluna).Value = "Ticket Médio"
            .Cells(linha, coluna).Value = ticket_medio

            .Cells(5, coluna + 1).Value = "Classificação"
            If ticket_medio > .Range("B3").Value Then .Cells(linha, coluna + 1).Value = "Potencial"
            If ticket_medio > .Range("C3").Value Then .Cells(linha, coluna + 1).Value = "VIP"
            
            .Cells(5, coluna + 2).Value = "Recorrente"
            If n_meses_venda >= 6 Then .Cells(linha, coluna + 2).Value = "X"

            .Cells(5, coluna + 3).Value = "Cliente Novo"
            If Year(Date) = Year(primeira_venda) Then .Cells(linha, coluna + 3).Value = "X"

            coluna = coluna + 4
            For Each am In anos_meses
                .Cells(5, coluna).Value = am
                bvendas.Range("A5").AutoFilter Field:=24, Criteria1:=am

                .Cells(linha, coluna).Value = IIf(.Cells(linha, coluna - anos_meses.Count - 4).Value > 0, 1, 0)
                coluna = coluna + 1
            Next
            .Cells(5, coluna).Value = "Ticket Médio"
            .Cells(linha, coluna).Value = IIf(.Cells(linha, coluna - anos_meses.Count - 4).Value > 0, 1, 0)
        Next
    End With
    bvendas.Range("A5").AutoFilter
    Call liga_desliga(True)
End Sub