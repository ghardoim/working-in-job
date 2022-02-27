Sub set_resumo()
    Call liga_desliga(False)

    ThisWorkbook.Sheets("BASE_VENDAS").Range("A5").AutoFilter Field:=11, Criteria1:="Autorizado"
    ThisWorkbook.Sheets("BASE_VENDAS").Range("A5").AutoFilter Field:=10, Operator:=xlFilterValues, _
        Criteria1:=Array("Venda de Produto pelo PDV", "Pedido de Venda", "Devolução de Venda", "Devolução (Emissão do Cliente)")
    ThisWorkbook.Sheets("BASE_VENDAS").Range("A5").AutoFilter Field:=12, Operator:=xlFilterValues, _
                        Criteria1:=Array("Clientes - Vendas PDV", "Clientes - Vendas Malinha / Whatsapp", _
                        "Clientes - Vendas Farfetch Nacional", "Clientes - Vendas Farfetch Internacional", _
                        "Clientes - Vendas Site Neriage", "Devoluções de Vendas de Mercadoria", _
                        "Devoluções de Compra de Mercadoria de Revenda")

    Dim linha_inicio As Integer: linha_inicio = 9
    Set situacoes = all_unique("K", "BASE_VENDAS")
    Set canais = all_unique("L", "BASE_VENDAS")
    Set am = all_unique("X", "BASE_VENDAS")

    With Sheets("BASE_RESUMO")

        format_cell(.Cells(5, am.Count + 3), cor:=RGB(170, 210, 230)).Value = "total"
        For ano_mes = 1 To am.Count

            format_cell(.Cells(6, ano_mes + 2), cor:=RGB(180, 250, 120)).Value = "'" & am(ano_mes)
            format_cell(.Cells(5, ano_mes + 2), cor:=RGB(170, 210, 230)).Value = MonthName(Right(am(ano_mes), 2), True)

            Call total_ano_mes(linha_inicio, "R:R", am.Count, ano_mes, am(ano_mes), "Total venda Captada Liquida", "*")
            Call total_ano_mes(linha_inicio + 1, "Y:Y", am.Count, ano_mes, am(ano_mes), "Atendido Bruto", "Autorizado")
            Call total_ano_mes(linha_inicio + 2, "P:P", am.Count, ano_mes, am(ano_mes), "Atendido Líquido", "Autorizado")
            format_cell(.Cells(linha_inicio + 3, 1), cor:=RGB(170, 210, 230)).Value = "Atendido % Desconto"
            format_cell(.Cells(linha_inicio + 4, 1), cor:=RGB(170, 210, 230)).Value = "Participação % Atendido"

            On Error Resume Next
            format_cell(.Cells(linha_inicio + 3, ano_mes + 2), "Percent").Value = (.Cells(linha_inicio + 1, ano_mes + 2) - .Cells(linha_inicio + 2, ano_mes + 2)) / .Cells(linha_inicio + 1, ano_mes + 2)
            format_cell(.Cells(linha_inicio + 4, ano_mes + 2), "Percent").Value = .Cells(linha_inicio + 2, ano_mes + 2) / .Cells(linha_inicio, ano_mes + 2)

            format_cell(.Cells(linha_inicio + 3, am.Count + 3), "Percent").Value = (.Cells(linha_inicio + 1, am.Count + 3) - .Cells(linha_inicio + 2, am.Count + 3)) / .Cells(linha_inicio + 1, am.Count + 3)
            format_cell(.Cells(linha_inicio + 4, am.Count + 3), "Percent").Value = .Cells(linha_inicio + 2, am.Count + 3) / .Cells(linha_inicio, am.Count + 3)

            Dim linha_inicio_bloco As Integer: linha_inicio_bloco = 20
            Call soma_por_situacao("Autorizado", linha_inicio_bloco, linha_inicio, canais, am.Count, ano_mes, am(ano_mes), "P:P")
            Call soma_por_situacao("Cancelado", linha_inicio_bloco, linha_inicio, canais, am.Count, ano_mes, am(ano_mes), "P:P")
            Call soma_por_situacao("Devolvido*", linha_inicio_bloco, linha_inicio, canais, am.Count, ano_mes, am(ano_mes), "P:P")
            Call soma_por_situacao("Inutilizado", linha_inicio_bloco, linha_inicio, canais, am.Count, ano_mes, am(ano_mes), "P:P")
            Call soma_por_situacao("Rejeitado", linha_inicio_bloco, linha_inicio, canais, am.Count, ano_mes, am(ano_mes), "P:P")

            Call soma_por_situacao("Autorizado", linha_inicio_bloco, linha_inicio, canais, am.Count, ano_mes, am(ano_mes), "P:P", " VENDA BRUTA", False)
            Call soma_por_situacao("Autorizado", linha_inicio_bloco, linha_inicio, canais, am.Count, ano_mes, am(ano_mes), "Q:Q", " DESCONTO REALIZADO", False)

            format_cell(.Cells(linha_inicio_bloco - 1, 1), cor:=RGB(170, 210, 230)).Value = "Atendido % DESCONTO REALIZADO"
            format_cell(.Cells(linha_inicio_bloco, 1), cor:=RGB(210, 210, 230)).Value = "TOTAL % DESCONTO REALIZADO"
            format_cell(.Cells(linha_inicio_bloco, ano_mes + 2), "Percent").Value = get_soma_se("Q:Q", am(ano_mes), "Autorizado") / get_soma_se("P:P", am(ano_mes), "Autorizado")
            format_cell(.Cells(linha_inicio_bloco, am.Count + 3), "Percent").Value = get_soma_se("Q:Q", am(ano_mes), "Autorizado") / get_soma_se("P:P", am(ano_mes), "Autorizado")
            For Each canal In canais
                linha_inicio_bloco = linha_inicio_bloco + 1
                format_cell(.Cells(linha_inicio_bloco, 1), cor:=RGB(210, 210, 230)).Value = canal
                format_cell(.Cells(linha_inicio_bloco, ano_mes + 2), "Percent").Value = get_soma_se("Q:Q", am(ano_mes), "Autorizado", canal) / get_soma_se("P:P", am(ano_mes), "Autorizado", canal)
                format_cell(.Cells(linha_inicio_bloco, am.Count + 3), "Percent").Value = get_soma_se("Q:Q", situacao:="Autorizado", canal:=canal) / get_soma_se("P:P", situacao:="Autorizado", canal:=canal)
            Next
            On Error GoTo 0
        Next
    End With
    Call MsgBox("agora todos os resumos estão aqui! :D", vbInformation, "Resumo Atualizado")
    Sheets("BASE_VENDAS").Range("A5").AutoFilter
    Call liga_desliga(True)
End Sub

Function get_soma_se(coluna_soma As String, Optional ByVal ano_mes As String = "*", Optional ByVal situacao As String = "*", Optional ByVal canal As String = "*") As Double
    With Sheets("BASE_VENDAS")
        .Range("A5").AutoFilter Field:=24, Criteria1:=ano_mes
        .Range("A5").AutoFilter Field:=11, Criteria1:=situacao
        .Range("A5").AutoFilter Field:=12, Criteria1:=canal
        get_soma_se = WorksheetFunction.Subtotal(9, .Range(coluna_soma))
    End With
End Function

Sub soma_por_situacao(situacao As String, ByRef linha_inicio_bloco As Integer, linha_inicio As Integer, canais As Variant, am_count As Integer, ByVal ano_mes As Integer, ByVal mes_ano As String, Optional coluna_soma As String = "R:R", Optional complemento As String = "", Optional need_percent As Boolean = True)
    Dim linha_percent As Integer: linha_percent = linha_inicio_bloco - 1

    With Sheets("BASE_RESUMO")
        format_cell(.Cells(linha_inicio_bloco, 1), cor:=RGB(210, 210, 230)).Value = "TOTAL" & complemento
        format_cell(.Cells(linha_inicio_bloco - 1, 1), cor:=RGB(170, 210, 230)).Value = situacao & complemento

        Call soma_por_canal(linha_inicio_bloco, canais, am_count, ano_mes, mes_ano, situacao, coluna_soma)

        format_cell(.Cells(linha_inicio_bloco - (canais.Count + 1), ano_mes + 2), "Currency").Value = WorksheetFunction.Sum(.Range(.Cells(linha_inicio_bloco + 1 - (canais.Count + 1), ano_mes + 2), .Cells(linha_inicio_bloco + (canais.Count + 1), ano_mes + 2)))
        format_cell(.Cells(linha_inicio_bloco - (canais.Count + 1), am_count + 3), "Currency").Value = WorksheetFunction.Sum(.Range(.Cells(linha_inicio_bloco + 1 - (canais.Count + 1), ano_mes + 3), .Cells(linha_inicio_bloco + (canais.Count + 1), ano_mes + 3)))
    End With

    If need_percent Then Call percentual(linha_percent, linha_inicio, am_count, ano_mes)
    linha_inicio_bloco = linha_inicio_bloco + 10 - (canais.Count + 1)
End Sub

Sub soma_por_canal(ByRef linha_inicio_bloco As Integer, canais As Variant, am_count As Integer, ByVal ano_mes As Integer, ByVal mes_ano As String, ByVal situacao As String, coluna_soma As String)
    With Sheets("BASE_RESUMO")
        For Each canal_venda In canais
            Sheets("BASE_VENDAS").Range("A5").AutoFilter Field:=12, Criteria1:=canal_venda
            format_cell(.Cells(linha_inicio_bloco + 1, 1), cor:=RGB(210, 210, 230)).Value = canal_venda
            format_cell(.Cells(linha_inicio_bloco + 1, ano_mes + 2), "Currency").Value = get_soma_se(coluna_soma, mes_ano, situacao, canal_venda)
            format_cell(.Cells(linha_inicio_bloco + 1, am_count + 3), "Currency").Value = WorksheetFunction.Sum(.Range(.Cells(linha_inicio_bloco + 1, 2), .Cells(linha_inicio_bloco + 1, am_count + 2)))
            linha_inicio_bloco = linha_inicio_bloco + 1
        Next
    End With
End Sub

Sub percentual(linha_percent As Integer, linha_inicio As Integer, am_count As Variant, ByVal ano_mes As Integer)
    With Sheets("BASE_RESUMO")
        On Error Resume Next
            format_cell(.Cells(linha_percent, ano_mes + 2), "Percent") = .Cells(linha_percent + 1, ano_mes + 2) / .Cells(linha_inicio, ano_mes + 2)
            format_cell(.Cells(linha_percent, am_count + 3), "Percent") = .Cells(linha_percent + 1, am_count + 3) / .Cells(linha_inicio, am_count + 3)
        On Error GoTo 0
    End With
End Sub

Sub total_ano_mes(linha_inicio As Integer, coluna_soma As String, am_count As Integer, ByVal ano_mes As Integer, ByVal mes_ano As String, texto As String, situacao As String)
    With Sheets("BASE_RESUMO")
        format_cell(.Cells(linha_inicio, 1), cor:=RGB(170, 210, 230)).Value = texto
        format_cell(.Cells(linha_inicio, ano_mes + 2), "Currency").Value = get_soma_se(coluna_soma, mes_ano, situacao)
        format_cell(.Cells(linha_inicio, am_count + 3), "Currency").Value = WorksheetFunction.Sum(.Range(.Cells(linha_inicio, 2), .Cells(linha_inicio, am_count + 2)))
    End With
End Sub