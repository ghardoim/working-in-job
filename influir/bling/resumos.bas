Sub set_resumo()
    Call liga_desliga(False)
    Dim linha_inicio As Integer: linha_inicio = 9
    situacoes = all_unique("W", "BASE_VENDAS")
    canais = all_unique("Y", "BASE_VENDAS")
    am = all_unique("Q", "BASE_VENDAS")

    With Sheets("BASE_RESUMO")

        format_cell(.Cells(5, UBound(am) + 3), cor:=RGB(170, 210, 230)).Value = "total"
        For ano_mes = 0 To UBound(am)

            format_cell(.Cells(6, ano_mes + 2), cor:=RGB(180, 250, 120)).Value = "'" & am(ano_mes)
            format_cell(.Cells(5, ano_mes + 2), cor:=RGB(170, 210, 230)).Value = MonthName(Right(am(ano_mes), 2), True)

            Call total_ano_mes(linha_inicio, "F:F", am, ano_mes, "Total venda Captada Liquida", "*")
            Call total_ano_mes(linha_inicio + 1, "G:G", am, ano_mes, "Atendido Bruto", "Atendido")
            Call total_ano_mes(linha_inicio + 2, "F:F", am, ano_mes, "Atendido Líquido", "Atendido")
            format_cell(.Cells(linha_inicio + 3, 1), cor:=RGB(170, 210, 230)).Value = "Atendido % Desconto"
            format_cell(.Cells(linha_inicio + 4, 1), cor:=RGB(170, 210, 230)).Value = "Participação % Atendido"

            On Error Resume Next
            format_cell(.Cells(linha_inicio + 3, ano_mes + 2), "Percent").Value = (.Cells(linha_inicio + 1, ano_mes + 2) - .Cells(linha_inicio + 2, ano_mes + 2)) / .Cells(linha_inicio + 1, ano_mes + 2)
            format_cell(.Cells(linha_inicio + 4, ano_mes + 2), "Percent").Value = .Cells(linha_inicio + 2, ano_mes + 2) / .Cells(linha_inicio, ano_mes + 2)

            format_cell(.Cells(linha_inicio + 3, UBound(am) + 3), "Percent").Value = (.Cells(linha_inicio + 1, UBound(am) + 3) - .Cells(linha_inicio + 2, UBound(am) + 3)) / .Cells(linha_inicio + 1, UBound(am) + 3)
            format_cell(.Cells(linha_inicio + 4, UBound(am) + 3), "Percent").Value = .Cells(linha_inicio + 2, UBound(am) + 3) / .Cells(linha_inicio, UBound(am) + 3)

            Dim linha_inicio_bloco As Integer: linha_inicio_bloco = 20
            Call soma_por_situacao("Atendido", linha_inicio_bloco, linha_inicio, canais, am, ano_mes)
            Call soma_por_situacao("Em aberto", linha_inicio_bloco, linha_inicio, canais, am, ano_mes)
            Call soma_por_situacao("Em andamento", linha_inicio_bloco, linha_inicio, canais, am, ano_mes)
            Call soma_por_situacao("Cancelado", linha_inicio_bloco, linha_inicio, canais, am, ano_mes)

            Call soma_por_situacao("Atendido", linha_inicio_bloco, linha_inicio, canais, am, ano_mes, "G:G", " VENDA BRUTA", False)
            Call soma_por_situacao("Atendido", linha_inicio_bloco, linha_inicio, canais, am, ano_mes, "H:H", " DESCONTO REALIZADO", False)

            format_cell(.Cells(linha_inicio_bloco - 1, 1), cor:=RGB(170, 210, 230)).Value = "Atendido % DESCONTO REALIZADO"
            format_cell(.Cells(linha_inicio_bloco, 1), cor:=RGB(210, 210, 230)).Value = "TOTAL % DESCONTO REALIZADO"
            format_cell(.Cells(linha_inicio_bloco, ano_mes + 2), "Percent").Value = get_soma_se("H:H", am(ano_mes), "Atendido") / get_soma_se("G:G", am(ano_mes), "Atendido")
            format_cell(.Cells(linha_inicio_bloco, UBound(am) + 3), "Percent").Value = get_soma_se("H:H", am(ano_mes), "Atendido") / get_soma_se("G:G", am(ano_mes), "Atendido")
            For Each canal In canais
                linha_inicio_bloco = linha_inicio_bloco + 1
                format_cell(.Cells(linha_inicio_bloco, 1), cor:=RGB(210, 210, 230)).Value = canal
                format_cell(.Cells(linha_inicio_bloco, ano_mes + 2), "Percent").Value = get_soma_se("H:H", am(ano_mes), "Atendido", canal) / get_soma_se("G:G", am(ano_mes), "Atendido", canal)
                format_cell(.Cells(linha_inicio_bloco, UBound(am) + 3), "Percent").Value = get_soma_se("H:H", situacao:="Atendido", canal:=canal) / get_soma_se("G:G", situacao:="Atendido", canal:=canal)
            Next
            On Error GoTo 0
        Next
    End With
    Call MsgBox("agora todos os resumos estão aqui! :D", vbInformation, "Resumo Atualizado")
    Call liga_desliga(True)
End Sub

Function get_soma_se(coluna_soma As String, Optional ByVal ano_mes As String = "*", Optional ByVal situacao As String = "*", Optional ByVal canal As String = "*") As Double
    get_soma_se = WorksheetFunction.SumIfs(Sheets("BASE_VENDAS").Range(coluna_soma), Sheets("BASE_VENDAS").Range("Q:Q"), ano_mes, Sheets("BASE_VENDAS").Range("W:W"), situacao, Sheets("BASE_VENDAS").Range("Y:Y"), canal)
End Function

Sub soma_por_situacao(situacao As String, ByRef linha_inicio_bloco As Integer, linha_inicio As Integer, canais As Variant, am As Variant, ByVal ano_mes As Integer, Optional coluna_soma As String = "F:F", Optional complemento As String = "", Optional need_percent As Boolean = True)
    Dim linha_percent As Integer: linha_percent = linha_inicio_bloco - 1

    Call header_situacao(situacao, linha_inicio_bloco, complemento)
    Call soma_por_canal(linha_inicio_bloco, canais, am, coluna_soma, ano_mes, situacao)
    Call total(linha_inicio_bloco, am, canais, ano_mes)
    If need_percent Then Call percentual(linha_percent, linha_inicio, am, ano_mes)
    linha_inicio_bloco = linha_inicio_bloco + 8 - (UBound(canais) + 1)
End Sub

Sub soma_por_canal(ByRef linha_inicio_bloco As Integer, canais As Variant, am As Variant, coluna_soma As String, ByVal ano_mes As Integer, ByVal situacao As String)
    With Sheets("BASE_RESUMO")
        For Each canal_venda In canais
            format_cell(.Cells(linha_inicio_bloco + 1, 1), cor:=RGB(210, 210, 230)).Value = canal_venda
            format_cell(.Cells(linha_inicio_bloco + 1, ano_mes + 2), "Currency").Value = get_soma_se(coluna_soma, am(ano_mes), situacao, canal_venda)
            format_cell(.Cells(linha_inicio_bloco + 1, UBound(am) + 3), "Currency").Value = WorksheetFunction.Sum(.Range(.Cells(linha_inicio_bloco + 1, 2), .Cells(linha_inicio_bloco + 1, UBound(am) + 2)))
            linha_inicio_bloco = linha_inicio_bloco + 1
        Next
    End With
End Sub

Sub total(linha_inicio_bloco As Integer, am As Variant, canais As Variant, ByVal ano_mes As Integer)
    With Sheets("BASE_RESUMO")
        format_cell(.Cells(linha_inicio_bloco - (UBound(canais) + 1), ano_mes + 2), "Currency").Value = WorksheetFunction.Sum(.Range(.Cells(linha_inicio_bloco + 1 - (UBound(canais) + 1), ano_mes + 2), .Cells(linha_inicio_bloco + (UBound(canais) + 1), ano_mes + 2)))
        format_cell(.Cells(linha_inicio_bloco - (UBound(canais) + 1), UBound(am) + 3), "Currency").Value = WorksheetFunction.Sum(.Range(.Cells(linha_inicio_bloco + 1 - (UBound(canais) + 1), ano_mes + 3), .Cells(linha_inicio_bloco + (UBound(canais) + 1), ano_mes + 3)))
    End With
End Sub

Sub percentual(linha_percent As Integer, linha_inicio As Integer, am As Variant, ByVal ano_mes As Integer)
    With Sheets("BASE_RESUMO")
        On Error Resume Next
        format_cell(.Cells(linha_percent, ano_mes + 2), "Percent") = .Cells(linha_percent + 1, ano_mes + 2) / .Cells(linha_inicio, ano_mes + 2)
        format_cell(.Cells(linha_percent, UBound(am) + 3), "Percent") = .Cells(linha_percent + 1, UBound(am) + 3) / .Cells(linha_inicio, UBound(am) + 3)
        On Error GoTo 0
    End With
End Sub

Sub header_situacao(ByVal situacao As String, linha_inicio_bloco As Integer, Optional complemento As String = "")
    format_cell(Sheets("BASE_RESUMO").Cells(linha_inicio_bloco, 1), cor:=RGB(210, 210, 230)).Value = "TOTAL" & complemento
    format_cell(Sheets("BASE_RESUMO").Cells(linha_inicio_bloco - 1, 1), cor:=RGB(170, 210, 230)).Value = situacao & complemento
End Sub

Sub total_ano_mes(linha_inicio As Integer, coluna_soma As String, am As Variant, ByVal ano_mes As Integer, texto As String, situacao As String)
    With Sheets("BASE_RESUMO")
        format_cell(.Cells(linha_inicio, 1), cor:=RGB(170, 210, 230)).Value = texto
        format_cell(.Cells(linha_inicio, ano_mes + 2), "Currency").Value = get_soma_se(coluna_soma, am(ano_mes), situacao)
        format_cell(.Cells(linha_inicio, UBound(am) + 3), "Currency").Value = WorksheetFunction.Sum(.Range(.Cells(linha_inicio, 2), .Cells(linha_inicio, UBound(am) + 2)))
    End With
End Sub