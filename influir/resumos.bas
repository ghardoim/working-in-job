Sub set_resumo()
    Dim linha_inicio As Integer: linha_inicio = 9
    situacoes = all_unique("V", "BASE_VENDAS")
    canais = all_unique("X", "BASE_VENDAS")
    am = all_unique("P", "BASE_VENDAS")

    With Sheets("BASE_RESUMO")

        format_cell(.Cells(5, UBound(am) + 3), cor:=RGB(170, 210, 230)).Value = "total"
        For ano_mes = 0 To UBound(am)

            format_cell(.Cells(6, ano_mes + 2), cor:=RGB(180, 250, 120)).Value = "'" & am(ano_mes)
            format_cell(.Cells(5, ano_mes + 2), cor:=RGB(170, 210, 230)).Value = MonthName(Right(am(ano_mes), 2), True)
            Call total_ano_mes(linha_inicio, "D:D", am, ano_mes, "Realizado")
            Call total_ano_mes(linha_inicio + 1, "E:E", am, ano_mes, "Venda Bruta")
            Call total_ano_mes(linha_inicio + 2, "F:F", am, ano_mes, "Desconto")

            Dim linha_inicio_bloco As Integer: linha_inicio_bloco = 20
            For Each situacao In situacoes
                Dim linha_percent As Integer: linha_percent = linha_inicio_bloco - 1
                Call header_situacao(situacao, linha_inicio_bloco)
                Call soma_por_canal(linha_inicio_bloco, canais, am, "D:D", ano_mes, situacao)
                Call total_e_percentual(linha_inicio_bloco, linha_percent, linha_inicio, am, canais, ano_mes)
                linha_inicio_bloco = linha_inicio_bloco + 9 - (UBound(canais) + 1)
            Next
            For Each situacao In situacoes
                If "Atendido" = situacao Then
                    linha_percent = linha_inicio_bloco - 1
                    Call header_situacao(situacao, linha_inicio_bloco, " VENDA BRUTA")
                    Call soma_por_canal(linha_inicio_bloco, canais, am, "E:E", ano_mes, situacao)
                    Call total_e_percentual(linha_inicio_bloco, linha_percent, linha_inicio + 1, am, canais, ano_mes)

                    linha_inicio_bloco = linha_inicio_bloco + 9 - (UBound(canais) + 1)
                    linha_percent = linha_inicio_bloco - 1

                    Call header_situacao(situacao, linha_inicio_bloco, " DESCONTO REALIZADO")
                    Call soma_por_canal(linha_inicio_bloco, canais, am, "F:F", ano_mes, situacao)
                    Call total_e_percentual(linha_inicio_bloco, linha_percent, linha_inicio + 2, am, canais, ano_mes)
                End If
            Next
        Next
    End With
    Call MsgBox("agora todos os resumos estão aqui! :D", vbInformation, "Resumo Atualizado")
End Sub

Sub soma_por_canal(ByRef linha_inicio_bloco As Integer, canais As Variant, am As Variant, coluna_soma As String, ByVal ano_mes As Integer, ByVal situacao As String)
    With Sheets("BASE_RESUMO")
        For Each canal_venda In canais
            format_cell(.Cells(linha_inicio_bloco + 1, 1), cor:=RGB(210, 210, 230)).Value = canal_venda
            format_cell(.Cells(linha_inicio_bloco + 1, ano_mes + 2), "Currency").Value = WorksheetFunction.SumIfs(Sheets("BASE_VENDAS").Range(coluna_soma), Sheets("BASE_VENDAS").Range("O:O"), am(ano_mes), Sheets("BASE_VENDAS").Range("S:S"), situacao, Sheets("BASE_VENDAS").Range("U:U"), canal_venda)
            format_cell(.Cells(linha_inicio_bloco + 1, UBound(am) + 3), "Currency").Value = WorksheetFunction.Sum(.Range(.Cells(linha_inicio_bloco + 1, 2), .Cells(linha_inicio_bloco + 1, UBound(am) + 2)))     
            linha_inicio_bloco = linha_inicio_bloco + 1
        Next
    End With
End Sub

Sub total_e_percentual(linha_inicio_bloco As Integer, linha_percent As Integer, linha_inicio As Integer, am As Variant, canais As Variant, ByVal ano_mes As Integer)
    With Sheets("BASE_RESUMO")
        format_cell(.Cells(linha_inicio_bloco - (UBound(canais) + 1), ano_mes + 2), "Currency").Value = WorksheetFunction.Sum(.Range(.Cells(linha_inicio_bloco + 1 - (UBound(canais) + 1), ano_mes + 2), .Cells(linha_inicio_bloco + (UBound(canais) + 1), ano_mes + 2)))
        format_cell(.Cells(linha_inicio_bloco - (UBound(canais) + 1), UBound(am) + 3), "Currency").Value = WorksheetFunction.Sum(.Range(.Cells(linha_inicio_bloco + 1 - (UBound(canais) + 1), ano_mes + 3), .Cells(linha_inicio_bloco + (UBound(canais) + 1), ano_mes + 3)))

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

Sub total_ano_mes(linha_inicio As Integer, coluna_soma As String, am As Variant, ByVal ano_mes As Integer, texto As String)
    With Sheets("BASE_RESUMO")
        format_cell(.Cells(linha_inicio, 1), cor:=RGB(170, 210, 230)).Value = texto
        format_cell(.Cells(linha_inicio, ano_mes + 2), "Currency").Value = WorksheetFunction.SumIfs(Sheets("BASE_VENDAS").Range(coluna_soma), Sheets("BASE_VENDAS").Range("O:O"), am(ano_mes))
        format_cell(.Cells(linha_inicio, UBound(am) + 3), "Currency").Value = WorksheetFunction.Sum(.Range(.Cells(linha_inicio, 2), .Cells(linha_inicio, UBound(am) + 2)))
    End With
End Sub