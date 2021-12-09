Sub set_resumo()
    Dim linha_realizado As Integer: linha_realizado = 9
    situacoes = all_unique("V", "BASE_VENDAS")
    canais = all_unique("X", "BASE_VENDAS")
    am = all_unique("P", "BASE_VENDAS")

    With Sheets("BASE_RESUMO")

        format_cell(.Cells(5, UBound(am) + 3), cor:=RGB(170, 210, 230)).Value = "total"
        For ano_mes = 0 To UBound(am)

            format_cell(.Cells(6, ano_mes + 2), cor:=RGB(180, 250, 120)).Value = "'" & am(ano_mes)
            format_cell(.Cells(5, ano_mes + 2), cor:=RGB(170, 210, 230)).Value = MonthName(Right(am(ano_mes), 2), True)

            format_cell(.Cells(linha_realizado, 1), cor:=RGB(170, 210, 230)).Value = "Realizado"
            format_cell(.Cells(linha_realizado, ano_mes + 2), "Currency").Value = WorksheetFunction.SumIfs(Sheets("BASE_VENDAS").Range("E:E"), Sheets("BASE_VENDAS").Range("P:P"), am(ano_mes))
            format_cell(.Cells(linha_realizado, UBound(am) + 3), "Currency").Value = WorksheetFunction.Sum(.Range(.Cells(linha_realizado, 2), .Cells(linha_realizado, UBound(am) + 2)))

            Dim linha_inicio_bloco As Integer: linha_inicio_bloco = 15
            For Each situacao In situacoes
                format_cell(.Cells(linha_inicio_bloco, 1), cor:=RGB(210, 210, 230)).Value = "TOTAL"
                format_cell(.Cells(linha_inicio_bloco - 1, 1), cor:=RGB(170, 210, 230)).Value = situacao

                Dim linha_percent As Integer: linha_percent = linha_inicio_bloco - 1
                For Each canal_venda In canais

                    format_cell(.Cells(linha_inicio_bloco + 1, 1), cor:=RGB(210, 210, 230)).Value = canal_venda
                    format_cell(.Cells(linha_inicio_bloco + 1, ano_mes + 2), "Currency").Value = WorksheetFunction.SumIfs(Sheets("BASE_VENDAS").Range("E:E"), Sheets("BASE_VENDAS").Range("P:P"), am(ano_mes), Sheets("BASE_VENDAS").Range("V:V"), situacao, Sheets("BASE_VENDAS").Range("X:X"), canal_venda)
                    format_cell(.Cells(linha_inicio_bloco + 1, UBound(am) + 3), "Currency").Value = WorksheetFunction.Sum(.Range(.Cells(linha_inicio_bloco + 1, 2), .Cells(linha_inicio_bloco + 1, UBound(am) + 2)))

                    linha_inicio_bloco = linha_inicio_bloco + 1
                Next
                format_cell(.Cells(linha_inicio_bloco - (UBound(canais) + 1), ano_mes + 2), "Currency").Value = WorksheetFunction.Sum(.Range(.Cells(linha_inicio_bloco + 1 - (UBound(canais) + 1), ano_mes + 2), .Cells(linha_inicio_bloco + (UBound(canais) + 1), ano_mes + 2)))
                format_cell(.Cells(linha_inicio_bloco - (UBound(canais) + 1), UBound(am) + 3), "Currency").Value = WorksheetFunction.Sum(.Range(.Cells(linha_inicio_bloco + 1 - (UBound(canais) + 1), ano_mes + 3), .Cells(linha_inicio_bloco + (UBound(canais) + 1), ano_mes + 3)))

                format_cell(.Cells(linha_percent, ano_mes + 2), "Percent") = .Cells(linha_percent + 1, ano_mes + 2) / .Cells(linha_realizado, ano_mes + 2)
                format_cell(.Cells(linha_percent, UBound(am) + 3), "Percent") = .Cells(linha_percent + 1, UBound(am) + 3) / .Cells(linha_realizado, UBound(am) + 3)
                
                linha_inicio_bloco = linha_inicio_bloco + 5 - (UBound(canais) + 1)
            Next
        Next
    End With
    Call MsgBox("agora todos os resumos estão aqui! :D", vbInformation, "Resumo Atualizado")
End Sub