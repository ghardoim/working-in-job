Sub set_resumo()
    situacoes = all_unique("S", "BASE_VENDAS")
    canais = all_unique("U", "BASE_VENDAS")
    am = all_unique("M", "BASE_VENDAS")
    
    With Sheets("BASE_RESUMO")
    
        format_cell(.Cells(5, UBound(am) + 3), False, RGB(170, 210, 230)).Value = "total"
        For ano_mes = 0 To UBound(am)
            
            format_cell(.Cells(6, ano_mes + 2), False, RGB(180, 250, 120)).Value = "'" & am(ano_mes)
            format_cell(.Cells(5, ano_mes + 2), False, RGB(170, 210, 230)).Value = MonthName(Right(am(ano_mes), 2), True)
            
            Dim linha_inicio_bloco As Integer: linha_inicio_bloco = 9
            For Each situacao In situacoes
                format_cell(.Cells(linha_inicio_bloco, 1), False, RGB(210, 210, 230)).Value = "TOTAL"
                format_cell(.Cells(linha_inicio_bloco - 1, 1), False, RGB(170, 210, 230)).Value = situacao
                
                For Each canal_venda In canais
                    format_cell(.Cells(linha_inicio_bloco + 1, 1), False, RGB(210, 210, 230)).Value = canal_venda
                    
                    format_cell(.Cells(linha_inicio_bloco + 1, ano_mes + 2)).Value = WorksheetFunction.SumIfs(Sheets("BASE_VENDAS").Range("E:E"), Sheets("BASE_VENDAS").Range("M:M"), am(ano_mes), Sheets("BASE_VENDAS").Range("S:S"), situacao, Sheets("BASE_VENDAS").Range("U:U"), canal_venda)
                    format_cell(.Cells(linha_inicio_bloco + 1, UBound(am) + 3)).Value = WorksheetFunction.Sum(.Range(.Cells(linha_inicio_bloco + 1, 2), .Cells(linha_inicio_bloco + 1, UBound(am) + 2)))
                    
                    linha_inicio_bloco = linha_inicio_bloco + 1
                Next
                format_cell(.Cells(linha_inicio_bloco - (UBound(canais) + 1), ano_mes + 2)).Value = WorksheetFunction.Sum(.Range(.Cells(linha_inicio_bloco + 1 - (UBound(canais) + 1), ano_mes + 2), .Cells(linha_inicio_bloco + (UBound(canais) + 1), ano_mes + 2)))
                format_cell(.Cells(linha_inicio_bloco - (UBound(canais) + 1), UBound(am) + 3)).Value = WorksheetFunction.Sum(.Range(.Cells(linha_inicio_bloco + 1 - (UBound(canais) + 1), ano_mes + 3), .Cells(linha_inicio_bloco + (UBound(canais) + 1), ano_mes + 3)))
                
                linha_inicio_bloco = linha_inicio_bloco + 5 - (UBound(canais) + 1)
            Next
        Next
    End With
    Call MsgBox("agora todos os resumos estão aqui! :D", vbInformation, "Resumo Atualizado")
End Sub