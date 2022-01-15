Sub get_vendas()
    Call liga_desliga(False)
    Dim base_neriage As Workbook: Set base_neriage = Workbooks.Open(Application.GetOpenFilename("Excel Files (*.xlsx), *"))
    Dim ultima_linha As Integer: ultima_linha = base_neriage.Sheets(1).Range("A1048576").End(xlUp).Row - 1

    With ThisWorkbook.Sheets("BASE_VENDAS")
        .Range("A6:S" & ultima_linha + 3) = base_neriage.Sheets(1).Range("A3:S" & ultima_linha).Value
        base_neriage.Close (False)
        .Range("A6:M" & ultima_linha + 3).SpecialCells(xlCellTypeBlanks).FormulaR1C1 = "=R[-1]C"
        .Range("A6:M" & ultima_linha + 3).Value = .Range("A6:M" & ultima_linha + 3).Value

        For linha = 6 To ultima_linha + 3
            On Error Resume Next
            For Each tamanho In tamanhos
                descricao = Split(.Cells(linha, 9), " ")
                If descricao(UBound(descricao)) = tamanho Then .Cells(linha, 20).Value = tamanho
            Next
            On Error GoTo 0
            For Each cor In cores
                Call set_atributo(cor, linha, 21, "BASE_VENDAS")
            Next
        Next 
    End With
    Call MsgBox("agora todos as vendas da planilha escolhida estão aqui! :D", vbInformation, "Base Atualizada")
    Call liga_desliga(True)
End Sub

Sub drop_vendas()
    Sheets("BASE_VENDAS").Rows("6:1048576").Delete
End Sub