Sub get_produtos()
    Call liga_desliga(False)
    Dim base_neriage As Workbook: Set base_neriage = Workbooks.Open(Application.GetOpenFilename("Excel Files (*.xlsx), *"))
    Dim ultima_linha As Integer: ultima_linha = base_neriage.Sheets(1).Range("A1048576").End(xlUp).Row - 1
    With ThisWorkbook.Sheets("BASE_PRODUTOS")
        .Range("A6:L" & ultima_linha + 3) = base_neriage.Sheets(1).Range("A3:L" & ultima_linha).Value
        base_neriage.Close (False)

        For linha = 6 To ultima_linha + 3
            On Error Resume Next
            .Cells(linha, 13).Value = Trim("'" & Split(.Cells(linha, 1), "-")(0))
            For Each tamanho In tamanhos
                descricao = Split(.Cells(linha, 2), " ")
                If descricao(UBound(descricao)) = tamanho Then .Cells(linha, 16).Value = tamanho
            Next
            On Error GoTo 0
            Call set_atributo("ACERVO", linha, 14, "BASE_PRODUTOS")
            Call set_atributo("PILOTO", linha, 14, "BASE_PRODUTOS")
            For Each cor In cores
                Call set_atributo(cor, linha, 15, "BASE_PRODUTOS")
            Next
            Call set_atributo("ÚNICO", linha, 16, "BASE_PRODUTOS")
        Next
    End With
    Call MsgBox("agora todos os produtos da planilha escolhida estão aqui! :D", vbInformation, "Base Atualizada")
    Call liga_desliga(True)
End Sub

Sub drop_produtos()
    Sheets("BASE_PRODUTOS").Rows("6:1048576").Delete
End Sub