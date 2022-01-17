Sub get_produtos()
    Call liga_desliga(False)
    var_cores = cores: var_subcores = sub_cores
    Dim base_neriage As Workbook: Set base_neriage = Workbooks.Open(Application.GetOpenFilename("Excel Files (*.xlsx), *"))
    Dim ultima_linha As Integer: ultima_linha = base_neriage.Sheets(1).Range("A1048576").End(xlUp).Row - 1
    With ThisWorkbook.Sheets("BASE_PRODUTOS")
        .Range("A6:L" & ultima_linha + 3) = base_neriage.Sheets(1).Range("A3:L" & ultima_linha).Value
        base_neriage.Close (False)
        Call remove_acento(.Range("A6:B" & ultima_linha + 3))

        For linha = 6 To ultima_linha + 3
            On Error Resume Next
            .Cells(linha, 13).Value = Trim("'" & Split(.Cells(linha, 1), "-")(0))
            For Each tamanho In tamanhos
                descricao = Split(.Cells(linha, 2), " ")
                If descricao(UBound(descricao)) = tamanho Then .Cells(linha, 16).Value = tamanho
            Next
            On Error GoTo 0

            Call set_atributo("ACERVO", linha, 2, 14, 2, "BASE_PRODUTOS")
            Call set_atributo("PILOTO", linha, 2, 14, 2, "BASE_PRODUTOS")

            For Each cor In var_cores
                Call set_atributo(cor, linha, 2, 15, 2, "BASE_PRODUTOS")
            Next
            For Each subcor In var_subcores
                Call set_atributo(subcor, linha, 2, 15, 2, "BASE_PRODUTOS")
            Next
            Call set_atributo("ÚNICO", linha, 1, 16, 2, "BASE_PRODUTOS")
            .Cells(linha, 17).Value = Trim(.Cells(linha, 2).Value & " " & .Cells(linha, 15).Value)
        Next
        ThisWorkbook.Sheets("BASE_APOIO").Range("D6:T" & ultima_linha + 3) = .Range("A6:Q" & ultima_linha + 3).Value
    End With
    Call MsgBox("agora todos os produtos da planilha escolhida estão aqui! :D", vbInformation, "Base Atualizada")
    Call liga_desliga(True)
End Sub

Sub drop_produtos()
    Sheets("BASE_PRODUTOS").Rows("6:1048576").Delete
End Sub