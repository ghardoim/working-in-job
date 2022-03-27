Sub get_produtos()
    Call liga_desliga(False)
    Call drop_produtos
    var_cores = cores: var_subcores = sub_cores
    Dim base_neriage As Workbook: Set base_neriage = Workbooks.Open(Application.GetOpenFilename("Excel Files (*.xlsx), *"))
    Dim ultima_linha As Integer: ultima_linha = base_neriage.Sheets(1).Range("A1048576").End(xlUp).Row - 1
    With ThisWorkbook.Sheets("BASE_PRODUTOS")
        .Range("A6:L" & ultima_linha + 3) = base_neriage.Sheets(1).Range("A3:L" & ultima_linha).Value
        base_neriage.Close (False)

        .Range("A5").AutoFilter Field:=1, Criteria1:="="
        .Range("A6:C" & ultima_linha + 3).SpecialCells(xlCellTypeVisible).Clear
        .Range("A5").AutoFilter

        .Range("A6:G" & ultima_linha + 3).SpecialCells(xlCellTypeBlanks).FormulaR1C1 = "=R[-1]C"
        .Range("A6:G" & ultima_linha + 3).Value = .Range("A6:R" & ultima_linha + 3).Value
        Call remove_acento(.Range("A6:B" & ultima_linha + 3))

        For linha = 6 To ultima_linha + 3
            On Error Resume Next
            For Each tamanho In tamanhos
                descricao = Split(.Cells(linha, 2), " ")
                If descricao(UBound(descricao)) = tamanho Then .Cells(linha, 16).Value = tamanho
            Next
            On Error GoTo 0
            Call set_atributo("ACERVO", linha, 2, 13, 2, "BASE_PRODUTOS")
            Call set_atributo("PILOTO", linha, 2, 13, 2, "BASE_PRODUTOS")
            
            For Each cor In var_cores
                Call set_atributo(cor, linha, 2, 14, 2, "BASE_PRODUTOS")
            Next
            For Each subcor In var_subcores
                Call set_atributo(subcor, linha, 2, 14, 2, "BASE_PRODUTOS")
            Next
            Call set_atributo("ÚNICO", linha, 1, 15, 2, "BASE_PRODUTOS")
            .Cells(linha, 16).Value = Trim(.Cells(linha, 2).Value & " " & .Cells(linha, 14).Value)
        Next
        ThisWorkbook.Sheets("BASE_APOIO").Range("A6:A" & ultima_linha + 3) = .Range("C6:C" & ultima_linha + 3).Value
        ThisWorkbook.Sheets("BASE_APOIO").Range("B6:B" & ultima_linha + 3) = .Range("P6:P" & ultima_linha + 3).Value
    End With

    With ThisWorkbook.Sheets("BASE_APOIO").Sort
            .SortFields.Clear
            .SortFields.Add2 Key:=Range("A:A"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SetRange Rows("5:1048576")
            .Header = xlYes
            .Apply
    End With
    Call MsgBox("agora todos os produtos da planilha escolhida estão aqui! :D", vbInformation, "Base Atualizada")
    Call liga_desliga(True)
End Sub

Sub drop_produtos()
    Sheets("BASE_PRODUTOS").Rows("6:1048576").Delete
End Sub