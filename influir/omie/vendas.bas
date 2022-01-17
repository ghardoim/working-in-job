Sub get_vendas()
    Call liga_desliga(False)
    var_cores = cores: var_subcores = sub_cores
    Dim base_neriage As Workbook: Set base_neriage = Workbooks.Open(Application.GetOpenFilename("Excel Files (*.xlsx), *"))
    Dim ultima_linha As Integer: ultima_linha = base_neriage.Sheets(1).Range("A1048576").End(xlUp).Row - 1

    With ThisWorkbook.Sheets("BASE_VENDAS")
        .Range("A6:S" & ultima_linha + 3) = base_neriage.Sheets(1).Range("A3:S" & ultima_linha).Value
        base_neriage.Close (False)
        .Range("A6:M" & ultima_linha + 3).SpecialCells(xlCellTypeBlanks).FormulaR1C1 = "=R[-1]C"
        .Range("A6:M" & ultima_linha + 3).Value = .Range("A6:M" & ultima_linha + 3).Value
        Call remove_acento(.Range("I6:I" & ultima_linha + 3))

        For linha = 6 To ultima_linha + 3

            data_tratada = Split(.Cells(linha, 7).Value, "/")
            If data_tratada(2) = 2022 Then .Cells(linha, 7).Value = CDate(data_tratada(1) & "/" & data_tratada(0) & "/" & data_tratada(2))
             .Cells(linha, 7).Value = CDate(.Cells(linha, 7).Value)

            Call set_atributo("ACERVO", linha, 9, 22, 9, "BASE_VENDAS")
            Call set_atributo("PILOTO", linha, 9, 22, 9, "BASE_VENDAS")
            On Error Resume Next
            For Each tamanho In tamanhos
                descricao = Split(.Cells(linha, 9), " ")
                If descricao(UBound(descricao)) = tamanho Then .Cells(linha, 20).Value = tamanho
            Next
            On Error GoTo 0
            For Each cor In var_cores
                Call set_atributo(cor, linha, 9, 21, 9, "BASE_VENDAS")
            Next
            For Each subcor In var_subcores
                Call set_atributo(subcor, linha, 9, 21, 9, "BASE_VENDAS")
            Next
            If InStr(.Cells(linha, 9).Value, " - ROS ") <> 0 Then
                .Cells(linha, 9).Value = Trim(Split(.Cells(linha, 9).Value, " - ROS ")(0))
                .Cells(linha, 21).Value = "ROSE"
            ElseIf InStr(.Cells(linha, 9).Value, " ROS ") <> 0 Then
                .Cells(linha, 9).Value = Trim(Split(.Cells(linha, 9).Value, " ROS ")(0))
                .Cells(linha, 21).Value = "ROSE"
            End If
            .Cells(linha, 22).Value = Trim(.Cells(linha, 9).Value & " " & .Cells(linha, 21).Value)
        Next
    End With
    Call MsgBox("agora todos as vendas da planilha escolhida estão aqui! :D", vbInformation, "Base Atualizada")
    Call liga_desliga(True)
End Sub

Sub drop_vendas()
    Sheets("BASE_VENDAS").Rows("6:1048576").Delete
End Sub