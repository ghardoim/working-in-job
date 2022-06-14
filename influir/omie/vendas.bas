Sub get_vendas()
    Call liga_desliga(False)
    Call drop_vendas
    var_cores = cores: var_subcores = sub_cores
    Dim base_neriage As Workbook: Set base_neriage = Workbooks.Open(Application.GetOpenFilename("Excel Files (*.xlsx), *"))
    Dim ultima_linha As Integer: ultima_linha = base_neriage.Sheets(1).Range("A1048576").End(xlUp).Row - 1
    With ThisWorkbook.Sheets("BASE_VENDAS")
        .Range("A6:W" & ultima_linha + 3) = base_neriage.Sheets(1).Range("A3:W" & ultima_linha).Value
        base_neriage.Close (False)

        .Range("A6:R" & ultima_linha + 3).SpecialCells(xlCellTypeBlanks).FormulaR1C1 = "=R[-1]C"
        .Range("A6:R" & ultima_linha + 3).Value = .Range("A6:R" & ultima_linha + 3).Value
        Call remove_acento(.Range("I6:I" & ultima_linha + 3))

        .Range("G6:G" & ultima_linha + 3).FormulaR1C1 = "=DATE(RC[9],MONTH(DATEVALUE(RC[11]&""1"")),RC[10])"
        .Range("G6:G" & ultima_linha + 3).Value = .Range("G6:G" & ultima_linha + 3).Value
        .Range("P6:R" & ultima_linha + 3).Delete Shift:=xlToLeft

        For linha = 6 To ultima_linha + 3
            Call set_atributo("ACERVO", linha, 9, 23, 9, "BASE_VENDAS")
            Call set_atributo("PILOTO", linha, 9, 23, 9, "BASE_VENDAS")
            On Error Resume Next
            For Each tamanho In tamanhos
                descricao = Split(.Cells(linha, 9), " ")
                If descricao(UBound(descricao)) = tamanho Then .Cells(linha, 21).Value = tamanho
            Next
            On Error GoTo 0
            For Each cor In var_cores
                Call set_atributo(cor, linha, 9, 22, 9, "BASE_VENDAS")
            Next
            For Each subcor In var_subcores
                Call set_atributo(subcor, linha, 9, 22, 9, "BASE_VENDAS")
            Next
            If InStr(.Cells(linha, 9).Value, " - ROS ") <> 0 Then
                .Cells(linha, 9).Value = Trim(Split(.Cells(linha, 9).Value, " - ROS ")(0))
                .Cells(linha, 22).Value = "ROSE"
            ElseIf InStr(.Cells(linha, 9).Value, " ROS ") <> 0 Then
                .Cells(linha, 9).Value = Trim(Split(.Cells(linha, 9).Value, " ROS ")(0))
                .Cells(linha, 22).Value = "ROSE"
            End If
            .Cells(linha, 23).Value = Trim(.Cells(linha, 9).Value & " " & .Cells(linha, 22).Value)
            .Cells(linha, 24).Value = "'" & Year(.Cells(linha, 7).Value) & "." & Format(Month(.Cells(linha, 7).Value), "00")
            On Error Resume Next
                .Cells(linha, 25).Value = indice_corresp(.Cells(linha, 23).Value & "*", "P:P", "I:I", "BASE_PRODUTOS")
                .Cells(linha, 26).Value = indice_corresp(.Cells(linha, 23).Value & "*", "P:P", "C:C", "BASE_PRODUTOS")
            On Error GoTo 0
        Next
        .Range("P6:S" & ultima_linha + 3).Style = "Currency"
        .Range("Y6:Y" & ultima_linha + 3).Style = "Currency"

    End With
    Call MsgBox("agora todos as vendas da planilha escolhida estão aqui! :D", vbInformation, "Base Atualizada")
    Call liga_desliga(True)
End Sub

Sub drop_vendas()
    Sheets("BASE_VENDAS").Rows("6:1048576").Delete
End Sub