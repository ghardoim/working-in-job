Sub get_produtos()
    Dim base_neriage As Workbook: Set base_neriage = Workbooks.Open(Application.GetOpenFilename("Excel Files (*.xlsx), *"))
    Dim ultima_linha As Integer: ultima_linha = base_neriage.Sheets(1).Range("A1048576").End(xlUp).Row - 1
    With ThisWorkbook.Sheets("BASE_PRODUTOS")
        .Range("A6:L" & ultima_linha + 3) = base_neriage.Sheets(1).Range("A3:L" & ultima_linha).Value
        base_neriage.Close (False)

        For linha = 6 To ultima_linha + 3
            Call set_clasificacao("ACERVO", linha)
            Call set_clasificacao("PILOTO", linha)
        Next
    End With
End Sub

Sub set_clasificacao(classificacao As String, linha As Integer)
    With ThisWorkbook.Sheets("BASE_PRODUTOS")
        If InStr(UCase(.Cells(linha, 1)), classificacao) <> 0 Then .Cells(linha, 13).Value = classificacao
    End With
End Sub