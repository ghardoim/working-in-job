Sub get_produtos()
    Call liga_desliga(False)
    Dim base_neriage As Workbook: Set base_neriage = Workbooks.Open(Application.GetOpenFilename("Excel Files (*.xlsx), *"))
    Dim ultima_linha As Integer: ultima_linha = base_neriage.Sheets(1).Range("A1048576").End(xlUp).Row - 1
    With ThisWorkbook.Sheets("BASE_PRODUTOS")
        .Range("A6:L" & ultima_linha + 3) = base_neriage.Sheets(1).Range("A3:L" & ultima_linha).value
        base_neriage.Close (False)

        For linha = 6 To ultima_linha + 3
            On Error Resume Next
            .Cells(linha, 13).value = Trim("'" & Split(.Cells(linha, 1), "-")(0))
            For Each tamanho In Array("PP", "P", "M", "G", "GG", "U", "ÚNICO", "34", "35", "36", "37", "38", "39", "40", "42", "44", "46")
                descricao = Split(.Cells(linha, 2), " ")
                If descricao(UBound(descricao)) = tamanho Then .Cells(linha, 16).value = tamanho
            Next
            On Error GoTo 0
            Call set_atributo("ACERVO", linha, 14)
            Call set_atributo("PILOTO", linha, 14)
            For Each cor In Array("AMARELO", "BEGE", "PRETO", "OFF", "OFF WHITE", "OFFWHITE", "OFF WHITHE", "BRANCO", "VERMELHO", "ROSÊ", _
                                    "ROSE", "ROSA", "AZUL", "MARINHO", "AZUL MARINHO", "CINZA", "VERDE", "MARFIM", "LISTRADO", "DIJON", _
                                    "CORAL", "CARAMELO", "ESTAMPA CLARA", "ESTAMPA ESCURA", "DOURADO", "TOMATE", "ROXO", "AREIA", "MARROM", _
                                    "COBRE", "XADREZ", "LARANJA", "SALMÃO", "RISCA DE GIZ", "BRUMA", "NATURAL", "VERMELHA")
                Call set_atributo(cor, linha, 15)
            Next
            Call set_atributo("ÚNICO", linha, 16)
        Next
    End With
    Call MsgBox("agora todos os produtos da planilha escolhida estão aqui! :D", vbInformation, "Base Atualizada")
    Call liga_desliga(True)
End Sub

Sub set_atributo(ByVal valor As String, ByVal linha As Integer, coluna As Integer)
    With ThisWorkbook.Sheets("BASE_PRODUTOS")
        If InStr(UCase(.Cells(linha, 1)), valor) <> 0 Then .Cells(linha, coluna).value = valor
    End With
End Sub

Sub drop_produtos()
    Sheets("BASE_PRODUTOS").Rows("6:1048576").Delete
End Sub