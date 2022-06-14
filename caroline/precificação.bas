Private Sub Worksheet_Change(ByVal Target As Range)
    Application.EnableEvents = False

    On Error Resume Next
    With WorksheetFunction

        If 4 = Target.Row And 3 = Target.Column Then
            codigo = Cells(4, 3).Value
            site = .Index(Sheets("BASE").Range("C:C"), .Match(codigo, Sheets("BASE").Range("F:F"), 0))
            descricao = .Index(Sheets("BASE").Range("G:G"), .Match(codigo, Sheets("BASE").Range("F:F"), 0))
            valor_material = .Index(Sheets("BASE").Range("N:N"), .Match(codigo, Sheets("BASE").Range("F:F"), 0))
            valor_energia = .Index(Sheets("BASE").Range("O:O"), .Match(codigo, Sheets("BASE").Range("F:F"), 0))
            variacao = .Index(Sheets("BASE").Range("U:U"), .Match(codigo, Sheets("BASE").Range("F:F"), 0))
            chave_busca_qnt = .Index(Sheets("BASE").Range("H:H"), .Match(codigo, Sheets("BASE").Range("F:F"), 0))
            quantidade = .Index(Sheets("QUANT").Range("H:H"), .Match(codigo, Sheets("QUANT").Range("F:F"), 0))

            Cells(5, 3).Value = .VLookup(site, Range("G5:H9"), 2, 0)
            Cells(6, 2).Value = descricao
            Cells(11, 3).Value = valor_material
            Cells(14, 3).Value = valor_material
            Cells(14, 3).Borders.Color = IIf(0 = variacao, 255, 0)
            Cells(16, 3).Value = valor_energia
            Cells(21, 3).Value = quantidade

        ElseIf (18 = Target.Row Or 20 = Target.Row) And 3 = Target.Column Then
            tipo_caminhao = Cells(18, 3).Value
            cidade = Cells(20, 3).Value
            nome_aba = "Tabela Frete At. 23.03.22"
            maximo = .MaxIfs(Sheets(nome_aba).Range("O:O"), Sheets(nome_aba).Range("F:F"), cidade, Sheets(nome_aba).Range("B:B"), tipo_caminhao)
            Cells(22, 3).Value = maximo

        ElseIf 3 = Target.Column Then
            Cells(25, 3).Value = Cells(8, 3).Value - Cells(14, 3).Value - Cells(16, 3).Value - Cells(23, 3).Value
            Cells(26, 3).Value = Cells(25, 3).Value / Cells(8, 3).Value
        End If
    End With

    On Error GoTo 0
    Application.EnableEvents = True
End Sub