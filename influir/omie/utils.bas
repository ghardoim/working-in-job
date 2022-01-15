Public Function tamanhos()
    tamanhos = Array("PP", "P", "M", "G", "GG", "U", "ÚNICO", "34", "35", "36", "37", "38", "39", "40", "42", "44", "46")
End Function
Public Function cores()
    cores = Array("AMARELO", "BEGE", "PRETO", "OFF", "OFF WHITE", "OFFWHITE", "OFF WHITHE", "BRANCO", "VERMELHO", "ROSÊ", _
                    "ROSE", "ROSA", "AZUL", "MARINHO", "AZUL MARINHO", "CINZA", "VERDE", "MARFIM", "LISTRADO", "DIJON", _
                    "CORAL", "CARAMELO", "ESTAMPA CLARA", "ESTAMPA ESCURA", "DOURADO", "TOMATE", "ROXO", "AREIA", "MARROM", _
                    "COBRE", "XADREZ", "LARANJA", "SALMÃO", "RISCA DE GIZ", "BRUMA", "NATURAL", "VERMELHA")
End Function
Sub set_atributo(ByVal valor As String, ByVal linha As Integer, coluna As Integer, sheets_name As String)
    With ThisWorkbook.Sheets(sheets_name)
        If InStr(UCase(.Cells(linha, 1)), valor) <> 0 Then .Cells(linha, coluna).Value = valor
    End With
End Sub