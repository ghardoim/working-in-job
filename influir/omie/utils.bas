Public Function tamanhos()
    tamanhos = Array("PP", "P", "M", "G", "GG", "U", "ÚNICO", "34", "35", "36", "37", "38", "39", "40", "42", "44", "46")
End Function

Public Function sub_cores()
    sub_cores = Array("BRANCO", "OFF WHITE", "PRETO", "OFF", "AMARELO", "ROSE", "AZUL", "MARINHO", "RISCA DE GIZ", "CARAMELO")
End Function

Public Function cores()
    cores = Array("AMARELO/OFF", "BEGE", "PRETO/ BRANCO", "OFFWHITE", "OFF WHITHE", "VERMELHO", "ROSA", "AZUL MARINHO", "CINZA", _
                    "VERDE", "MARFIM", "LISTRADO", "DIJON", "CORAL", "OFF/CARAMELO", "ESTAMPA CLARA", "ESTAMPA ESCURA", "DOURADO", _
                    "TOMATE", "ROXO", "AREIA", "MARROM", "COBRE", "XADREZ", "LARANJA", "SALMAO", "BRUMA", "NATURAL", "VERMELHA", _
                    "PRETO/ OFF WHITE")
End Function

Sub set_atributo(ByVal valor As String, ByVal linha As Integer, c_desc As Integer, c_valor As Integer, c_produto As Integer, _
                    sheet_name As String)
    With ThisWorkbook.Sheets(sheet_name)
        If InStr(UCase(.Cells(linha, c_desc).Value), valor) <> 0 Then
            .Cells(linha, c_valor).Value = valor
            With .Cells(linha, c_produto)
                .Value = Trim(Split(UCase(.Value), valor)(0))
                If Right(.Value, 1) = "-" Then .Value = Trim(Left(.Value, Len(.Value) - 1))
            End With
        End If
    End With
End Sub

Sub remove_acento(celulas As Range)
    For Each letra In Array(Array("Ã", "A"), Array("Á", "A"), Array("Â", "A"), Array("É", "E"), Array("Ê", "E"), _
                            Array("Í", "I"), Array("Ô", "O"), Array("Ó", "O"), Array("Ú", "U"), Array("Ç", "C"))
        celulas.Replace What:=letra(0), Replacement:=letra(1)
    Next
End Sub

Function indice_corresp(ByVal isso As String, essa_coluna As String, naquela_coluna As String, sheet_name As String) As String
    With WorksheetFunction
        indice_corresp = .Index(Sheets(sheet_name).Range(naquela_coluna), .Match(isso, Sheets(sheet_name).Range(essa_coluna), 0))
    End With
End Function