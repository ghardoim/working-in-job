Sub set_atributo(ByVal valor As String, ByVal linha As Integer, coluna As Integer, sheets_name As String)
    With ThisWorkbook.Sheets(sheets_name)
        If InStr(UCase(.Cells(linha, 1)), valor) <> 0 Then .Cells(linha, coluna).Value = valor
    End With
End Sub