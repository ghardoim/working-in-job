Public Function atualiza_total(lbl As MSForms.Label, aba As Worksheet, coluna As String) As Double
  atualiza_total = WorksheetFunction.Sum(aba.Range(coluna & ":" & coluna))
  lbl.Caption = "R$ " & atualiza_total
End Function

Public Sub filtrar(filtro_1 As String, col_1 As Integer, filtro_2 As String, col_2 As Integer, abaName As String)
  Call limpar_filtro
  With planilha.Sheets(abaName)
    .UsedRange.AutoFilter
    Call .UsedRange.AutoFilter(col_1, filtro_1)
    Call .UsedRange.AutoFilter(col_2, filtro_2)
    .Range("A1").CurrentRegion.Copy planilha.Sheets("AUXILIAR").Range("A1")
  End With
End Sub