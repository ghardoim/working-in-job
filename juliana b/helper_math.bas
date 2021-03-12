Public Function atualiza_total(lbl As MSForms.Label, aba As Worksheet, coluna As String) As Double
  atualiza_total = WorksheetFunction.Sum(aba.Range(coluna & ":" & coluna))
  If Not lbl Is Nothing Then lbl.Caption = "R$ " & atualiza_total
End Function