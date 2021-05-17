Function GetLastRow(sheetName As String, Optional Column As String = "A") As Integer
  GetLastRow = Sheets(sheetName).Range(Column & "1048576").End(xlUp).Row + 1
End Function

Function GetLastColumn(sheetName As String) As String
  GetLastColumn = Split(Sheets(sheetName).Range("A1").End(xlToRight).Address, "$")(1)
End Function

Function IsValid(txtBox As MSForms.TextBox) As Boolean
  IsValid = True: If txtBox.value = "" Or txtBox.value = 0 Then txtBox.SetFocus: IsValid = False
End Function

Function total(sheetName As String) As Double
  total = WorksheetFunction.Sum(Sheets(sheetName).Range("C:C"))
End Function

Function GetTaxa(cmbBox As String, columnTax As String, columnValue As Integer) As Double
  With Sheets("Taxas")
    Dim linha As Integer: linha = .Range(columnTax & ":" & columnTax).Find(cmbBox).Row
    GetTaxa = .Cells(linha, columnValue).value
  End With
End Function