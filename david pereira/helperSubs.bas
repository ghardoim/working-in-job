Sub calculaValorHora(txtBox As MSForms.TextBox, Column As String, totalHorasDias As Integer)
  On Error Resume Next
  With Sheets("Despesas")
    .Range(Column & "1").value = txtBox.value
    .Range("J1").value = total("Materiais") / totalHorasDias
  End With
  On Error GoTo 0
End Sub

Sub spinChange(txtBox As MSForms.TextBox, limitValue As Integer, Optional sequence As Boolean = True)
  With txtBox
    If sequence Then If .value < limitValue Then .value = .value + 1: Exit Sub
    If .value > limitValue Then .value = .value - 1
  End With
End Sub

Sub removerItem(Lista As MSForms.ListBox, sheetName As String)
  With Lista
    If MsgBox("DESEJA REMOVER O PRODUTO: " & .List(.ListIndex, 1), vbYesNo, "EXCLUIR") = vbYes Then
      Call Sheets(sheetName).Range("A:A").Find(.List(.ListIndex, 0)).EntireRow.Delete
    End If
  End With
End Sub

Sub alterar(List As MSForms.ListBox, fieldToChange As String, sheetName As String, _
            Optional nameField As String = "", Optional valueField As String = "")

  With List
    If -1 = .ListIndex Then Call errorMessage(fieldToChange): Exit Sub

    Dim NOME As String: NOME = nameField
    Dim VALOR As Double: VALOR = IIf("" = valueField, 0, valueField)

    If MsgBox("DESEJA ALTERAR " & fieldToChange & ": " & .List(.ListIndex, 1), vbYesNo, "ALTERAR") = vbYes Then

      Dim linha As Integer: linha = Sheets(sheetName).Range("A:A").Find(.List(.ListIndex, 0)).Row

      With Sheets(sheetName)
        If NOME <> "" Then .Cells(linha, 2).value = NOME
        If VALOR <> 0 Then .Cells(linha, 3).value = VALOR
      End With
    End If
  End With
End Sub

Sub atualizar(List As MSForms.ListBox, sheetName As String)
  List.RowSource = sheetName & "!A2:" & GetLastColumn(sheetName) & GetLastRow(sheetName)
End Sub

Sub limpar(Form As Controls)
  For Each Item In Form
    If TypeName(Item) = "TextBox" Then Item.value = ""
  Next
End Sub

Sub errorMessage(Field As String)
  Call MsgBox("INFORME " & Field, vbExclamation, Field & " N�O INFORMADO")
End Sub