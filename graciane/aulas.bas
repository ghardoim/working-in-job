Sub analise_absenteismo()
  Dim celula As Range
  Dim linha As Long
  Dim qnt_janeiro As Long: qnt_janeiro = 0
  Dim qnt_fevereiro As Long: qnt_fevereiro = 0

  'para cada objeto em uma lista
  For Each celula In Range("A1:A91085")
    linha = celula.Row
    If celula.Value = "JANEIRO" Then
      qnt_janeiro = qnt_janeiro + 1
    ElseIf celula.Value = "FEVEREIRO" Then
      qnt_fevereiro = qnt_fevereiro + 1
    End If
    If celula.Value = "" Then
      Exit For
    End If
  Next
  MsgBox "achei " & qnt_janeiro & " linhas com JANEIRO e " & qnt_fevereiro & " linhas com FEVEREIRO"

  'para um inicio até um final de 1 em 1
  For linha = 1 To 91085 Step 1
    If Range("A" & linha).Value = "JANEIRO" Then
      qnt_janeiro = qnt_janeiro + 1
    ElseIf Range("A" & linha).Value = "FEVEREIRO" Then
      qnt_fevereiro = qnt_fevereiro + 1
    ElseIf Range("A" & linha).Value = "MARÇO" Then
      MsgBox "Achei março"
    ElseIf Range("A" & linha).Value = "ABRIL" Then
    End If
  Next
  MsgBox "achei " & qnt_janeiro & " linhas com JANEIRO e " & qnt_fevereiro & " linhas com FEVEREIRO"

  'enquanto isso for verdadeiro
  linha = 1
  Do While Range("A" & linha).Value <> ""
    If Range("A" & linha).Value = "JANEIRO" Then
      qnt_janeiro = qnt_janeiro + 1
    ElseIf Range("A" & linha).Value = "FEVEREIRO" Then
      qnt_fevereiro = qnt_fevereiro + 1
    End If
    linha = linha + 1
  Loop
  MsgBox "achei " & qnt_janeiro & " linhas com JANEIRO e " & qnt_fevereiro & " linhas com FEVEREIRO"

  'até que isso seja verdadeiro
  linha = 1
  Do Until Range("A" & linha).Value = ""
    If Range("A" & linha).Value = "JANEIRO" Then
      qnt_janeiro = qnt_janeiro + 1
    ElseIf Range("A" & linha).Value = "FEVEREIRO" Then
      qnt_fevereiro = qnt_fevereiro + 1
    End If
    linha = linha + 1
  Loop

  MsgBox "achei " & qnt_janeiro & " linhas com JANEIRO e " & qnt_fevereiro & " linhas com FEVEREIRO"
End Sub