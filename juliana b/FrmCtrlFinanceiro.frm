Private Sub UserForm_Initialize()
  Dim nome As String: nome = ThisWorkbook.Sheets("AUXILIAR").Name
  
  With FrmCtrlEntrada
    .cbx_advogado.RowSource = nome & "!B2:B" & ultima_linha(nome, "B") - 1
    .cbx_tipo.RowSource = nome & "!C2:C" & ultima_linha(nome, "C") - 1

    .txt_imposto.Value = 12
  End With

  With FrmCtrlSaida
    .cbx_funcionario.RowSource = nome & "!A2:A" & ultima_linha(nome) - 1
    .cbx_tipo.RowSource = nome & "!D2:D" & ultima_linha(nome, "D") - 1
  End With

  Me.lbl_total.Caption = "R$ " & ThisWorkbook.Sheets("AUXILIAR").Range("E2")
  Me.lbl_titulo = "BEM - VINDO " & UCase(Application.UserName)
End Sub

Private Sub btn_fechar_arquivo_Click()
  If planilha Is Nothing Then Exit Sub
  planilha.Save
  planilha.Close
  Set planilha = Nothing
End Sub

Private Sub btn_entrada_Click()
  If planilha Is Nothing Then Exit Sub
  planilha.Sheets("ENTRADA").UsedRange.AutoFilter

  Me.Hide
  FrmCtrlEntrada.Show
End Sub

Private Sub btn_saida_Click()
  If planilha Is Nothing Then Exit Sub
  planilha.Sheets("SAÍDA").UsedRange.AutoFilter

  Me.Hide
  FrmCtrlSaida.Show
End Sub

Private Sub btn_novo_arquivo_Click()
  Call novo_arquivo
  ThisWorkbook.Activate
End Sub

Private Sub btn_importar_Click()
  Call abre_arquivo
  ThisWorkbook.Activate
End Sub

Private Sub UserForm_Terminate()
  Call liga_desliga(True)
  If Not planilha Is Nothing Then
    planilha.Save
    planilha.Close
  End If
End Sub

Private Sub btn_total_Click()
  If planilha Is Nothing Then Exit Sub
  Dim i As Integer, j As Integer, total As Double
  
  With FrmCtrlEntrada
    For i = 0 To .cbx_advogado.ListCount - 1
      For j = 0 To .cbx_tipo.ListCount - 1
        Call filtrar(.cbx_advogado.List(i), 1, .cbx_tipo.List(j), 3, "ENTRADA")
        total = atualiza_total(Me.lbl_para_cada_total, planilha.Sheets("AUXILIAR"), "J")
        'escrever numa tabela
        'celula(linha = i + 2, coluna = j + 2)
        MsgBox "ADVOGADO: " & .cbx_advogado.List(i) & " / TIPO: " & .cbx_tipo.List(j) & vbNewLine & " / TIPO: " & total
      Next
    Next
  End With
End Sub