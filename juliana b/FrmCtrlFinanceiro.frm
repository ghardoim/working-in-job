Private Sub UserForm_Initialize()
  Dim nome As String: nome = ThisWorkbook.Sheets("AUX").Name
  
  With FrmCtrlEntrada
    .cbx_advogado.RowSource = nome & "!B2:B" & ultima_linha(nome, "B") - 1
    .cbx_tipo.RowSource = nome & "!C2:C" & ultima_linha(nome, "C") - 1

    .txt_imposto.Value = 12
  End With

  With FrmCtrlSaida
    .cbx_funcionario.RowSource = nome & "!A2:A" & ultima_linha(nome) - 1
    .cbx_tipo.RowSource = nome & "!D2:D" & ultima_linha(nome, "D") - 1
  End With

  Me.lbl_total.Caption = "R$ " & ThisWorkbook.Sheets("AUX").Range("E2")
  Me.lbl_titulo = "BEM - VINDO " & UCase(Application.UserName)
End Sub

Private Sub btn_fechar_arquivo_Click()
  If planilha Is Nothing Then Exit Sub
  planilha.Save
  planilha.Close
  Set planilha = Nothing
  Call MsgBox("PLANILHA DE CONTROLE FECHADA COM SUCESSO", vbInformation, "SUCESSO")
End Sub

Private Sub btn_entrada_Click()
  If planilha Is Nothing Then Exit Sub
  Call limpar_filtro("ENTRADA")
  Me.Hide
  FrmCtrlEntrada.Show
End Sub

Private Sub btn_saida_Click()
  If planilha Is Nothing Then Exit Sub
  Call limpar_filtro("SAÍDA")
  Me.Hide
  FrmCtrlSaida.Show
End Sub

Private Sub btn_novo_arquivo_Click()
  Call novo_arquivo
  ThisWorkbook.Activate
  Call MsgBox("NOVA PLANILHA DE CONTROLE CRIADA COM SUCESSO", vbInformation, "SUCESSO")
End Sub

Private Sub btn_importar_Click()
  Call abre_arquivo
  ThisWorkbook.Activate
  Call MsgBox("PLANILHA DE CONTROLE IMPORTADA COM SUCESSO", vbInformation, "SUCESSO")
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

  If planilha.PivotCaches.Count > 0 Then planilha.Sheets("RESULTADO").UsedRange.ClearContents

  Call cria_tabela("ENTRADA", 1, "B", "ADVOGADO", 2)
  Call cria_tabela("SAÍDA", 4, "E")  

  Call MsgBox("TOTAL CALCULADO COM SUCESSO", vbInformation, "SUCESSO")
End Sub