Private Sub UserForm_Initialize()
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
End Sub