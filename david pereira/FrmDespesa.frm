Private Sub btn_adicionar_Click()
  If Not IsValid(Me.txt_nome) Then Call errorMessage("NOME"): Exit Sub
  If Not IsValid(Me.txt_valor) Then Call errorMessage("VALOR"): Exit Sub

  Dim ULTIMA_LINHA As Integer: ULTIMA_LINHA = GetLastRow("Despesas")
  With Sheets("Despesas")

    .Cells(ULTIMA_LINHA, 1).value = "P" & ULTIMA_LINHA - 1 & Int(Rnd * 10)
    .Cells(ULTIMA_LINHA, 2).value = Me.txt_nome.value
    .Cells(ULTIMA_LINHA, 3).value = Me.txt_valor.value
  End With

  Call limpar(Me.Controls)
  Call atualizar(Me.lst_despesa, "Despesas")
End Sub

Private Sub btn_alterar_Click()
  Call alterar(Me.lst_despesa, "A DESPESA", "Despesas", Me.txt_nome, Me.txt_valor)
  Call limpar(Me.Controls)
End Sub

Private Sub btn_valorHora_Click()
  FrmValorHora.Show
End Sub

Private Sub lst_despesa_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
  Call removerItem(Me.lst_despesa, "Despesas")
End Sub

Private Sub UserForm_Activate()
  Call atualizar(Me.lst_despesa, "Despesas")
End Sub

Private Sub UserForm_Terminate()
  Call atualizar(FrmOrcamento.lst_procedimento, "Procedimentos")
End Sub