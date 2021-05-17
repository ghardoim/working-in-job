Private Sub btn_adicionar_Click()
  If Not IsValid(Me.txt_nome) Then Call errorMessage("O NOME"): Exit Sub
  If Not IsValid(Me.txt_valor) Then Call errorMessage("O VALOR"): Exit Sub

  Dim ULTIMA_LINHA As Integer: ULTIMA_LINHA = GetLastRow("Materiais")
  With Sheets("Materiais")
    .Cells(ULTIMA_LINHA, 1).value = "P" & ULTIMA_LINHA - 1 & Int(Rnd * 10)
    .Cells(ULTIMA_LINHA, 2).value = Me.txt_nome.value
    .Cells(ULTIMA_LINHA, 3).value = Me.txt_valor.value
  End With

  Call limpar(Me.Controls)
  Call atualizar(Me.lst_produto, "Materiais")
End Sub

Private Sub btn_alterar_Click()
  Call alterar(Me.lst_produto, "O PRODUTO", "Materiais", Me.txt_nome.value, Me.txt_valor.value)
  Call limpar(Me.Controls)
End Sub

Private Sub lst_produto_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
  Call removerItem(Me.lst_produto, "Materiais")
End Sub

Private Sub UserForm_Activate()
  Call atualizar(Me.lst_produto, "Materiais")
End Sub

Private Sub UserForm_Terminate()
  Call atualizar(FrmOrcamento.lst_procedimento, "Procedimentos")
End Sub