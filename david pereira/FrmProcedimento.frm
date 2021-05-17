Private Sub btn_addProduto_Click()
  If "" = Me.cmb_produtos Then Call errorMessage("O PRODUTO"): Exit Sub
  Dim ULTIMA_LINHA As Integer: ULTIMA_LINHA = GetLastRow("Procedimentos")
  With Me.lst_procedimento
    If -1 = .ListIndex Then Call MsgBox("SELECIONE O PROCEDIMENTO", vbExclamation, ""): Exit Sub

    Dim procedimento As String: procedimento = .List(.ListIndex, 1)
    Dim ID As String: ID = .List(.ListIndex, 0)

    With Sheets("Procedimentos")
      .Cells(ULTIMA_LINHA, 1).value = ID
      .Cells(ULTIMA_LINHA, 2).value = procedimento
      .Cells(ULTIMA_LINHA, 3).value = Me.cmb_produtos.value
      .Cells(ULTIMA_LINHA, 4).value = Me.txt_qnt.value
      .Cells(ULTIMA_LINHA, 5).value = Me.txt_qnt.value * Sheets("Materiais").Cells(Sheets("Materiais").Range("B:B").Find(Me.cmb_produtos.value).Row, 3).value
    End With
  End With
  Call atualizar(Me.lst_procedimento, "Procedimentos")

  Me.cmb_produtos.value = ""
  Me.txt_qnt.value = 1
End Sub

Private Sub btn_adicionar_Click()
  If Not IsValid(Me.txt_nome) Then Call errorMessage("O NOME"): Exit Sub
  If "" = Me.cmb_produtos Then Call errorMessage("O PRODUTO"): Exit Sub

  Dim ULTIMA_LINHA As Integer: ULTIMA_LINHA = GetLastRow("Procedimentos")
  With Sheets("Procedimentos")
    .Cells(ULTIMA_LINHA, 1).value = "P" & ULTIMA_LINHA - 1 & Int(Rnd * 10)
    .Cells(ULTIMA_LINHA, 2).value = Me.txt_nome.value
    .Cells(ULTIMA_LINHA, 3).value = Me.cmb_produtos.value
    .Cells(ULTIMA_LINHA, 4).value = Me.txt_qnt.value
    .Cells(ULTIMA_LINHA, 5).value = Me.txt_qnt.value * Sheets("Materiais").Cells(Sheets("Materiais").Range("B:B").Find(Me.cmb_produtos.value).Row, 3).value
  End With

  Call limpar(Me.Controls)
  Call atualizar(Me.lst_procedimento, "Procedimentos")

  Me.cmb_produtos.value = ""
  Me.txt_qnt.value = 1
End Sub

Private Sub lst_procedimento_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
  Call removerItem(Me.lst_procedimento, "Procedimentos")
End Sub

Private Sub spn_qnt_SpinDown()
  Call spinChange(Me.txt_qnt, 1, False)
End Sub

Private Sub spn_qnt_SpinUp()
  Call spinChange(Me.txt_qnt, 50)
End Sub

Private Sub UserForm_Activate()
  Me.cmb_produtos.RowSource = "Materiais!B2:B" & GetLastRow("Materiais", "B") - 1
  Call atualizar(Me.lst_procedimento, "Procedimentos")
End Sub

Private Sub UserForm_Terminate()
  Call atualizar(FrmOrcamento.lst_procedimento, "Procedimentos")
End Sub