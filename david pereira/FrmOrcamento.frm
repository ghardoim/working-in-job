Private Sub btn_despesa_Click()
  FrmDespesa.Show
End Sub

Private Sub btn_limpar_Click()
  For Each aba In ThisWorkbook.Sheets
    Dim sheetName As String: sheetName = aba.name
    If "Taxas" <> sheetName And "Interface" <> sheetName Then
      Dim ultimaColuna As String: ultimaColuna = GetLastColumn(sheetName)
      Sheets(sheetName).Range("A2:" & ultimaColuna & GetLastRow(sheetName, ultimaColuna)).ClearContents
    End If
  Next
  Call atualizar(Me.lst_procedimento, "Procedimentos")
  Call limpar(Me.Controls)
End Sub

Private Sub btn_produtos_Click()
  FrmProduto.Show
End Sub

Private Sub btn_procedimentos_Click()
  FrmProcedimento.Show
End Sub

Private Sub btn_total_Click()
  Dim imposto As Double: imposto = Sheets("Taxas").Range("E2")
  Dim lucro As Double: lucro = Sheets("Taxas").Range("F2")
  Dim desconto As Double: desconto = GetTaxa(Me.cmb_tipoPagamento, "C", 4)
  Dim taxa As Double: taxa = GetTaxa(Me.cmb_nParcelas, "A", 2)
  Dim total As Double: total = WorksheetFunction.Sum(Sheets("Procedimentos").Range("E:E"))

  Me.txt_desconto.value = desconto * 100 & "%"
  Me.txt_taxas.value = taxa * 100 & "%"

  Me.txt_total.value = Round((total / (100 - (imposto + taxa + (lucro - desconto)))) * 100, 2)
End Sub

Private Sub cmb_tipoPagamento_Change()
  With Me.cmb_nParcelas
    If "CRÉDITO" = Me.cmb_tipoPagamento.value Then .Enabled = True Else .value = "": .Enabled = False
  End With
End Sub

Private Sub UserForm_Activate()
    Me.cmb_tipoPagamento.RowSource = "Taxas!C2:C" & GetLastRow("Taxas", "C") - 1
    Me.cmb_nParcelas.RowSource = "Taxas!A2:A" & GetLastRow("Taxas") - 1
    
    Call atualizar(Me.lst_procedimento, "Procedimentos")
End Sub