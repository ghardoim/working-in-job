Private Sub UserForm_Activate()
  planilha.Activate
  Call atualiza_total(Me.lbl_total, planilha.Sheets("SAÍDA"), "F")
  Call listar(Me.lst_saida, "SAÍDA", 6, "F")
End Sub

Private Sub btn_adicionar_Click()
  If Not eh_valido(Me.txt_despesa) Then Call msg_de_nao_preenchido("DESPESA", "A"): Exit Sub
  If Not eh_valido(Me.txt_valor) Then Call msg_de_nao_preenchido("VALOR"): Exit Sub
  If Not eh_valido(Me.txt_data) Then Call msg_de_nao_preenchido("DATA", "A"): Exit Sub
  If Not eh_valido(Me.cbx_tipo) Then Call msg_de_nao_preenchido("TIPO DE DESPESA"): Exit Sub

  With planilha.Sheets("SAÍDA")
    Dim ultLin_saida As Integer: ultLin_saida = ultima_linha(.Name)
    Dim opcao As Object: Set opcao = Nothing

    .Cells(ultLin_saida, 1) = Me.txt_data.Value
    .Cells(ultLin_saida, 2) = UCase(Me.cbx_funcionario.Value)
    .Cells(ultLin_saida, 3) = UCase(Me.txt_cliente.Value)
    .Cells(ultLin_saida, 4) = Me.cbx_tipo.Value
    .Cells(ultLin_saida, 5) = UCase(Me.txt_despesa.Value)

    .Columns("F:F").NumberFormat = "$#,##0.00"
    .Cells(ultLin_saida, 6) = Me.txt_valor.Value

    Me.lst_saida.SetFocus
    Call listar(Me.lst_saida, .Name, 6, "F")
  End With

  Call atualiza_total(Me.lbl_total, planilha.Sheets("SAÍDA"), "F")
  Call limpar_campos
End Sub

Private Sub btn_filtrar_Click()
  If Not eh_valido(Me.txt_cliente) And Not eh_valido(Me.txt_despesa) And _
    Not eh_valido(Me.cbx_funcionario) And Not eh_valido(Me.cbx_tipo) Then _
        Call msg_de_nao_preenchido("CLIENTE/FUNCIONÁRIO/TIPO/DESPESA"): Exit Sub

  With planilha.Sheets("SAÍDA")
    Call limpar_filtro

    .UsedRange.AutoFilter
    If Me.txt_cliente <> "" Then Call .UsedRange.AutoFilter(3, "*" & Me.txt_cliente & "*")
    If Me.txt_despesa <> "" Then Call .UsedRange.AutoFilter(5, "*" & Me.txt_despesa & "*")
    If Me.cbx_funcionario <> "" Then Call .UsedRange.AutoFilter(2, Me.cbx_funcionario)
    If Me.cbx_tipo <> "" Then Call .UsedRange.AutoFilter(4, Me.cbx_tipo)
    .Range("A1").CurrentRegion.Copy planilha.Sheets("AUXILIAR").Range("A1")

    Call listar(Me.lst_saida, "AUXILIAR", 6, "F")
  End With

  Me.lst_saida.SetFocus
  Call atualiza_total(Me.lbl_total, planilha.Sheets("AUXILIAR"), "F")
  Call limpar_campos
End Sub

Private Sub btn_voltar_Click()
  Me.Hide
  FrmCtrlFinanceiro.Show
End Sub

Private Sub limpar_campos()
  Me.txt_cliente = ""
  Me.txt_valor = ""
  Me.txt_data = ""
  Me.txt_despesa = ""
  Me.cbx_funcionario = ""
  Me.cbx_tipo = ""
End Sub