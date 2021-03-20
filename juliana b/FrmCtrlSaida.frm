Private Sub UserForm_Activate()
  Call atualiza_total(Me.lbl_total, planilha.Sheets("SAÍDA"), "F")
  Call listar(Me.lst_saida, "SAÍDA", 6, "F")
End Sub

Private Sub btn_adicionar_Click()
  If Not eh_valido(Me.txt_despesa) Then Call msg_de_nao_preenchido("DESPESA", "A"): Exit Sub
  If Not eh_valido(Me.txt_valor) Then Call msg_de_nao_preenchido("VALOR"): Exit Sub
  If Not eh_valido(Me.cbx_tipo) Then Call msg_de_nao_preenchido("TIPO DE DESPESA"): Exit Sub

  With planilha.Sheets("SAÍDA")
    Dim ultLin_saida As Integer: ultLin_saida = ultima_linha(.Name, "F", planilha)
    Dim opcao As Object: Set opcao = Nothing

    .Cells(ultLin_saida, 1) = Me.txt_data.Value
    .Cells(ultLin_saida, 2) = UCase(Me.cbx_funcionario.Value)
    .Cells(ultLin_saida, 3) = UCase(Me.txt_cliente.Value)
    .Cells(ultLin_saida, 4) = Me.cbx_tipo.Value
    .Cells(ultLin_saida, 5) = UCase(Me.txt_despesa.Value)

    With planilha.Sheets("EXTRATO")
      Dim ultLin_extrato As Integer: ultLin_extrato = ultima_linha(.Name, "E", planilha)
      .Cells(ultLin_extrato, 1) = Me.txt_data.Value
      .Cells(ultLin_extrato, 2) = Me.cbx_funcionario.Value
      .Cells(ultLin_extrato, 4) = Me.txt_cliente.Value
      .Cells(ultLin_extrato, 5) = Me.txt_despesa.Value
      With .Cells(ultLin_extrato, 6)
        .Value = Me.txt_valor.Value
        .Style = "Currency"
      End With
    End With

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

  Call filtrar("SAÍDA", "*" & Me.txt_cliente & "*", 3, "*" & Me.txt_despesa & "*", 5, Me.cbx_funcionario, 2, cbx_tipo, 4)
  Call listar(Me.lst_saida, "AUXILIAR", 6, "F")

  Me.lst_saida.SetFocus
  Call atualiza_total(Me.lbl_total, planilha.Sheets("AUXILIAR"), "F")
  Call limpar_campos
End Sub

Private Sub btn_limpa_filtro_Click()
  Call limpar_filtro("SAÍDA")
  Call listar(Me.lst_saida, "SAÍDA", 6, "F")
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