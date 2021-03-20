Private Sub UserForm_Activate()
  Call atualiza_total(Me.lbl_total, planilha.Sheets("ENTRADA"), "J")
  Call listar(Me.lst_entrada, "ENTRADA", 10)
End Sub

Private Sub btn_adicionar_Click()

  If Not eh_valido(Me.txt_cliente) Then Call msg_de_nao_preenchido("CLIENTE"): Exit Sub
  If Not eh_valido(Me.txt_valor_ref) Then Call msg_de_nao_preenchido("VALOR"): Exit Sub
  If Not eh_valido(Me.cbx_advogado) Then Call msg_de_nao_preenchido("ADVOGADO"): Exit Sub
  If Not eh_valido(Me.cbx_tipo) Then Call msg_de_nao_preenchido("TIPO DE ENTRADA"): Exit Sub

  With planilha.Sheets("ENTRADA")

    Dim ultLin_entrada As Integer: ultLin_entrada = ultima_linha(.Name, plan:=planilha)
    Dim valor As Double: valor = Me.txt_valor_ref.Value
    Dim imposto As Double: imposto = valor * (Me.txt_imposto.Value / 100)
    Dim valor_pago As Double: If Me.txt_valor_pag.Value <> "" Then valor_pago = Me.txt_valor_pag.Value _
                      Else valor_pago = 0: imposto = 0: .Range("A" & ultLin_entrada & ":J" & ultLin_entrada).Interior.Color = 255
        
    .Cells(ultLin_entrada, 1) = Me.cbx_advogado.Value
    .Cells(ultLin_entrada, 2) = UCase(Me.txt_cliente.Value)
    .Cells(ultLin_entrada, 3) = Me.cbx_tipo.Value
    .Cells(ultLin_entrada, 4) = Me.txt_vencimento.Value

    With planilha.Sheets("EXTRATO")
      Dim ultLin_extrato As Integer: ultLin_extrato = ultima_linha(.Name, "E", planilha)
      .Cells(ultLin_extrato, 1) = Me.txt_vencimento.Value
      .Cells(ultLin_extrato, 2) = Me.cbx_advogado.Value
      .Cells(ultLin_extrato, 4) = Me.txt_cliente.Value
      .Cells(ultLin_extrato, 5) = Me.cbx_tipo.Value
      With .Cells(ultLin_extrato, 6)
        .Value = valor_pago - imposto
        .Style = "Currency"
      End With
    End With

    If ckb_boleto Then .Cells(ultLin_entrada, 5) = "-"
    If ckb_nfe Then .Cells(ultLin_entrada, 6) = "-"

    .Columns("G:J").NumberFormat = "$#,##0.00"
    .Cells(ultLin_entrada, 7) = valor
    .Cells(ultLin_entrada, 8) = valor_pago
    .Cells(ultLin_entrada, 9) = imposto
    .Cells(ultLin_entrada, 10) = valor_pago - imposto

    Me.lst_entrada.SetFocus

    Call listar(Me.lst_entrada, .Name, 10)
  End With

  Call atualiza_total(Me.lbl_total, planilha.Sheets("ENTRADA"), "J")
  Call limpar_campos
End Sub

Private Sub btn_filtrar_Click()
  If Not eh_valido(Me.txt_cliente) And Not eh_valido(Me.cbx_tipo) And Not eh_valido(Me.cbx_advogado) Then _
      Call msg_de_nao_preenchido("CLIENTE/ADVOGADO/TIPO"): Exit Sub

  Call filtrar("ENTRADA", "*" & Me.txt_cliente & "*", 2, Me.cbx_advogado, 1, Me.cbx_tipo, 3)
  Call listar(Me.lst_entrada, "AUXILIAR", 10)

  Me.lst_entrada.SetFocus
  Call atualiza_total(Me.lbl_total, planilha.Sheets("AUXILIAR"), "J")
  Call limpar_campos
End Sub

Private Sub btn_limpa_filtro_Click()
  Call limpar_filtro("ENTRADA")
  Call listar(Me.lst_saida, "ENTRADA", 10)
End Sub

Private Sub btn_voltar_Click()
  Me.Hide
  FrmCtrlFinanceiro.Show
End Sub

Private Sub spn_imposto_SpinUp()
  Me.txt_imposto.Value = Me.txt_imposto.Value + 1
End Sub

Private Sub spn_imposto_SpinDown()
  Me.txt_imposto.Value = Me.txt_imposto.Value - 1
End Sub

Private Sub limpar_campos()
  Me.txt_cliente = ""
  Me.txt_valor_pag = ""
  Me.txt_valor_ref = ""
  Me.txt_vencimento = ""
  Me.cbx_advogado = ""
  Me.cbx_tipo = ""
  Me.ckb_boleto = False
  Me.ckb_nfe = False
End Sub