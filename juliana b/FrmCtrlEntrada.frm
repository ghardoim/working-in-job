Private Sub UserForm_Activate()
  ThisWorkbook.Activate
  Dim nome As String: nome = ThisWorkbook.Sheets("AUXILIAR").Name

  Me.cbx_advogado.RowSource = nome & "!B2:B" & ultima_linha(nome, "B") - 1
  Me.cbx_tipo.RowSource = nome & "!C2:C" & ultima_linha(nome, "C") - 1
  Me.txt_imposto.Value = 12

  planilha.Activate
  Call atualiza_total(Me.lbl_total, planilha.Sheets("ENTRADA"), "J")
  Call listar(Me.lst_entrada, "ENTRADA", 10)
End Sub

Private Sub btn_adicionar_Click()

  If Not eh_valido(Me.txt_cliente) Then Call msg_de_nao_preenchido("CLIENTE"): Exit Sub
  If Not eh_valido(Me.txt_valor_ref) Then Call msg_de_nao_preenchido("VALOR"): Exit Sub
  If Not eh_valido(Me.cbx_advogado) Then Call msg_de_nao_preenchido("ADVOGADO"): Exit Sub
  If Not eh_valido(Me.cbx_tipo) Then Call msg_de_nao_preenchido("TIPO DE ENTRADA"): Exit Sub

  With planilha.Sheets("ENTRADA")

    Dim ultLin_entrada As Integer: ultLin_entrada = ultima_linha(.Name)
    Dim imposto As Double: imposto = Me.txt_valor_ref.Value * (Me.txt_imposto.Value / 100)
    Dim valor_pago As Double: If Me.txt_valor_pag.Value <> "" Then valor_pago = Me.txt_valor_pag.Value _
                      Else valor_pago = 0: imposto = 0: .Range("A" & ultLin_entrada & ":J" & ultLin_entrada).Interior.Color = 255
        
    .Cells(ultLin_entrada, 1) = Me.cbx_advogado.Value
    .Cells(ultLin_entrada, 2) = UCase(Me.txt_cliente.Value)
    .Cells(ultLin_entrada, 3) = Me.cbx_tipo.Value
    .Cells(ultLin_entrada, 4) = Me.txt_vencimento.Value

    If ckb_boleto Then .Cells(ultLin_entrada, 5) = "-"
    If ckb_nfe Then .Cells(ultLin_entrada, 6) = "-"

    .Columns("G:J").NumberFormat = "$#,##0.00"
    .Cells(ultLin_entrada, 7) = Me.txt_valor_ref.Value
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

  With planilha.Sheets("ENTRADA")
    planilha.Sheets("AUXILIAR").Range("G1").CurrentRegion.Clear

    .UsedRange.AutoFilter
    If Me.txt_cliente <> "" Then Call .UsedRange.AutoFilter(2, "*" & UCase(Me.txt_cliente) & "*")
    If Me.cbx_advogado <> "" Then Call .UsedRange.AutoFilter(1, Me.cbx_advogado)
    If Me.cbx_tipo <> "" Then Call .UsedRange.AutoFilter(3, Me.cbx_tipo)
    .Range("A1").CurrentRegion.Copy planilha.Sheets("AUXILIAR").Range("A1")

    Call listar(Me.lst_entrada, "AUXILIAR", 10)
  End With

  Me.lst_entrada.SetFocus
  Call atualiza_total(Me.lbl_total, planilha.Sheets("AUXILIAR"), "J")
  Call limpar_campos
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