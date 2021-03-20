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
  If Not planilha Is Nothing Then Exit Sub
  Call novo_arquivo
  ThisWorkbook.Activate
  Call MsgBox("NOVA PLANILHA DE CONTROLE CRIADA COM SUCESSO EM MEUS DOCUMENTOS" & vbNewLine & _
        vbNewLine & "PARA CONTINUAR, ABRA O ARQUIVO GERADO.", vbInformation, "SUCESSO")
End Sub

Private Sub btn_importar_Click()
  If Not planilha Is Nothing Then Exit Sub
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
  On Error Resume Next
  If planilha Is Nothing Then Exit Sub

  If planilha.PivotCaches.Count > 0 Then planilha.Sheets("RESULTADO").UsedRange.ClearContents
  Call cria_tabela("ENTRADA", 1, "VALOR LÍQUIDO", "B", "ADVOGADO", 2, "IMPOSTO", "C")
  Call cria_tabela("SAÍDA", 5, "VALOR", "F")

  Dim fundo_escritorio As Double: fundo_escritorio = Me.lbl_total.Value
  Dim total_bruno As Double: total_bruno = 0
  Dim total_paulo As Double: total_bruno = 0
  Dim total_isa As Double: total_isa = 0

  With planilha.Sheets("RESULTADO")
    With .PivotTables("TOTAL ENTRADA")

      Dim i As Integer
      For i = 0 To FrmCtrlEntrada.cbx_tipo.ListCount - 1
        Dim j As Integer
        For j = 0 To FrmCtrlEntrada.cbx_advogado.ListCount - 1

          Dim advogado As String: advogado = FrmCtrlEntrada.cbx_advogado.List(j)
          Dim valor As Double: valor = 0
          valor = .GetPivotData("VALOR LÍQUIDO", "TIPO", FrmCtrlEntrada.cbx_tipo.List(i), "ADVOGADO", advogado)
                    
          Select Case advogado
            Case "BRUNO"
              total_bruno = add_no_(total_bruno, 0.6, valor)
              total_paulo = add_no_(total_paulo, 0.4, valor)

            Case "PAULO"
              total_paulo = add_no_(total_paulo, 0.6, valor)
              total_bruno = add_no_(total_bruno, 0.4, valor)

            Case "ISABELA"
              total_isa = add_no_(total_isa, 0.5, valor)
              total_paulo = add_no_(total_paulo, 0.25, valor)
              total_bruno = add_no_(total_bruno, 0.25, valor)

            Case "BRUNO & PAULO"
              total_paulo = add_no_(total_paulo, 0.5, valor)
              total_bruno = add_no_(total_bruno, 0.5, valor)

            Case "CONTA À PARTE"
              fundo_escritorio = fundo_escritorio + valor

            Case Else
              Call MsgBox("O ADVOGADO " & advogado & " NÃO TEM NENHUMA REGRA DE LUCRO CADASTRADA", vbExclamation, "REGRA NÃO CADASTRADA")
          End Select
        Next
      Next

      total_isa = add_no_(total_isa, 0.02, total_bruno) + add_no_(total_isa, 0.02, total_paulo)
      total_bruno = add_no_(total_bruno, -0.02, total_bruno)
      total_paulo = add_no_(total_paulo, -0.02, total_paulo)
    End With

    .Range("E" & ultima_linha("RESULTADO", "E", planilha) + 2) = "BRUNO: " & total_bruno
    .Range("E" & ultima_linha("RESULTADO", "E", planilha)) = "PAULO: " & total_paulo

    With .PivotTables("TOTAL ENTRADA")

      Dim total_saida As Double: total_saida = 0
      Dim n As Integer
      For n = 0 To FrmCtrlSaida.cbx_tipo.ListCount - 1
        Dim tipo_saida As String: tipo_saida = FrmCtrlSaida.cbx_tipo.List(n)
        Dim valor_saida As Double: valor_saida = 0
        valor_saida = .GetPivotData("VALOR", "TIPO", tipo_saida)
                
        Select Case tipo_saida
          Case "CONTA À PARTE"
            fundo_escritorio = fundo_escritorio - valor

          Case Else
            total_saida = total_saida + valor_saida
        End Select
      Next

      total_bruno = add_no_(total_bruno, -0.5, total_saida)
      total_paulo = add_no_(total_paulo, -0.5, total_saida)

    End With

    .Range("E" & ultima_linha("RESULTADO", "E", planilha)) = "ISABELA" & total_isa
    .Range("E" & ultima_linha("RESULTADO", "E", planilha) + 1) = "BRUNO - DESPESAS: " & total_bruno
    .Range("E" & ultima_linha("RESULTADO", "E", planilha) + 1) = "PAULO - DESPESAS: " & total_paulo
  End With

  Call MsgBox("TOTAL CALCULADO COM SUCESSO", vbInformation, "SUCESSO")
  On Error GoTo 0
End Sub