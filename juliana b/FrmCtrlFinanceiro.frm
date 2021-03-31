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
  ThisWorkbook.Save
End Sub

Private Sub btn_total_Click()
  On Error Resume Next
  If planilha Is Nothing Then Exit Sub

  If planilha.PivotCaches.Count > 0 Then planilha.Sheets("RESULTADO").UsedRange.ClearContents
  Call cria_tabela("ENTRADA", 1, "VALOR LÍQUIDO", "B", "ADVOGADO", 2, "IMPOSTO", "C")
  Call cria_tabela("SAÍDA", 5, "VALOR", "F")

  Dim fundo_escritorio As Double: fundo_escritorio = Me.lbl_total
  Dim total_bruno As Double: total_bruno = 0
  Dim total_paulo As Double: total_bruno = 0
  Dim total_isa As Double: total_isa = 0

  With planilha.Sheets("RESULTADO")
    With .PivotTables("TOTAL ENTRADA")

      Dim tipo_entrada As PivotItem
      For Each tipo_entrada In .PivotFields("TIPO").PivotItems()

        Dim advogado As PivotItem
        For Each advogado In .PivotFields("ADVOGADO").PivotItems()

          Dim valor_entrada As Double: valor_entrada = 0
          valor_entrada = .GetPivotData("VALOR LÍQUIDO", "TIPO", tipo_entrada.Caption, "ADVOGADO", advogado.Caption)
          If 0 = valor_entrada Then GoTo Continue

          If "CONTA À PARTE" = tipo_entrada.Caption Or "REEMBOLSO" = tipo_entrada.Caption Then
              'LUCROS CONTA À PARTE -> CRÉDITO DIRETO NO FUNDO DO ESCRITÓRIO
              fundo_escritorio = fundo_escritorio + valor_entrada
              GoTo Continue

          ElseIf "MENSALIDADE" = tipo_entrada.Caption Or "HONORÁRIOS" = tipo_entrada.Caption Then
            valor_entrada = add_no_(valor_entrada, -0.05, valor_entrada)
            fundo_escritorio = add_no_(fundo_escritorio, 0.05, valor_entrada)
          End If

          Select Case advogado.Caption
            Case "BRUNO"
              total_bruno = add_no_(total_bruno, 0.6, valor_entrada)
              total_paulo = add_no_(total_paulo, 0.4, valor_entrada)

            Case "PAULO"
              total_paulo = add_no_(total_paulo, 0.6, valor_entrada)
              total_bruno = add_no_(total_bruno, 0.4, valor_entrada)
          
            Case "ISABELA"
              total_isa = add_no_(total_isa, 0.5, valor_entrada)
              total_paulo = add_no_(total_paulo, 0.25, valor_entrada)
              total_bruno = add_no_(total_bruno, 0.25, valor_entrada)

            Case "BRUNO & PAULO"
              total_paulo = add_no_(total_paulo, 0.5, valor_entrada)
              total_bruno = add_no_(total_bruno, 0.5, valor_entrada)

            Case Else
              Call MsgBox("O ADVOGADO " & advogado.Caption & " NÃO TEM NENHUMA REGRA DE LUCRO CADASTRADA", vbExclamation, "REGRA NÃO CADASTRADA")
          End Select
Continue:
        Next
      Next
      total_isa = add_no_(total_isa, 0.02, total_bruno) + add_no_(total_isa, 0.02, total_paulo)
      total_bruno = add_no_(total_bruno, -0.02, total_bruno)
      total_paulo = add_no_(total_paulo, -0.02, total_paulo)
    End With

    .Range("E" & ultima_linha("RESULTADO", "E", planilha) + 2) = "BRUNO: " & total_bruno
    .Range("E" & ultima_linha("RESULTADO", "E", planilha)) = "PAULO: " & total_paulo

    With .PivotTables("TOTAL SAÍDA")

      Dim total_saida As Double: total_saida = 0

      Dim tipo_saida As PivotItem
      For Each tipo_saida In .PivotFields("TIPO").PivotItems()

        Dim valor_saida As Double: valor_saida = 0
        valor_saida = .GetPivotData("VALOR", "TIPO", tipo_saida)

        If "CONTA À PARTE" = tipo_saida Or "REEMBOLSO" = tipo_saida Then
          fundo_escritorio = fundo_escritorio - valor_saida
        Else
          total_saida = total_saida + valor_saida
        End If
      Next

      total_bruno = add_no_(total_bruno, -0.5, total_saida)
      total_paulo = add_no_(total_paulo, -0.5, total_saida)
    End With

    .Range("E" & ultima_linha("RESULTADO", "E", planilha)) = "ISABELA: " & total_isa
    .Range("E" & ultima_linha("RESULTADO", "E", planilha) + 1) = "BRUNO - DESPESAS: " & total_bruno
    .Range("E" & ultima_linha("RESULTADO", "E", planilha) + 1) = "PAULO - DESPESAS: " & total_paulo
  End With

  ThisWorkbook.Sheets("AUX").Range("E2") = fundo_escritorio
  Call MsgBox("TOTAL CALCULADO COM SUCESSO", vbInformation, "SUCESSO")
  On Error GoTo 0
End Sub