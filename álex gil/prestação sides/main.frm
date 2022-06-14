Private Sub campo_vazio(campo As String)
    Call MsgBox("Por favor, preencha o campo " & campo, vbExclamation, "Campo " & campo & " � obrigat�rio!")
End Sub

Private Sub btn_adicionar_Click()
    With Sheets("Receitas")
        Dim ult_linha As Integer: ult_linha = .Range("A1048576").End(xlUp).Row + 1

        If "" = Me.txt_beneficiario.Value Then Call campo_vazio("BENEFICIÁRIO"): Exit Sub
        If "" = Me.cmb_dia.Value Or "" = Me.cmb_mes.Value Or "" = Me.cmb_ano.Value Then Call campo_vazio("DATA"): Exit Sub
        If "" = Me.txt_despesa.Value Then Call campo_vazio("DESPESA"): Exit Sub
        If "" = Me.cmb_tipo.Value Then Call campo_vazio("TIPO DE DESPESA"): Exit Sub

        If "Recebido" <> Me.cmb_tipo.Value Then
            If "" = Me.txt_pagamento.Value Then Call campo_vazio("N° PAGAMENTO"): Exit Sub
            If "" = Me.txt_notafiscal.Value Then Call campo_vazio("NOTA FISCAL"): Exit Sub
        End If

        .Cells(ult_linha, 1).Value = "01"
        If 2 <> ult_linha Then .Cells(ult_linha, 1).Value = .Cells(ult_linha - 1, 1).Value + 1

        On Error GoTo informe_valor
        .Cells(ult_linha, 3).Value = Abs(CDbl(Me.txt_despesa.Value))
        If "Recebido" <> Me.cmb_tipo.Value Then
            .Cells(ult_linha, 3).Value = .Cells(ult_linha, 3).Value * -1
            With .Cells(ult_linha, 3).Interior
                .Color = 255
                .TintAndShade = 0.6
            End With
        End If
        .Cells(ult_linha, 2).Value = CDate(Me.cmb_dia.Value & "/" & Me.cmb_mes.Value & "/" & Me.cmb_ano.Value)
        .Cells(ult_linha, 4).Value = Me.txt_beneficiario.Value
        .Cells(ult_linha, 5).Value = IIf("" <> Me.txt_notafiscal.Value, Me.txt_notafiscal.Value, "-")
        .Cells(ult_linha, 6).Value = IIf("" <> Me.txt_pagamento.Value, Me.txt_pagamento.Value, "-")
        .Cells(ult_linha, 7).Value = Me.cmb_tipo.Value
        .Cells(ult_linha, 8).Value = IIf("" <> Me.txt_observacoes.Value, Me.txt_observacoes.Value, "-")
        .Range(.Cells(ult_linha, 1), .Cells(ult_linha, 8)).Borders().LineStyle = xlContinuous
    End With
    Me.txt_beneficiario.Value = ""
    Me.txt_observacoes.Value = ""
    Me.txt_notafiscal.Value = ""
    Me.txt_pagamento.Value = ""
    Me.txt_despesa.Value = ""
    Me.cmb_tipo.Value = ""
    Me.cmb_dia.Value = ""
    Me.cmb_mes.Value = ""
    Me.lst_despesas.RowSource = "Receitas!A2:H" & ult_linha

informe_valor:
    If 13 = Err.Number Then Call MsgBox("Por favor, informe um n�mero no campo DESPESA!", vbExclamation, "O campo DESPESA � num�rico!")
    On Error GoTo 0
End Sub

Private Sub UserForm_Activate()
    Me.lst_despesas.RowSource = "Receitas!A2:H" & Sheets("Receitas").Range("A1048576").End(xlUp).Row + 1
End Sub

Private Sub UserForm_Terminate()
    Call atualizar_resumo
End Sub