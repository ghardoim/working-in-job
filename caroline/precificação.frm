Private Function option_selected_in(frame As MSForms.frame) As String
    For Each ctrl In frame.Controls
        If CBool(ctrl.value) Then option_selected_in = CStr(ctrl.Caption)
    Next
End Function

Private Function convert(label As MSForms.label) As Double
    If "" <> label And "-" <> label Then convert = CDbl(Split(label, " ")(1))
End Function

Private Sub atualiza_all()
    Call atualiza_qnt(Me.txt_codigo.value)
    Call atualiza_frete(Me.cmb_cidade.value, option_selected_in(Me.frm_meio_transporte))
    Call atualiza_valores_finais
End Sub

Private Sub atualiza_valores_finais()
    preco_net = CDbl(IIf("" <> Me.txt_preco_net.value, Me.txt_preco_net.value, 0))
    Me.lbl_valor_added_value.Caption = FormatCurrency(preco_net _
            - convert(Me.lbl_valor_material) - convert(Me.lbl_valor_energia) - convert(Me.lbl_valor_frete))
    On Error Resume Next
        Me.lbl_valor_av.Caption = FormatPercent(convert(Me.lbl_valor_added_value) / preco_net, 2)
    On Error GoTo 0
End Sub

Private Sub atualiza_qnt(codigo As String)
    codigo = CLng(Me.txt_codigo.value)
    On Error Resume Next
        coluna_qnt = IIf("BAÚ" = option_selected_in(Me.frm_meio_transporte), "G:G", "H:H")
        With WorksheetFunction
            chave_busca_qnt = .Index(Sheets("BASE").Range("H:H"), .Match(codigo, Sheets("BASE").Range("F:F"), 0))
            quantidade = .Index(Sheets("QUANT").Range(coluna_qnt), .Match(chave_busca_qnt, Sheets("QUANT").Range("F:F"), 0))
        End With
        Me.lbl_valor_qnt.Caption = quantidade
    On Error GoTo 0
End Sub

Private Sub atualiza_frete(cidade As String, tipo_caminhao As String)
    If "FOB" = tipo_caminhao Then
        Me.lbl_valor_frete.Caption = "-"
    Else
        nome_aba = "Tabela Frete At. 23.03.22"
        maximo = WorksheetFunction.MaxIfs(Sheets(nome_aba).Range("O:O"), _
                                        Sheets(nome_aba).Range("F:F"), cidade, _
                                        Sheets(nome_aba).Range("B:B"), tipo_caminhao)
        Me.lbl_valor_frete.Caption = FormatCurrency(maximo, 2)
    End If
End Sub

Private Sub cmb_cidade_AfterUpdate()
    Call atualiza_all
End Sub

Private Sub ALLOptionButtons_Click()
    Call atualiza_all
End Sub

Private Sub txt_preco_net_AfterUpdate()
    Call atualiza_all
End Sub

Private Sub txt_codigo_AfterUpdate()
    On Error Resume Next
        codigo = CLng(Me.txt_codigo.value)
        With WorksheetFunction
            site = .Index(Sheets("BASE").Range("C:C"), .Match(codigo, Sheets("BASE").Range("F:F"), 0))
            Me.lbl_descricao.Caption = .Index(Sheets("BASE").Range("G:G"), .Match(codigo, Sheets("BASE").Range("F:F"), 0))

            variacao = .Index(Sheets("BASE").Range("U:U"), .Match(codigo, Sheets("BASE").Range("F:F"), 0))
            If 0 = variacao Then
                Me.lbl_valor_material.BorderStyle = fmBorderStyleSingle
                Me.lbl_valor_material.BorderColor = RGB(250, 0, 0)
            Else
                valor_material = .Index(Sheets("BASE").Range("N:N"), .Match(codigo, Sheets("BASE").Range("F:F"), 0))
                Me.lbl_valor_material.Caption = FormatCurrency(valor_material, 2)
            End If

            valor_energia = .Index(Sheets("BASE").Range("O:O"), .Match(codigo, Sheets("BASE").Range("F:F"), 0))
            Me.lbl_valor_energia.Caption = IIf(valor_energia <> 0, FormatCurrency(valor_energia, 2), "-")

            Call atualiza_all
        End With
    On Error GoTo 0
End Sub