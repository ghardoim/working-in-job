Sub get_clientes()
    Call liga_desliga(False)
    With Sheets("BASE_CLIENTES")
        .Rows("6:1048576").Delete

        Dim response As String: Dim cliente As Dictionary: Dim json_obj As Dictionary
        Dim request As New WinHttp.WinHttpRequest: Dim objeto_retornado As New Dictionary
        Dim ult_inclusao As Date: ult_inclusao = CDate(WorksheetFunction.Max(.Range("J:J"))) + 1
        Dim page As Integer: page = 1: Dim ult_linha As Integer: ult_linha = .Range("A1048576").End(xlUp).Row + 1

        anos_meses = all_unique("Q", "BASE_VENDAS")
        .Cells(5, UBound(anos_meses) + 15).Value = "Ticket Médio"
        .Cells(5, UBound(anos_meses) + 16).Value = "Classificação"

        Do While True
            With request
                .Open "GET", api_url & "contatos/page=" & page & "/json/?filters=" & _
                    "dataInclusao[" & ult_inclusao & " TO " & Date & "]&apikey=" & api_key, False
                .Send
            End With
            response = request.ResponseText
            If InStr(response, "erros") <> 0 Then Exit Do

            Set json_obj = JsonConverter.ParseJson(response)
            For Each objeto_retornado In json_obj("retorno")("contatos")
                Set cliente = objeto_retornado("contato")

                nome = Trim(StrConv(cliente("nome"), vbProperCase))
                .Cells(ult_linha, 1).Value = cliente("id")
                .Cells(ult_linha, 2).Value = nome
                .Cells(ult_linha, 3).Value = cliente("tipo")
                .Cells(ult_linha, 4).Value = cliente("cnpj")
                .Cells(ult_linha, 5).Value = cliente("bairro")
                .Cells(ult_linha, 6).Value = cliente("cidade")
                .Cells(ult_linha, 7).Value = cliente("uf")
                .Cells(ult_linha, 8).Value = IIf(IsEmpty(cliente("celular")), cliente("fone"), cliente("celular"))
                .Cells(ult_linha, 9).Value = cliente("email")
                .Cells(ult_linha, 10).Value = cliente("clienteDesde")

                ultima_compra = WorksheetFunction.MaxIfs(Sheets("BASE_VENDAS").Range("P:P"), Sheets("BASE_VENDAS").Range("Z:Z"), nome)
                valor_total = WorksheetFunction.SumIfs(Sheets("BASE_VENDAS").Range("F:F"), Sheets("BASE_VENDAS").Range("Z:Z"), nome)
                .Cells(ult_linha, 11).Value = IIf(ultima_compra <> 0, ultima_compra, "")
                .Cells(ult_linha, 12).Value = IIf(ultima_compra <> 0, Date - ultima_compra, "")
                .Cells(ult_linha, 13).Value = valor_total

                For am = 0 To UBound(anos_meses)
                    .Cells(5, am + 14).Value = "'" & anos_meses(am)
                    .Cells(ult_linha, am + 14).Value = WorksheetFunction.SumIfs(Sheets("BASE_VENDAS").Range("F:F"), Sheets("BASE_VENDAS").Range("Q:Q"), anos_meses(am), Sheets("BASE_VENDAS").Range("Z:Z"), nome)
                Next
                Set range_anomes = .Range(.Cells(ult_linha, 14), .Cells(ult_linha, UBound(anos_meses) + 14))
                n_meses_venda = WorksheetFunction.CountIf(range_anomes, ">0")
                If n_meses_venda <> 0 Then ticket_medio = WorksheetFunction.Sum(range_anomes) / n_meses_venda

                .Cells(ult_linha, UBound(anos_meses) + 15).Value = ticket_medio
                If ticket_medio > .Range("B3").Value Then .Cells(ult_linha, UBound(anos_meses) + 16).Value = "VIP"
                If ticket_medio > .Range("C3").Value Then .Cells(ult_linha, UBound(anos_meses) + 16).Value = "Potencial"
                If WorksheetFunction.CountIf(.Range(.Cells(ult_linha, 13), .Cells(ult_linha, UBound(anos_meses) + 15)), "<>0") _
                    >= 6 Then .Cells(ult_linha, UBound(anos_meses) + 16).Value = "Recorrente"
                ult_linha = ult_linha + 1
            Next
            page = page + 1
        Loop
        .Range("M:" & Split(Cells(1, UBound(anos_meses) + 15).Address(True, False), "$")(0)).Style = "Currency"
        .Range("A1").Select
    End With
    Call MsgBox("agora todos os clientes cadastrados no bling estão aqui! :D", vbInformation, "Base Atualizada")
    Call liga_desliga(True)
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Address = Range("B3").Address Or Target.Address = Range("C3").Address Then Call get_clientes
End Sub