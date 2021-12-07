Option Explicit

Sub get_vendas()
    With Sheets("BASE_VENDAS")            
        Dim item_vendido As Dictionary: Dim linha as Integer
        Dim response As String: Dim venda As Dictionary: Dim json_obj As Dictionary
        Dim request As New WinHttp.WinHttpRequest: Dim objeto_retornado As New Dictionary
        Dim ult_inclusao As Date: ult_inclusao = CDate(WorksheetFunction.Max(.Range("K:K")))
        Dim page As Integer: page = 1: Dim ult_linha As Integer: ult_linha = .Range("A1048576").End(xlUp).Row + 1
        
        For linha = ult_linha - 1 To 6 Step -1
            If ult_inclusao = .Range("K" & linha).Value Then .Range("A" & linha).EntireRow.Delete Else Exit For
            ult_linha = linha
        Next
        Do While True
            With request
                .Open "GET", api_url & "pedidos/page=" & page & "/json/?loja=" & id_loja & _
                    "&filters=dataEmissao[" & ult_inclusao & " TO " & Date & "]&historico=true&apikey=" & api_key, False
                .Send
            End With
            response = request.ResponseText
            If InStr(response, "erros") <> 0 Then Exit Do
            
            Set json_obj = JsonConverter.ParseJson(response)
            For Each objeto_retornado In json_obj("retorno")("pedidos")
                Set venda = objeto_retornado("pedido")
                
                If IsEmpty(venda("itens")) Then GoTo proximo
                For Each item_vendido In venda("itens")
                    
                    .Cells(ult_linha, 1).Value = item_vendido("item")("descricao")
                    .Cells(ult_linha, 2).Value = Trim(item_vendido("item")("codigo"))
                    .Cells(ult_linha, 3).Value = item_vendido("item")("quantidade")
                    
                    'tamanho | loja AVLE
                    If Not IsNumeric(Right(item_vendido("item")("codigo"), 1)) And Right(item_vendido("item")("codigo"), 1) <> "" Then
                        .Cells(ult_linha, 4).Value = Right(item_vendido("item")("codigo"), 1)
                    ElseIf InStr(item_vendido("item")("descricao"), " - ") <> 0 And InStr(item_vendido("item")("descricao"), ",") = 0 Then
                        .Cells(ult_linha, 4).Value = Trim(Right(item_vendido("item")("descricao"), 2))
                    ElseIf InStr(item_vendido("item")("descricao"), ":") <> 0 And InStr(item_vendido("item")("descricao"), ";") = 0 Then
                        .Cells(ult_linha, 4).Value = Trim(Right(item_vendido("item")("descricao"), Len(item_vendido("item")("descricao")) - InStr(item_vendido("item")("descricao"), ":")))
                    End If

                    .Cells(ult_linha, 5).Value = Replace(item_vendido("item")("valorunidade"), ".", ",") * (1 - (Replace(venda("desconto"), ",", ".") / Replace(venda("totalprodutos"), ",", ".")))
                    .Cells(ult_linha, 6).Value = item_vendido("item")("precocusto")
                    .Cells(ult_linha, 7).Value = item_vendido("item")("descontoItem")
                    .Cells(ult_linha, 8).Value = Replace(venda("desconto"), ",", ".")
                    .Cells(ult_linha, 9).Value = Replace(venda("valorfrete"), ",", ".")
                    .Cells(ult_linha, 10).Value = Replace(venda("totalprodutos"), ",", ".")
                    .Cells(ult_linha, 11).Value = Replace(venda("totalvenda"), ",", ".")
                    .Cells(ult_linha, 12).Value = venda("data")
                    .Cells(ult_linha, 13).Value = "'" & Year(venda("data")) & "." & Format(Month(venda("data")), "00")
                    If Not IsEmpty(venda("parcelas")) Then
                        If venda("parcelas")(1)("parcela")("obs") <> "" Then .Cells(ult_linha, 14).Value = Trim(Split(Split(venda("parcelas")(1)("parcela")("obs"), "|")(1), ":")(1))
                        .Cells(ult_linha, 15).Value = venda("parcelas").Count
                    End If
                    .Cells(ult_linha, 16).Value = venda("numero")
                    .Cells(ult_linha, 17).Value = venda("numeroPedidoLoja")
                    .Cells(ult_linha, 18).Value = venda("vendedor")
                    .Cells(ult_linha, 19).Value = venda("situacao")
                    .Cells(ult_linha, 20).Value = venda("loja")
                    .Cells(ult_linha, 21).Value = "SITE"
                    
                    'origem_venda | loja AVLE
                    If venda("loja") = "" Then .Cells(ult_linha, 21).Value = "LOJA BH"
                    
                    .Cells(ult_linha, 22).Value = venda("cliente")("nome")
                    .Cells(ult_linha, 23).Value = venda("cliente")("cnpj")
                    .Cells(ult_linha, 24).Value = venda("cliente")("ie")
                    .Cells(ult_linha, 25).Value = venda("cliente")("rg")
                    .Cells(ult_linha, 26).Value = venda("cliente")("endereco")
                    .Cells(ult_linha, 27).Value = venda("cliente")("numero")
                    .Cells(ult_linha, 28).Value = venda("cliente")("complemento")
                    .Cells(ult_linha, 29).Value = venda("cliente")("cidade")
                    .Cells(ult_linha, 30).Value = venda("cliente")("bairro")
                    .Cells(ult_linha, 31).Value = venda("cliente")("cep")
                    .Cells(ult_linha, 32).Value = venda("cliente")("uf")
                    .Cells(ult_linha, 33).Value = venda("cliente")("email")
                    .Cells(ult_linha, 34).Value = venda("cliente")("celular")
                    .Cells(ult_linha, 35).Value = venda("cliente")("fone")

                    ult_linha = ult_linha + 1
                Next
proximo:
            Next
            page = page + 1
        Loop
        .Columns("A:AF").ColumnWidth = 25
        .Columns("D:J").Style = "Currency"
        Call format_header(.Name)
        
        .Range("A1").Select
    End With
    Call MsgBox("agora todas as vendas cadastradas no bling estão aqui! :D", vbInformation, "Base Atualizada")
End Sub