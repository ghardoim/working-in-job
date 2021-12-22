Option Explicit

Sub get_produtos()
    Call liga_desliga(False)
    With Sheets("BASE_PRODUTOS")
        Dim response As String: Dim produto As Dictionary: Dim json_obj As Dictionary
        Dim request As New WinHttp.WinHttpRequest: Dim objeto_retornado As New Dictionary
        Dim ult_inclusao As Date: ult_inclusao = CDate(WorksheetFunction.Max(.Range("Q:Q"))) + 1
        Dim page As Integer: page = 1: Dim ult_linha As Integer: ult_linha = .Range("A1048576").End(xlUp).Row + 1
        Dim lixo_html As Variant: Dim tags_html() As Variant: tags_html = Array("<strong>", "</strong>", "&#8211;", "&nbsp;", "<p>", "</p>", "<br />", "&lt;3")

        Do While True
            With request
                .Open "GET", api_url & "produtos/page=" & page & "/json/?loja=" & id_loja & _
                    "&filters=dataInclusao[" & ult_inclusao & " TO " & Date & "]&imagem=S&estoque=S&apikey=" & api_key, False
                .Send
            End With
            response = request.ResponseText
            If InStr(response, "erros") <> 0 Then Exit Do

            Set json_obj = JsonConverter.ParseJson(response)
            For Each objeto_retornado In json_obj("retorno")("produtos")
                Set produto = objeto_retornado("produto")
                For Each deposito In produto("depositos")

                    .Cells(ult_linha, 1).Value = RTrim(produto("descricao"))
                    If IsNull(produto("descricaoCurta")) Then produto("descricaoCurta") = ""
                    For Each lixo_html In tags_html
                        produto("descricaoCurta") = Replace(produto("descricaoCurta"), lixo_html, "")
                    Next
                    .Cells(ult_linha, 2).Value = RTrim(Replace(Replace(Replace(Replace(produto("descricaoCurta"), vbCrLf, " "), " :", ":"), " !", "!"), " ,", ","))
                    .Cells(ult_linha, 3).Value = produto("codigoPai")

                    If InStr(LCase(produto("descricao")), "cor:") <> 0 Then .Cells(ult_linha, 4).Value = get_color(produto("descricao"), "cor:", 3)
                    If InStr(LCase(produto("descricao")), "color:") <> 0 Then .Cells(ult_linha, 4).Value = get_color(produto("descricao"), "color:", 5)

                    'set_tamanho | loja FELINE
                    If achei("tam", produto("descricao")) And Not achei("estampa", produto("descricao")) And Not achei("tamanho", produto("descricao")) And Not achei("tamiris", produto("descricao")) And Not achei("tamires", produto("descricao")) Then
                        .Cells(ult_linha, 5).Value = get_tamanho(produto("descricao"), "TAM", 3)
                    ElseIf achei("size", produto("descricao")) Then
                        .Cells(ult_linha, 5).Value = get_tamanho(produto("descricao"), "SIZE", 4)
                    ElseIf achei("tamanho", produto("descricao")) Then
                        .Cells(ult_linha, 5).Value = get_tamanho(produto("descricao"), "TAMANHO", 7)
                    End If
                    If achei("color", .Cells(ult_linha, 5).Value) Or achei("cor", .Cells(ult_linha, 5).Value) Then
                        .Cells(ult_linha, 5).Value = Left(.Cells(ult_linha, 5).Value, InStr(LCase(.Cells(ult_linha, 5).Value), "co") - 2)
                    End If

                    'set_tamanho | loja AVLE
                    If Not IsNumeric(Right(produto("codigo"), 1)) Then .Cells(ult_linha, 5).Value = Right(produto("codigo"), 1)

                    .Cells(ult_linha, 6).Value = produto("codigo")
                    .Cells(ult_linha, 7).Value = produto("estoqueAtual")
                    .Cells(ult_linha, 8).Value = produto("preco")
                    .Cells(ult_linha, 9).Value = Replace(produto("preco"), ".", ",") * produto("estoqueAtual")

                    If IsNull(produto("precoCusto")) Then produto("precoCusto") = 0
                    'preço_custo | loja FELINE
                    .Cells(ult_linha, 10).Value = Replace(produto("precoCusto"), ".", ",") / 2.5
                    .Cells(ult_linha, 11).Value = .Cells(ult_linha, 10).Value * produto("estoqueAtual")

                    'preço_custo | loja AVLE
                    .Cells(ult_linha, 10).Value = produto("precoCusto")
                    .Cells(ult_linha, 11).Value = Replace(produto("precoCusto"), ".", ",") * produto("estoqueAtual")

                    If Not IsEmpty(produto("produtoLoja")) Then .Cells(ult_linha, 12).Value = produto("produtoLoja")("preco")("preco")
                    If Not IsEmpty(produto("produtoLoja")) Then .Cells(ult_linha, 13).Value = produto("produtoLoja")("preco")("precoPromocional")

                    .Cells(ult_linha, 14).Value = produto("idGrupoProduto")
                    .Cells(ult_linha, 15).Value = produto("grupoProduto")

                    If 0 < produto("imagem").Count Then .Cells(ult_linha, 16).Value = produto("imagem")(1)("link")
                    .Cells(ult_linha, 17).Value = produto("dataInclusao")
                    .Cells(ult_linha, 18).Value = deposito("deposito")("nome")
                    .Cells(ult_linha, 19).Value = deposito("deposito")("saldo")
                    
                    ult_linha = ult_linha + 1
                Next
            Next
            page = page + 1
        Loop
        .Columns("A:S").ColumnWidth = 25
        .Columns("H:M").Style = "Currency"
        Call format_header(.Name)

        .Range("A1").Select
    End With
    Call MsgBox("agora todos os produtos cadastrados no bling estão aqui! :D", vbInformation, "Base Atualizada")
    Call liga_desliga(True)
End Sub

Sub drop_produtos()
    Sheets("BASE_PRODUTOS").Range("A6:S1048576").Delete
End Sub