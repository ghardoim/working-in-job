Option Explicit

Sub get_produtos()
    Dim page As Integer: page = 1: Dim ult_linha As Integer: ult_linha = 6
    Dim response As String: Dim produto As Dictionary: Dim json_obj As Dictionary
    Dim request As New WinHttp.WinHttpRequest: Dim objeto_retornado As New Dictionary
    Dim lixo_html As Variant: Dim tags_html() As Variant: tags_html = Array("<strong>", "</strong>", "&#8211;", "&nbsp;", "<p>", "</p>", "<br />", "&lt;3")
        
    With Sheets("BASE_PRODUTOS")
        Do While True
            With request
                .Open "GET", api_url & "produtos/page=" & page & "/json/?loja=" & id_loja & "&imagem=S&estoque=S&apikey=" & api_key, False
                .Send
            End With
            response = request.ResponseText
            If InStr(response, "erros") <> 0 Then Exit Do
            
            Set json_obj = JsonConverter.ParseJson(response)
            For Each objeto_retornado In json_obj("retorno")("produtos")
                Set produto = objeto_retornado("produto")
                
                .Cells(ult_linha, 1).Value = RTrim(produto("descricao"))
                If IsNull(produto("descricaoCurta")) Then produto("descricaoCurta") = ""
                For Each lixo_html In tags_html
                    produto("descricaoCurta") = Replace(produto("descricaoCurta"), lixo_html, "")
                Next
                .Cells(ult_linha, 2).Value = RTrim(Replace(Replace(Replace(Replace(produto("descricaoCurta"), vbCrLf, " "), " :", ":"), " !", "!"), " ,", ","))
                .Cells(ult_linha, 3).Value = produto("codigoPai")
                
                'set_tamanho | loja FELINE
                If achei("tam", produto("descricao")) And Not achei("estampa", produto("descricao")) And Not achei("tamanho", produto("descricao")) And Not achei("tamiris", produto("descricao")) Then .Cells(ult_linha, 4).Value = Trim(Right(produto("descricao"), Len(produto("descricao")) - InStr(UCase(produto("descricao")), "TAM") - 3))
                
                'set_tamanho | loja AVLE
                If Not IsNumeric(Right(produto("codigo"), 1)) Then .Cells(ult_linha, 4).Value = Right(produto("codigo"), 1)
                
                .Cells(ult_linha, 5).Value = produto("codigo")
                .Cells(ult_linha, 6).Value = produto("estoqueAtual")
                
                .Cells(ult_linha, 7).Value = produto("preco")
                .Cells(ult_linha, 8).Value = Replace(produto("preco"), ".", ",") * produto("estoqueAtual")
                
                If IsNull(produto("precoCusto")) Then produto("precoCusto") = 0
                .Cells(ult_linha, 9).Value = produto("precoCusto")
                .Cells(ult_linha, 10).Value = Replace(produto("precoCusto"), ".", ",") * produto("estoqueAtual")
                
                If Not IsEmpty(produto("produtoLoja")) Then .Cells(ult_linha, 11).Value = produto("produtoLoja")("preco")("preco")
                If Not IsEmpty(produto("produtoLoja")) Then .Cells(ult_linha, 12).Value = produto("produtoLoja")("preco")("precoPromocional")
                
                .Cells(ult_linha, 13).Value = produto("idGrupoProduto")
                .Cells(ult_linha, 14).Value = produto("grupoProduto")
                
                If 0 < produto("imagem").Count Then .Cells(ult_linha, 15).Value = produto("imagem")(1)("link")
                
                ult_linha = ult_linha + 1
            Next
            page = page + 1
        Loop
        
        .Columns("A:O").ColumnWidth = 25
        .Columns("G:L").Style = "Currency"
        Call format_header(.Name)
        
        .Range("A1").Select
    End With
    Call MsgBox("agora todos os produtos cadastrados no bling est�o aqui! :D", vbInformation, "Base Atualizada")
End Sub