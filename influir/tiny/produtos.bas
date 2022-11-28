Public Const api_url As String = "https://api.tiny.com.br/api2/"
Public Const api_key As String = ""

Public Sub liga_desliga(on_off As Boolean)
  With Application
    If on_off Then .Calculation = xlCalculationAutomatic
    If Not on_off Then .Calculation = xlCalculationManual
    .ScreenUpdating = on_off
    .DisplayAlerts = on_off
    .EnableEvents = on_off
  End With
End Sub

Sub get_produtos()
    Call liga_desliga(False)
    Dim request As New WinHttp.WinHttpRequest
    Dim response As String
    
    Dim endpoint As String: endpoint = "produtos.pesquisa.php"
    Dim pagina As Integer: pagina = 1
    
    Dim json_obj As Dictionary
    Dim produto As Dictionary
    
    Dim coluna, linha As Integer

    linha = 6
    Do While True
        If 0 = pagina Mod 10 Then Application.Wait (Now + TimeValue("0:00:30"))

        With request
            .Open "POST", api_url & endpoint & "?token=" & api_key & "&formato=JSON&pagina=" & pagina
            .Send
        End With
    
        response = request.ResponseText
        
        Set json_obj = JsonConverter.ParseJson(response)("retorno")
        If json_obj("numero_paginas") < pagina Then Exit Do
        
        For Each produto In json_obj("produtos")
            Set produto = produto("produto")
            coluna = 1
            
            For Each campo In produto
                Sheets("BASE_PRODUTOS").Cells(linha, coluna) = produto(campo)
                coluna = coluna + 1
            Next
            linha = linha + 1
        Next
        pagina = pagina + 1
    Loop
    Call liga_desliga(True)
    Call MsgBox("agora todos os produtos cadastrados no tiny estão aqui! :D", vbInformation, "Base Atualizada")
End Sub