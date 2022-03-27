Sub get_extrato()
    Call liga_desliga(False)
    Call drop_extrato
    With Sheets("BASE_PAGAR.ME")

        Dim operation As New Dictionary: Dim coluna As Integer
        Dim request As New WinHttp.WinHttpRequest: Dim response As String: Dim json_obj As Dictionary
        Dim page As Integer: page = 1: Dim ult_linha As Integer: ult_linha = .Range("A1048576").End(xlUp).Row + 1
        Do While True
            With request
                .Open "GET", pagar_me_api_url & "/balance/operations?count=1000&page=" & page & "&api_key=" & pagar_me_api_key
                .Send
            End With
            response = request.ResponseText
            If "[]" = response Then Exit Do
            Set json_obj = JsonConverter.ParseJson("{ ""response"": " & response & "}")

            For Each operation In json_obj("response")
                coluna = 1
                For Each operation_key In operation.Keys
                    If "movement_object" <> operation_key Then
                        .Cells(ult_linha, coluna).Value = operation(operation_key)
                        coluna = coluna + 1
                    Else
                        For Each movement_key In operation(operation_key).Keys
                            If "bank_account" <> movement_key And "metadata" <> movement_key And "movement_object" <> movement_key Then
                                .Cells(ult_linha, coluna).Value = operation(operation_key)(movement_key)
                                coluna = coluna + 1
                            End If
                        Next
                    End If
                Next
                ult_linha = ult_linha + 1
            Next
            page = page + 1
        Loop
    End With
    Call MsgBox("agora todas as operações do Pagar.me estão aqui! :D", vbInformation, "Extrato Atualizado")
    Call liga_desliga(True)
End Sub

Sub drop_extrato()
    Sheets("BASE_PAGAR.ME").Rows("6:1048576").Delete
End Sub