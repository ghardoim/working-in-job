Sub dezenas_fora_da_combinacao()
    ThisWorkbook.Sheets("PLAN-DEZENAS FORA").Range("D12:DD1500").Delete

    With ThisWorkbook.Sheets("PLAN-COMBINAÇOES")
        matriz_combinacao = Application.Transpose(Application.Transpose( _
                                .Range(.Cells(5, 4), .Cells(5, .Cells(5, 4).End(xlToRight).Column))))

        For linha = 14 To .Range("D1048576").End(xlUp).Row
            jogo = Application.Transpose(Application.Transpose( _
                    .Range(.Cells(linha, 4), .Cells(linha, .Cells(linha, 4).End(xlToRight).Column))))

            With ThisWorkbook.Sheets("PLAN-DEZENAS FORA")
                ultima_linha = .Range("D1048576").End(xlUp).Row + 1
                For Each dezena In matriz_combinacao

                    If Not AcheiEsseNumero(dezena, jogo) Then
                        ultima_coluna = .Cells(ultima_linha, 15000).End(xlToLeft).Column + 1
                        .Cells(ultima_linha, ultima_coluna).Value = dezena
                    End If
                Next
        Next
    End With

    ThisWorkbook.Sheets("PLAN-DEZENAS FORA").Select
End Sub

Private Function AcheiEsseNumero(ByVal item As Integer, lista As Variant) As Boolean
    For Each n In lista: If item = n Then AcheiEsseNumero = True: Exit For
    Next
End Function