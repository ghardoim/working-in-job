Sub dezenas_fora_da_combinacao()

    With ThisWorkbook.Sheets("PLAN-COMBINAÇOES")
        matriz_combinacao = .Range(.Cells(5, 4), .Cells(5, .Cells(5, 4).End(xlToRight).Column))

        For linha = 14 To .Range("D1048576").End(xlUp).Row
            jogos = .Range(.Cells(linha, 4), .Cells(linha, .Cells(linha, 4).End(xlToRight).Column))

            With ThisWorkbook.Sheets("PLAN-DEZENAS FORA")
                ultima_linha = .Range("D1048576").End(xlUp).Row + 1
                For Each jogo In jogos
                
                    For Each numero In matriz_combinacao
                        If jogo = numero Then GoTo next_jogo
                    Next

                    ultima_coluna = .Cells(ultima_linha, 15000).End(xlToLeft).Column + 1
                    .Cells(ultima_linha, ultima_coluna).Value = jogo
next_jogo:
                Next
            End With
        Next
    End With
    ThisWorkbook.Sheets("PLAN-DEZENAS FORA").Select
End Sub