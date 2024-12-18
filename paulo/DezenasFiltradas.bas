Sub dezenas_filtradas()
    With ThisWorkbook.Sheets("Combinaçoes filtradas")
        .Range("D10:DD1500").ClearContents
        For Each comentario In .Comments
            comentario.Delete
        Next
    End With

    With ThisWorkbook.Sheets("Combinaçoes para filtrar")
        For linha = 10 To .Range("D1048576").End(xlUp).Row
            combinacao = Application.Transpose(Application.Transpose(.Range(.Cells(linha, 4), .Cells(linha, .Range("D" & linha).End(xlToRight).Column))))
            Set combinacao = remove_duplicadas(combinacao)
            For coluna = 1 To combinacao.Count
                Sheets("Combinaçoes filtradas").Cells(linha, coluna + 3).Value = combinacao(coluna)
            Next

            With Sheets("Combinaçoes filtradas").Cells(linha, coluna + 2)
                .AddComment
                .Comment.Visible = False
                .Comment.Text "Total de dezenas: " & combinacao.Count
                .Comment.Shape.Height = 12
                .Comment.Shape.Width = 87
            End With
            Sheets("Combinaçoes filtradas").Cells(linha, 104).Value = combinacao.Count
        Next
    End With
    ThisWorkbook.Sheets("Combinaçoes filtradas").Select
End Sub

Function remove_duplicadas(combinacao As Variant) As Collection
    Dim Coll As New Collection

    On Error Resume Next
    For Each dezena In combinacao
        Coll.Add dezena, CStr(dezena)
    Next dezena
    On Error GoTo 0

    Set remove_duplicadas = Coll
End Function