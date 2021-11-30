Public Const id_loja As String = ""
Public Const api_key As String = ""
Public Const api_url As String = "https://bling.com.br/Api/v2/"
        
Public Sub format_header(nome_planilha As String)
    With Sheets(nome_planilha)
        .Range("A5").CurrentRegion.AutoFilter
        .Range("A5").CurrentRegion.HorizontalAlignment = xlJustify
        .Range("A5").CurrentRegion.RowHeight = 15
        With .Range(Range("A5"), .Range("A5").End(xlToRight))
            .Interior.Pattern = xlSolid
            .Interior.Color = RGB(173, 216, 230)
            .Borders(xlDiagonalDown).LineStyle = xlNone
            .Borders(xlDiagonalUp).LineStyle = xlNone
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlInsideVertical).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
        End With
    End With
    Call handle_description(nome_planilha)
End Sub

Public Sub handle_description(sheet_name As String)
    Dim position As Integer: Dim celula As Range: Dim remover As Variant
    With Sheets(sheet_name)
        For Each celula In .Range("A6:A" & .Range("A1048576").End(xlUp).Row)
            For Each remover In Array("cor:", "tam:", "size:", "color:", "tamanho:")
                position = InStr(LCase(celula.Value), remover)
                If position <> 0 Then celula.Value = Trim(Left(celula.Value, position - 1))
            Next
            If "-" = Right(celula.Value, 1) Then celula.Value = Trim(Left(celula.Value, Len(celula.Value) - 1))
        Next
    End With
End Sub

Public Function achei(isso As String, naquilo As String) As Boolean
    achei = InStr(LCase(naquilo), isso) <> 0
End Function

Public Function get_tamanho(descricao As String, str_procura As String, str_len As Integer) As String
    On Error Resume Next
    get_tamanho = Trim(Right(descricao, Len(descricao) - InStr(UCase(descricao), str_procura) - str_len))
    On Error GoTo 0
End Function

Public Function all_unique(col_letter As String, sheet_name As String) As Variant
    Dim all_uniques(): all_uniques = Sheets(sheet_name).Range(col_letter & "6:" & col_letter & Sheets(sheet_name).Range("A1048576").End(xlUp).Row).Value
    Dim dict_uniques As New Scripting.Dictionary, linha As Integer
    For linha = 1 To UBound(all_uniques)
        dict_uniques(all_uniques(linha, 1)) = Empty
    Next
    all_unique = dict_uniques.Keys
End Function