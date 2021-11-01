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
End Sub

Public Function achei(isso As String, naquilo As String) As Boolean
    achei = InStr(LCase(naquilo), isso) <> 0
End Function