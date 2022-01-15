Public Const id_loja As String = ""
Public Const api_key As String = ""
Public Const api_url As String = "https://bling.com.br/Api/v2/"
        
Public Sub format_header(nome_planilha As String)
    With Sheets(nome_planilha)
        .Rows(5).AutoFilter
        .Range("A5").CurrentRegion.HorizontalAlignment = xlJustify
        .Range("A5").CurrentRegion.RowHeight = 15
        With .Range(Range("A5"), .Range("A5").End(xlToRight))
            .Interior.Pattern = xlSolid
            .Interior.Color = RGB(173, 216, 230)
            .Borders.LineStyle = xlContinuous
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
            .Cells(celula.Row, 41).Value = Trim(.Cells(celula.Row, 1).Value & " " & .Cells(celula.Row, 4).Value)
            If "-" = Right(celula.Value, 1) Then celula.Value = Trim(Left(celula.Value, Len(celula.Value) - 1))

            'set_cor | loja AVLE
            If InStr(celula.Value, " - ") = 0 And InStr(celula.Value, "(") = 0 Then .Cells(celula.Row, 4).Value = StrConv(Split(celula.Value, " ")(UBound(Split(celula.Value, " "))), vbProperCase)
        Next
    End With
End Sub

Public Function achei(isso As String, naquilo As String) As Boolean
    achei = InStr(LCase(naquilo), isso) <> 0
End Function

'get_color | loja TELA STUDIO
Function get_cor(descricao As String) As String
    Dim palavras() As String: palavras = Split(descricao, " ")
    If InStr(descricao, "cor:") <> 0 Then
        If InStr(descricao, ";t") = 0 Then get_cor = Trim(Right(descricao, Len(descricao) - InStr(descricao, "cor:") - 3))
        If InStr(descricao, ";t") <> 0 Then get_cor = Trim(Mid(descricao, InStr(descricao, "cor:") + 4, InStr(descricao, ";t") - (InStr(descricao, "cor:") + 4)))
    ElseIf InStr(descricao, "cores:") <> 0 Then
        get_cor = Trim(Right(descricao, Len(descricao) - InStr(descricao, "cores:") - 5))
    ElseIf InStr(descricao, "color:") <> 0 Then
        get_cor = Trim(Right(descricao, Len(descricao) - InStr(descricao, "color:") - 5))
    ElseIf Len(palavras(UBound(palavras))) = 1 Then
        get_cor = palavras(UBound(palavras) - 2) & " " & palavras(UBound(palavras) - 1)
    End If
    get_cor = StrConv(get_cor, vbProperCase)
End Function

Function get_color(ByVal descricao As String, label As String, label_len As Integer) As String
    If InStr(LCase(descricao), label) <> 0 Then
        descricao = Left(Right(descricao, Len(descricao) - InStr(LCase(descricao), label) - label_len), InStr(descricao, ";"))
        If InStr(descricao, ";") <> 0 Then get_color = Trim(Left(descricao, InStr(descricao, ";") - 1))
    End If
End Function

'get_tamanho | loja TELA STUDIO
Function get_tamanho(descricao As String) As String
    Dim palavras() As String: palavras = Split(descricao, " ")
    If InStr(descricao, "tamanho:") <> 0 Then
        If InStr(descricao, ";c") <> 0 Then get_tamanho = Trim(Mid(descricao, InStr(descricao, "tamanho:") + 8, InStr(descricao, ";") - (InStr(descricao, "tamanho:") + 8)))
        If InStr(descricao, ";c") = 0 Then get_tamanho = Trim(Right(descricao, Len(descricao) - (InStr(descricao, "tamanho:") + 7)))
    ElseIf InStr(descricao, "tamanhos:") Then
        get_tamanho = Trim(Mid(descricao, InStr(descricao, "tamanhos:") + 9, InStr(descricao, ";") - (InStr(descricao, "tamanhos:") + 9)))
    ElseIf Len(palavras(UBound(palavras))) = 1 Then
        get_tamanho = palavras(UBound(palavras))
    End If
    get_tamanho = UCase(get_tamanho)
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

Public Function format_cell(celula As Range, Optional cell_style As String = "Normal", Optional cor As Long = 0) As Range
    With celula
        .Style = cell_style
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlRight
        If cell_style = "Normal" Then
            .Font.Bold = True
            .Interior.Color = cor
        End If
    End With
    Set format_cell = celula
End Function

Public Sub liga_desliga(on_off As Boolean)
  With Application
    If on_off Then .Calculation = xlCalculationAutomatic
    If Not on_off Then .Calculation = xlCalculationManual
    .ScreenUpdating = on_off
    .DisplayAlerts = on_off
  End With
End Sub

'loja NERIAJE
Sub atualizar()
    Call liga_desliga(False)
    Dim base_neriage As Workbook: Set base_neriage = Workbooks.Open(Replace(ThisWorkbook.FullName, ".xlsm", ".xlsx"))
    Call from_base(base_neriage, "PRODUTOS")
    Call from_base(base_neriage, "VENDAS")
    base_neriage.Close
    Call liga_desliga(True)
End Sub

Private Sub from_base(base_neriage As Workbook, sufix As String)
    base_neriage.Sheets("BASE_" & sufix).Range("A1").CurrentRegion.Copy
    ThisWorkbook.Sheets("BASE_" & sufix).Range("A6").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
End Sub