Sub preenche_cnae_primario()
    Call preenche_template("CNAEs Primários")
End Sub

Sub preenche_cnae_secundario()
    Call preenche_template("CNAEs Secundários")
End Sub

Private Sub preenche_template(nome_planilha As String)
    Application.ScreenUpdating = False

    Dim ultima_info As Integer: ultima_info = Sheets(nome_planilha).Range("A1").End(xlToRight).Column
    Dim ufColumn As Integer: ufColumn = Sheets(nome_planilha).Range("1:1").Find("uf").Column

    Sheets("final").Select
    ActiveSheet.Cells.Delete Shift:=xlUp
    ActiveSheet.Shapes.SelectAll
    Selection.Delete

    ActiveSheet.Range("A1").Select
    For linha = 2 To Sheets(nome_planilha).Range("A1048576").End(xlUp).Row

        Select Case Sheets(nome_planilha).Cells(linha, ufColumn).Value

            Case "RO", "AC", "AM", "PA", "AP", "RR"
                Sheets("templates").Range("A1:I50").Copy

            Case "PR", "SC", "RS"
                Sheets("templates").Range("A51:I100").Copy

            Case "RJ", "SP", "MG", "ES"
                Sheets("templates").Range("A151:I200").Copy

            Case "MT", "MS", "GO", "DF"
                Sheets("templates").Range("A201:I250").Copy

            Case Else
                Sheets("templates").Range("A101:I150").Copy
        End Select

        If 2 < linha Then Sheets("final").Range("A" & (Sheets("final").Range("A1048576").End(xlUp).Row + 1)).Select
        Selection.PasteSpecial Paste:=xlPasteColumnWidths
        ActiveSheet.Paste

        On Error Resume Next
        For Each coluna In Sheets(nome_planilha).Range(Sheets(nome_planilha).Range("A1"), Sheets(nome_planilha).Cells(1, ultima_info))
            achei = Selection.Find("{{" & coluna.Value & "}}")
            If Not IsEmpty(achei) Then Call Selection.Replace(achei, Sheets(nome_planilha).Cells(linha, coluna.Column))
        Next
        On Error GoTo 0
    Next
    Sheets("final").Range("A1:I" & Sheets("final").Range("A1048576").End(xlUp).Row).ExportAsFixedFormat Filename:=ThisWorkbook.Path & "\Relat�rio - " & nome_planilha & ".pdf", Type:=xlTypePDF
    Sheets("main").Select
    Application.ScreenUpdating = True
End Sub