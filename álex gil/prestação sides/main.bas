Sub init()
    FRMSIdes.Show
End Sub

Sub salvarPDF()
    ThisWorkbook.ExportAsFixedFormat Filename:=ThisWorkbook.Path & "\Prestação SIDES" & ".pdf", Type:=xlTypePDF
End Sub

Sub atualizar_resumo()
    Call liga_desliga(False)
    Dim coluna As Integer: coluna = 3
    semestre = IIf( _
        6 >= Month(CDate(WorksheetFunction.Subtotal(4, Sheets("Receitas").Range("B:B")))), _
            Array("01/31", "02/28", "03/31", "04/30", "05/31", "06/30"), _
            Array("07/31", "08/31", "09/30", "10/31", "11/30", "12/31"))

    For Each mes In semestre
        Sheets("Receitas").Range("A1").AutoFilter field:=2, Operator:=xlFilterValues, Criteria2:=Array(1, mes)
        With Sheets("Prestação").Cells(1, coluna)
            .Value = MonthName(Month(CDate(mes)))
            .Interior.ThemeColor = xlThemeColorAccent5
            .Interior.TintAndShade = -0.25
            .Font.Color = RGB(255, 255, 255)
            .Font.Size = 14
        End With

        Dim linha As Integer: linha = 2
        For Each tipo In Sheets("aux").Range("C1:C5").Value
            Call preenche_total(linha, coluna, tipo)
            linha = linha + 2
        Next
        Call preenche_total(linha, coluna, "<>Recebido")
        Call preenche_total(linha + 2, coluna, "Recebido")
        coluna = coluna + 2
    Next
    Sheets("Receitas").Range("A1").AutoFilter
    Call preenche_total(18, 9, "<>Recebido")
    Call preenche_total(20, 9, "Recebido")
    Sheets("Receitas").Range("A1").AutoFilter

    Call atualizar_saldo
    Call liga_desliga(True)
End Sub

Sub atualizar_saldo()
    With Sheets("Prestação")
        .Cells(20, 3).Value = .Cells(18, 3).Value + .Cells(20, 9).Value - .Cells(18, 9).Value
        .Columns.AutoFit
    End With
End Sub

Private Sub preenche_total(linha As Integer, coluna As Integer, ByVal filtro As String)
    Sheets("Receitas").Range("A1").AutoFilter field:=7, Criteria1:=filtro
    Sheets("Prestação").Cells(linha, coluna).Value = WorksheetFunction.Subtotal(9, Sheets("Receitas").Range("C:C"))
End Sub

Sub liga_desliga(on_off As Boolean)
  With Application
    .Calculation = IIf(on_off, xlCalculationAutomatic, xlCalculationManual)
    .ScreenUpdating = on_off
    .DisplayAlerts = on_off
    .EnableEvents = on_off
  End With
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Call liga_desliga(False)
        If 18 = Target.Row And 3 = Target.Column Then Call atualizar_saldo
    Call liga_desliga(True)
End Sub