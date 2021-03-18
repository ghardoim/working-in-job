Public Sub abre_arquivo()
  Dim arquivo As String: arquivo = Application.GetOpenFilename("Excel Files (*.xlsm), *")
  If InStr(arquivo, "FINANCEIRO") <> 0 Then: Set planilha = Workbooks.Open(arquivo)
  Call liga_desliga(False)
End Sub

Public Sub novo_arquivo()
  Dim excelAPP As Application: Set excelAPP = New Excel.Application
  Set planilha = excelAPP.Workbooks.Add

  Dim VBModule As CodeModule: Set VBModule = excelAPP.VBE.ActiveVBProject.VBComponents.Item("EstaPastaDeTrabalho").CodeModule
  Call VBModule.AddFromString(PasswordInit)

  Dim aba_entrada As Worksheet: Set aba_entrada = planilha.Sheets.Add
  Dim aba_saida As Worksheet: Set aba_saida = planilha.Sheets.Add
  Dim aba_total As Worksheet: Set aba_total = planilha.Sheets.Add: aba_total.Name = "RESULTADO"

  With aba_entrada
    .Name = "ENTRADA"
    .Range("A1:J1") = Array("ADVOGADO", "CLIENTE", "TIPO", "VENCIMENTO", "BOLETO EMITIDO", "NFE EMITIDA", "VALOR", "VALOR PAGO", "IMPOSTO", "VALOR LÍQUIDO")
    Call formata(.Range("A1:J1"))
  End With

  With aba_saida
    .Name = "SAÍDA"
    .Range("A1:F1") = Array("DATA", "FUNCIONÁRIO", "CLIENTE", "TIPO", "DESPESA", "VALOR")
    Call formata(.Range("A1:F1"))
  End With
    
  With planilha.Sheets(4)
    .Name = "AUXILIAR"
    .Visible = False
  End With

  Call liga_desliga(False)
  Call planilha.SaveAs("FINANCEIRO #" & UCase(MonthName(Month(Date), True)) & Right(Year(Date), 2), xlOpenXMLWorkbookMacroEnabled)
  Dim novo As String: novo = planilha.FullName: Call excelAPP.Quit
  Set planilha = Workbooks.Open(novo, IgnoreReadOnlyRecommended:=True)
  Call liga_desliga(False)
End Sub

Private Sub linhas_de_borda(borda As Border)
  With borda
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
  End With
End Sub

Private Sub formata(cabecalho As Range)
  With cabecalho
    .AutoFilter
    .ColumnWidth = 16
    .Font.Bold = True
    With .Interior
      .Pattern = xlSolid
      .PatternColorIndex = xlAutomatic
      .Color = RGB(217, 225, 242)
      .TintAndShade = 0
      .PatternTintAndShade = 0
    End With
    Call linhas_de_borda(.Borders(xlEdgeLeft))
    Call linhas_de_borda(.Borders(xlEdgeTop))
    Call linhas_de_borda(.Borders(xlEdgeBottom))
    Call linhas_de_borda(.Borders(xlEdgeRight))
    Call linhas_de_borda(.Borders(xlInsideVertical))
    Call linhas_de_borda(.Borders(xlInsideHorizontal))
  End With
End Sub

Private Function PasswordInit() As String
  PasswordInit = "Private Sub Workbook_Open()" & vbNewLine & "Dim senha As String: senha = ""julianab""" & vbNewLine & _
    "Dim resposta As String: resposta = InputBox(""INFORME A SENHA PARA INICIAR"", ""SENHA"")" & vbNewLine & _
    "If senha <> resposta Then" & vbNewLine & "MsgBox (""VOCÊ NÃO TEM ACESSO A ESSA INFORMAÇÃO"")" & vbNewLine & _
    "End If" & vbNewLine & "End Sub"
End Function