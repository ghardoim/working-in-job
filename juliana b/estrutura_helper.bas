Public planilha As Workbook
Public Const ULTIMA_CELULA As String = "1048576"

Public Sub msg_de_nao_preenchido(nome As String, Optional oa As String = "O")
  Call MsgBox("POR FAVOR, INFORME " & oa & " " & nome & "!", vbExclamation, nome & " NÃO INFORMADO")
End Sub

Public Function eh_valido(campo As Object) As Boolean
  eh_valido = True:
  If campo.Value = "" Or campo.Value = False Or campo.Value = 0 Then campo.SetFocus: eh_valido = False
End Function

Public Function ultima_linha(nome_aba As String, Optional coluna As String = "A") As Integer
  ultima_linha = Sheets(nome_aba).Range(coluna & ULTIMA_CELULA).End(xlUp).Row + 1
End Function

Public Sub listar(tela As MSForms.ListBox, nome_aba As String, numero_de_colunas As Integer, Optional col_fim As String = "J")
  With tela
    .ColumnCount = numero_de_colunas
    .ColumnHeads = True
    .RowSource = nome_aba & "!A2:" & col_fim & ultima_linha(nome_aba, col_fim)
  End With
End Sub

Private Sub fixa_cabecalho()
  With ActiveWindow
    .SplitRow = 1
    .FreezePanes = True
  End With
End Sub

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

  With planilha.Sheets(3)
    .Name = "AUXILIAR"
    .Visible = False
  End With

  Call planilha.SaveAs("FINANCEIRO #" & UCase(MonthName(Month(Date), True)) & Right(Year(Date), 2), xlOpenXMLWorkbookMacroEnabled)
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
      .ThemeColor = xlThemeColorDark2
      .TintAndShade = -9.99786370433668E-02
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

  PasswordInit = "Private Sub Workbook_Open()" & vbNewLine & "With Application" & vbNewLine & _
                  ".DisplayAlerts = False" & vbNewLine & ".Visible = False" & vbNewLine & _
                  "Dim senha As String: senha = ""******""" & vbNewLine & _
                  "Dim resposta As String: resposta = InputBox(""INFORME A SENHA PARA INICIAR"", ""SENHA"")" & vbNewLine & _
                  "If senha <> resposta Then" & vbNewLine & "MsgBox (""VOCÊ NÃO TEM ACESSO A ESSA INFORMAÇÃO"")" & vbNewLine & _
                  ".Quit" & vbNewLine & "End If:" & vbNewLine & ".DisplayAlerts = True" & vbNewLine & _
                  ".Visible = True" & vbNewLine & "End With" & vbNewLine & "End Sub"

End Function

Public Sub liga_desliga(on_off As Boolean)
  With Application
    If on_off Then .Calculation = xlCalculationAutomatic
    If Not on_off Then .Calculation = xlCalculationManual
    .ScreenUpdating = on_off
    .DisplayAlerts = on_off
    .Visible = on_off
  End With
End Sub

Public Sub atualiza_total(lbl As MSForms.Label, aba As Worksheet, coluna As String)
  lbl.Caption = "R$ " & WorksheetFunction.Sum(aba.Range(coluna & ":" & coluna))
End Sub