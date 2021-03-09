Private Sub Workbook_Open()
  Call liga_desliga(False)
  Dim senha As String: senha = "*******"
  Dim resposta As String: resposta = InputBox("INFORME A SENHA PARA INICIAR", "SENHA")

  If senha <> resposta Then
    MsgBox ("VOCÊ NÃO TEM ACESSO A ESSA INFORMAÇÃO")
    Application.Quit
  End If

  FrmCtrlFinanceiro.Show
End Sub