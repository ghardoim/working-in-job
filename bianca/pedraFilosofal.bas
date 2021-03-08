Sub Pedra_Filosofal()
  Attribute Pedra_Filosofal.VB_ProcData.VB_Invoke_Func = "p\n14"

  Range("A1").Select

  Dim sPath As String: sPath = InputBox("Cole o endereço da pasta com as planilhas de dados brutos diários, no formato abaixo:", _
      "Pasta de dados diários", "C:\Users\$(whoami)\Desktop\teste\")
      
  ThisWorkbook.Worksheets("Auxiliar").Range("D8").value = sPath
  
  Call screenUpdate(False)
  Call insereCabecalho
  
  Dim shPadrao As Worksheet: Set shPadrao = ThisWorkbook.Sheets("Raw Data")
  Dim fileName As String: fileName = Dir(sPath & "\*.xlsx")
  
  Do While fileName <> ""
                  
    Workbooks.Open fileName:=sPath & fileName, UpdateLinks:=False
    
    Call findAndPaste(shPadrao, Sheets("Main"), "Date and Time", True)
    Call findAndPaste(shPadrao, Sheets("Main"), "main")
    
    For flareNumber = 1 To 3
      Dim flare As String: flare = "Flare_" & flareNumber
      Call findAndPaste(shPadrao, Sheets(flare), "LFG flow normalized*" & flareNumber, , True)
      If flareNumber <> 3 Then
        Call findAndPaste(shPadrao, Sheets(flare), "Exhaust gas temperature*" & flareNumber, , True)
        Call findAndPaste(shPadrao, Sheets(flare), "CH4 fraction exhaust gas*" & flareNumber, , True)
        Call findAndPaste(shPadrao, Sheets(flare), "O2 fraction exhaust gas*" & flareNumber, , True)
      End If
      Call findAndPaste(shPadrao, Sheets(flare), "LFG flow normalized LFG50*" & flareNumber, , True, True)
    Next
    Workbooks(fileName).Close SaveChanges:=False
        
ScapeB: fileName = Dir()
  
  Loop
  On Error GoTo 0
  
  'Acerto de Colunas e formata��o de datas
  Columns("C:Y").ColumnWidth = 15.71
  Columns("B:B").ColumnWidth = 22.14
  Columns("A:A").ColumnWidth = 15.71
  Range("A2").Select
  Call screenUpdate(True)
End Sub

Sub insereCabecalho()
  Sheets("Auxiliar").Select
  Range(Range("A1"), Range("A1").End(xlToRight)).Select
  Range(Selection, Selection.End(xlDown)).Copy
  Sheets("Raw Data").Select
  ActiveSheet.Paste
End Sub

Sub screenUpdate(on_off As Boolean)
  With Application
    .ScreenUpdating = on_off
    .DisplayAlerts = on_off
  End With
End Sub

Sub findAndPaste(planilhaPrincipal As Worksheet, abaOndeBusco As Variant, infoQueProcuro As String, _
                    Optional findColumnDate As Boolean = False, Optional findFlare As Boolean = False, _
                        Optional selectTwoColumns As Boolean = False)
  Sheets(abaOndeBusco.Name).Select
    
  Dim informacao As Range: Set informacao = planilhaPrincipal.Rows(3).Find(infoQueProcuro)
  Dim ultimalinha As Integer: ultimalinha = planilhaPrincipal.Cells("1048576", informacao.Column).End(xlUp).Row + 1
    
  Dim nomeDaInfo As String: nomeDaInfo = informacao.value
  If Not findColumnDate And Not findFlare Then
    nomeDaInfo = Left(informacao.value, InStr(1, informacao.value, infoQueProcuro) - 2)
  ElseIf findFlare Then
    nomeDaInfo = Left(informacao.value, InStr(1, informacao.value, "flare") - 2)
  End If
    
  Dim ondeAchei As Range: Set ondeAchei = Sheets(abaOndeBusco.Name).Rows(3).Find(nomeDaInfo)
  Cells(ondeAchei.Row, ondeAchei.Column).Offset(4, 0).Select
  If selectTwoColumns Then
    Range(Selection, Selection.Offset(0, 1)).Select
  End If
    
  Sheets(abaOndeBusco.Name).Range(Selection, Selection.End(xlDown)).Copy
  planilhaPrincipal.Cells(ultimalinha, informacao.Column).PasteSpecial _
      Paste:=xlPasteValuesAndNumberFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
End Sub
