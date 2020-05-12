Sub Pedra_Filosofal()
Attribute Pedra_Filosofal.VB_ProcData.VB_Invoke_Func = "p\n14"

Dim sPath As String, sName As String, fName As String
Dim r As Long
Dim shPadrao As Worksheet

    sPath = InputBox("Cole o endere�o da pasta com as planilhas de dados brutos di�rios, no formato abaixo:", "Pasta de dados di�rios", "C:\Users\c117012\Desktop\teste\")
    ThisWorkbook.Worksheets("Auxiliar").Range("D8").Value = sPath
    
With Application
    .ScreenUpdating = False
    .DisplayAlerts = False
End With

Set shPadrao = Sheets("Raw Data")

sName = Dir(sPath & "\*.xlsx")

Do While sName <> ""

    r = shPadrao.Cells(Rows.Count, "B").End(xlUp).Row

    fName = sPath & sName

    Workbooks.Open Filename:=fName, UpdateLinks:=False

    Sheets("Lines").Range("A7:A1446").Copy shPadrao.Range("A" & r + 1)
    Sheets("Main").Range("B7:B1446").Copy shPadrao.Range("B" & r + 1)
    Sheets("Main").Range("L7:L1446").Copy shPadrao.Range("C" & r + 1)
    Sheets("Flares").Range("E7:E1446").Copy shPadrao.Range("D" & r + 1)
    Sheets("Flares").Range("M7:O1446").Copy shPadrao.Range("E" & r + 1)
    Sheets("Flares").Range("U7:U1446").Copy shPadrao.Range("H" & r + 1)
    Sheets("Flares").Range("AC7:AE1446").Copy shPadrao.Range("I" & r + 1)
    Sheets("Flares").Range("AK7:AK1446").Copy shPadrao.Range("L" & r + 1)
    Sheets("Flares").Range("AS7:AU1446").Copy shPadrao.Range("M" & r + 1)
    Sheets("Flares").Range("BA7:BA1446").Copy shPadrao.Range("P" & r + 1)
    Sheets("Flares").Range("BI7:BK1446").Copy shPadrao.Range("Q" & r + 1)
       
ActiveWorkbook.Close SaveChanges:=False

ScapeB: sName = Dir()

Loop

On Error GoTo 0

'cabe�alho
Rows("1:5").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Sheets("Auxiliar").Range("A1:S6").Copy
    Sheets("Raw Data").Range("A1").Select
    ActiveSheet.Paste
    
'Acerto de Colunas e formata��o de datas
    Columns("C:O").ColumnWidth = 15.71
    Columns("B:B").ColumnWidth = 22.14
    Columns("A:A").ColumnWidth = 15.71
    Columns("A:B").Application.CutCopyMode = False
    Selection.Copy
 Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
 :=False, Transpose:=False

Application.CutCopyMode = False
Range("A2").Select

With Application
    .ScreenUpdating = True
    .DisplayAlerts = True
End With

End Sub
