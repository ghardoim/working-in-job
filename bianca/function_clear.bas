Sub LimparRawData()
Attribute LimparRawData.VB_ProcData.VB_Invoke_Func = " \n14"
'
' LimparRawData Macro
'

'

With Application
    .ScreenUpdating = False
    .DisplayAlerts = False
End With

    Range("A1:S177126").Select
    Selection.Clear
    
With Application
    .ScreenUpdating = True
    .DisplayAlerts = True
End With
    
End Sub
