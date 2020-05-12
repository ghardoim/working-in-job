Sub LimparRawData()
Attribute LimparRawData.VB_ProcData.VB_Invoke_Func = " \n14"
'
' LimparRawData Macro
'
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
    End With
    
    Range(Range("A1"), Range("A1").End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Clear
        
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
    End With
End Sub
