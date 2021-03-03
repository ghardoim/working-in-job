Sub apagarimagem()
'
' apagarimagem Macro
'
  With Application
    .ScreenUpdating = False
    .DisplayAlerts = False
  End With

  ActiveSheet.Shapes.Range(Array("Immagine 4")).Select
  Selection.Delete
  ActiveSheet.Shapes.Range(Array("Picture 1")).Select
  Selection.Delete

  With Application
    .ScreenUpdating = True
    .DisplayAlerts = True
  End With
End Sub