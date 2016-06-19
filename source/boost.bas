Attribute VB_Name = "boost"
Public Sub boostStart(Optional unDo$)
  If unDo <> "" Then ActiveDocument.BeginCommandGroup unDo
  Optimization = True
  EventsEnabled = False
  ActiveDocument.SaveSettings
  ActiveDocument.PreserveSelection = False
End Sub

Public Sub boostFinish(Optional ByVal endUndoGroup% = False, Optional ByVal doRedraw As Boolean = True)
  ActiveDocument.PreserveSelection = True
  ActiveDocument.RestoreSettings
  EventsEnabled = True
  Optimization = False
  If endUndoGroup Then ActiveDocument.EndCommandGroup
  
  If doRedraw Then
    CorelScript.RedrawScreen
    ActiveWindow.Refresh
    Application.Refresh
  End If
End Sub
