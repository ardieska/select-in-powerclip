Attribute VB_Name = "select_by_mouse"
Declare Function GetKeyState& Lib "user32" (ByVal vKey&)

Sub run()
  If ActiveDocument Is Nothing Then Exit Sub

  Dim s As Shape, s2 As Shape, sp As Shape, sr As New ShapeRange
  Dim x#, y#, Shift&, myMin As Boolean
  Dim sCon As ShapeRange
  Set sCon = ActiveSelectionRange

  If ActiveDocument.GetUserClick(x, y, Shift, -1, True, 313) Then Exit Sub

  boostStart

  Set s = ActivePage.SelectShapesAtPoint(x, y, False)
  Set sr = s.Shapes.All
  ActiveDocument.ClearSelection

  If sr.Count > 1 Then
    For i = sr.Count To 1 Step -1
      If Not sr(i).PowerClip Is Nothing Then Set s = sr(i): Exit For
    Next
    ActiveDocument.ClearSelection
    s.CreateSelection
  ElseIf sr.Count = 1 Then
    Set s = sr(1)
    If Not s.PowerClip Is Nothing Then
      ActiveDocument.ClearSelection
      s.CreateSelection
    End If
  ElseIf sr.Count = 0 Then
    boostFinish endUndoGroup:=False, doRedraw:=False
    Exit Sub
  End If

  'On Error Resume Next
  If Not s.PowerClip Is Nothing Then
    Set sp = s
    Set sr = sp.PowerClip.Shapes.All

    For i = 1 To sr.Count
      If sr(i).IsOnShape(x, y) <> cdrOutsideShape Then
        Set s2 = sr(i)
        Exit For
      End If
    Next

    If s2 Is Nothing Then
      If (GetKeyState(vbKeyControl) And &HFF80) Then
        ActiveDocument.ClearSelection
        sCon.CreateSelection
        sp.AddToSelection
      Else
        ActiveDocument.ClearSelection
        sCon.CreateSelection
      End If
      boostFinish endUndoGroup:=False, doRedraw:=False
      Exit Sub
    End If

    If (GetKeyState(vbKeyControl) And &HFF80) Then
      For i = 1 To sCon.Count
        If s2 Is sCon(i) Then myMin = True
      Next

      If myMin = True Then
        ActiveDocument.ClearSelection
        sCon.CreateSelection
        s2.RemoveFromSelection
      Else
        ActiveDocument.ClearSelection
        sCon.CreateSelection
        s2.AddToSelection
      End If
    Else
    ActiveDocument.ClearSelection
    s2.CreateSelection
    End If
  End If

  boostFinish endUndoGroup:=False, doRedraw:=True
End Sub
