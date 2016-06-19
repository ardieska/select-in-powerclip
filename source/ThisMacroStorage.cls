VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisMacroStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub GlobalMacroStorage_SelectionChange()
  Dim s1 As Shape
  Set s1 = ActiveShape
  
  On Error GoTo myEnd
  If sPowForm_Active = True Then
      If ActiveShape Is Nothing Then Exit Sub
      If ActiveTool = 1 Or 28 Then
          If Not s1.PowerClip Is Nothing Then
              main_form.update_bootom2
          Else
              If s1.PowerClipParent.Type <> 0 Then
              ' Do Nothing
              End If
          End If
      End If
  Else
      Exit Sub
  End If
  Exit Sub
myEnd:
  main_form.clerMyList
End Sub

