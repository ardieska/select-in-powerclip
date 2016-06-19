VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} main_form 
   Caption         =   "Select In Powerclip"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4455
   OleObjectBlob   =   "main_form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "main_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub del_bootom_Click()
  Dim i&, sr3 As New ShapeRange, sr4 As New ShapeRange, s4 As Shape, li As ListItem
  If myListObj.SelectedItem Is Nothing Then Exit Sub
  'i = CLng(myListObj.SelectedItem.Index)
  Set sr3 = sr2.ReverseRange

  For Each li In myListObj.ListItems
    If li.Selected = True Then
    i = CLng(li.Index)
    Set s4 = sr3.Item(i)
    sr4.Add s4
    End If
  Next li

  boostStart "Delete Shape"
  sPow.PowerClip.ExtractShapes
  sr3.RemoveRange sr4
  sr3.AddToPowerClip sPow, cdrFalse
  sr4.Delete

  ActiveDocument.ClearSelection
  sPow.CreateSelection
  boostFinish endUndoGroup:=True

  '==============================================
  For i = myListObj.ListItems.Count To 1 Step -1
    myListObj.ListItems.Remove (i)
  Next i
  scanObj
End Sub

Private Sub ext_bootom_Click()
  Dim i&, sr3 As New ShapeRange, sr4 As New ShapeRange, s4 As Shape, li As ListItem
  If myListObj.SelectedItem Is Nothing Then Exit Sub
  'i = CLng(myListObj.SelectedItem.Index)
  Set sr3 = sr2.ReverseRange

  For Each li In myListObj.ListItems
    If li.Selected = True Then
      i = CLng(li.Index)
      Set s4 = sr3.Item(i)
      sr4.Add s4
    End If
  Next li

  boostStart "Extract Shape"
  sPow.PowerClip.ExtractShapes
  sr3.RemoveRange sr4
  sr3.AddToPowerClip sPow, cdrFalse

  ActiveDocument.ClearSelection
  sPow.CreateSelection
  'sr3.Item(i).CreateSelection
  boostFinish endUndoGroup:=True

  '==============================================
  For i = myListObj.ListItems.Count To 1 Step -1
    myListObj.ListItems.Remove (i)
  Next i
  scanObj
  '==============================================

  ActiveDocument.ClearSelection
  sr4.CreateSelection
End Sub

Private Sub myAlign_Click()
  On Error Resume Next
  s = PopMenuList("Align and Distribute|Align and Distribute (Powerclip)", 0, 0)
  Select Case s
  Case 1
    myNewSelected_Click
    Application.FrameWork.Automation.Invoke "7109900f-4789-4451-9ba7-bb3df86db569"
  Case 2
    myNewSelected_Click
    sPow.AddToSelection
    Application.FrameWork.Automation.Invoke "7109900f-4789-4451-9ba7-bb3df86db569"
  Case 0
    Exit Sub
  End Select
  'retry True
End Sub

Private Sub myNewSelected_Click()
  Dim i&, sr3 As New ShapeRange, li As ListItem
  Dim myLS As Boolean
  If myListObj.SelectedItem Is Nothing Then Exit Sub
  Set sr3 = sr2.ReverseRange

  For Each li In myListObj.ListItems
    If li.Selected = True Then
      i = CLng(li.Index)
      If myLS = False Then
        ActiveDocument.ClearSelection
        sr3.Item(i).CreateSelection
        myLS = True
      Else
        sr3.Item(i).AddToSelection
      End If
    End If
  Next li
End Sub

Private Sub myPowSelAdd_Click()
  On Error Resume Next
  sPow.AddToSelection
End Sub

Private Sub update_bootom_Click()
  clerMyList
  scanObj
  myListObj2.Clear
  scanObj2
End Sub

Private Sub myselectMouse_Click()
  selectObj_mouse
End Sub

Private Sub myListObj_Click()
  Dim i&, sr3 As New ShapeRange, li As ListItem
  Dim myLS As Boolean
  If myListObj.SelectedItem Is Nothing Then Exit Sub
  Set sr3 = sr2.ReverseRange

  For Each li In myListObj.ListItems
    If li.Selected = True Then
      i = CLng(li.Index)
      If myLS = False Then
      ActiveDocument.ClearSelection
      sr3.Item(i).CreateSelection
      myLS = True
      Else
      sr3.Item(i).AddToSelection
      End If
    End If
  Next li
End Sub

Private Sub myListObj_ItemClick(ByVal Item As MSComctlLib.ListItem)
  myListObj_Click
End Sub

Private Sub myListObj2_Change()
  Dim i&, sr3 As New ShapeRange
  On Error Resume Next
  Set sr3 = sr2_2.ReverseRange

  ActiveDocument.ClearSelection
  sr3.Item(CLng(myListObj2.ListIndex) + 1).CreateSelection

  'update_bootom2
  'MsgBox myListObj2.ListIndex
End Sub

Private Sub myListObj_DblClick()
  Dim i&, sr3 As New ShapeRange
  If myListObj.SelectedItem Is Nothing Then Exit Sub
  i = CLng(myListObj.SelectedItem.Index)

  sPow.PowerClip.EnterEditMode
  ActiveDocument.ClearSelection
  Set sr3 = sr2.ReverseRange
  sr3.Item(i).CreateSelection

  '==============================================
  For i = myListObj.ListItems.Count To 1 Step -1
    myListObj.ListItems.Remove (i)
  Next i
  scanObj
End Sub

Sub update_bootom2()
  clerMyList
  scanObj
End Sub

Sub clerMyList()
  For i = myListObj.ListItems.Count To 1 Step -1
    myListObj.ListItems.Remove (i)
  Next i
End Sub

Private Sub UserForm_Initialize()
  sPowForm_Active = True
  'Me.Zoom = 90
  Me.Width = 138.5

  s = GetSetting("SanchoCorelVBA", "SelectInPowerclip", "Pos")
  If Len(s) Then
    startupPosition = 0
    Move CSng(Split(s, " ")(0)), CSng(Split(s, " ")(1))
  End If

  scanObj
  scanObj2
  Copirait.Caption = "Version 4"
  Copirait2.Caption = Chr(169) & " 2007 Sancho"
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  sPowForm_Active = False
  SaveSetting "SanchoCorelVBA", "SelectInPowerclip", "Pos", Left & " " & Top
End Sub

Private Sub scanObj()
  Dim sr As ShapeRange, s As Shape, sp As Shape, c&
  Dim n As Variant, li As ListItem, doc As Document, myIcoN$
  Dim li2 As ListItem, i&

  Set doc = CorelDRAW.ActiveDocument
  doc.Unit = cdrMillimeter

  sr2.RemoveAll

  Set sr = ActiveSelectionRange
  If sr Is Nothing Then GoTo myEnd
  If sr.Count <> 1 Then GoTo myEnd
  Set s = sr.Item(1)
  Set sPow = sr.Item(1)
  c = 0

  On Error Resume Next
  If Not s.PowerClip Is Nothing Then
    c = c + 1

    For Each sp In s.PowerClip.Shapes
    'n = sp.Type
    '===========================================================
    Select Case sp.Type
      Case cdrGroupShape
        n = "": myIcoN = "group"
      Case cdrCurveShape
        n = "": myIcoN = "curve2"
      Case cdrRectangleShape
        n = "": myIcoN = "rectangle"
      Case cdrEllipseShape
        n = "": myIcoN = "ellipse"
      Case cdrPolygonShape
        n = "": myIcoN = "polygon"
      Case cdrPerfectShape
        n = "": myIcoN = "perfect"
      Case cdrSymbolShape
        n = "": myIcoN = "symbol"
      Case cdrTextShape
        n = "": myIcoN = "text"
      Case cdrBitmapShape
        n = "": myIcoN = "bitmap"
        Select Case sp.Bitmap.Mode
          Case cdrCMYKColorImage: n = n + " CMYK"
          Case cdrGrayscaleImage: n = n + " Grayscale"
          Case cdrBlackAndWhiteImage: n = n + " BW"
          Case cdrRGBColorImage: n = n + " RGB"
        End Select
      Case cdrEPSShape
        n = "": myIcoN = "eps"
      Case cdrMeshFillShape
        n = "": myIcoN = "mesh"
      Case cdrArtisticMediaGroupShape
        n = "": myIcoN = "ArtMedia"
    End Select

    If Not sp.PowerClip Is Nothing Then n = n + " (PowerClip)"

    If sp.CanHaveFill Then
      Select Case sp.Fill.Type
      Case cdrNoFill
        n = n + " (No Fill)"
      Case cdrUniformFill
        If sp.Fill.UniformColor.Name = "unnamed color" Then
        n = n + " " + sp.Fill.UniformColor.Name(True)
        Else
        n = n + " " + sp.Fill.UniformColor.Name
        End If
      Case cdrFountainFill
        n = n + " " + "Fountain Fill"
      End Select
    End If

    n = n + " (" & Left(sp.SizeWidth, 6) & " x " & Left(sp.SizeHeight, 6) & ")"
    '===========================================================

    myListObj.SmallIcons = ImageList_ico
    Set li = myListObj.ListItems.Add(c, , n, , myIcoN)
    li.Tag = "obj" + CStr(c)

    sr2.Add sp
    Next
  End If

myEnd:
  If myListObj.ListItems.Count < 1 Then
    ext_bootom.Enabled = False
    del_bootom.Enabled = False
    myNewSelected.Enabled = False
    myAlign.Enabled = False
    myPowSelAdd.Enabled = False
  Else
    ext_bootom.Enabled = True
    del_bootom.Enabled = True
    myNewSelected.Enabled = True
    myAlign.Enabled = True
    myPowSelAdd.Enabled = True
  End If

  For Each li2 In myListObj.ListItems
    If li2.Selected = True Then
      i = CLng(li2.Index)
      myListObj.ListItems(i).Selected = False
    End If
  Next li2
End Sub

Private Sub scanObj2()
  Dim p As Page, doc As Document, s As Shape, n$, c&, sr3 As New ShapeRange
  Dim li As ListItem

  Set doc = CorelDRAW.ActiveDocument
  doc.Unit = cdrMillimeter
  Set p = ActivePage
  c = 0

  sr2_2.RemoveAll

  scanObj2_skan p.Shapes.All, sr2_2
  If sr2_2.Count > 0 Then
    Set sr3 = sr2_2.ReverseRange
  Else
    Exit Sub
  End If

  For Each s In sr3
    c = c + 1
    n = "On "
    n = n + s.Layer.Name

    If s.CanHaveFill Then
      Select Case s.Fill.Type
      Case cdrNoFill
        n = n + " F:None"
      Case cdrUniformFill
        If s.Fill.UniformColor.Name = "unnamed color" Then
          n = n + " F: " + s.Fill.UniformColor.Name(True)
        Else
          n = n + " F:" + s.Fill.UniformColor.Name
        End If
      Case cdrFountainFill
        n = n + " F:" + "Fountain"
      End Select
    End If

    myListObj2.AddItem n
  Next
End Sub

Private Sub scanObj2_skan(sr As ShapeRange, sr2_2 As ShapeRange)
  Dim s As Shape
  On Error Resume Next

  For Each s In sr
    If s.Type = cdrGroupShape Then
      scanObj2_skan s.Shapes.All, sr2_2
    Else
      If Not s.PowerClip Is Nothing Then sr2_2.Add s
    End If
  Next s
End Sub
