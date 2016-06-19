Attribute VB_Name = "select_in_powerclip"
Public sr2 As New ShapeRange
Public sr2_2 As New ShapeRange
Public sPow As Shape
Public sPowForm_Active As Boolean

Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function SetLayeredWindowAttributes& Lib "user32" (ByVal hwnd&, ByVal crKey&, ByVal bAlpha As Byte, ByVal dwFlags&)
Private Declare Function SetWindowLongW& Lib "user32" (ByVal hwnd&, ByVal nIndex%, ByVal dwNewLong&)
Private Declare Function GetWindowLongW& Lib "user32" (ByVal hwnd&, ByVal nIndex%)
Public Type mySh
  x As Double
  y As Double
  W As Double
  H As Double
  mySh As Shape
  number As Integer
End Type
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Integer, ByVal lpNewItem As Any) As Long
Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hwnd As Long, lprc As RECT) As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As MSG, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Declare Function GetActiveWindow Lib "user32" () As Long
Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Type POINTAPI
  x As Long
  y As Long
End Type

Type MSG
  hwnd As Long
  message As Long
  wParam As Long
  lParam As Long
  time As Long
  pt As POINTAPI
End Type

Sub run()
  If ActiveDocument Is Nothing Then Exit Sub
  main_form.Show 0
End Sub

Function PopMenuList(str1 As String, mx As Long, my As Long) As Long
  Const MF_ENABLED = &H0
  Const TPM_LEFTALIGN = &H0
  Const MF_SEPARATOR = &H800
  Dim msgdata As MSG
  Dim rectdata As RECT
  Dim Cursor As POINTAPI
  Dim i As Long
  Dim j As Long
  Dim last As Long
  Dim hMenu As Long 'хэндл объекта окна меню
  Dim id As Integer
  Dim junk As Long
  
  hMenu = CreatePopupMenu() 'Создание объекта окна меню
  id = 1 ' Счетчик, задающий значение, которое вернет функция при выборе соответствующего пункта меню
  For i = 1 To 2 ' Добавление в меню пунктов
    junk = AppendMenu(hMenu, MF_ENABLED, i, CStr(Split(str1, "|")(i - 1))) ' Добавление нового пункта меню
    'If i < 4 Then junk = AppendMenu(hMenu, MF_SEPARATOR, 0, "") ' Добавление сепаратора
  Next i
  If mx = 0 And my = 0 Then
  Call GetCursorPos(Cursor) ' Получение текущих координат курсора мыши
  mx = Cursor.x - 5 ' Поправка
  my = Cursor.y + 10
  End If
  junk = TrackPopupMenu(hMenu, TPM_LEFTALIGN, mx, my, 0, GetActiveWindow(), rectdata) ' Визуализация объекта
  junk = GetMessage(msgdata, GetActiveWindow(), 0, 0) ' Ожидание события
  i = Abs(msgdata.wParam)
  If msgdata.message = 273 Then PopMenuList = i ' Присвоение возвращаемого значения
  Call DestroyMenu(hMenu)
End Function
