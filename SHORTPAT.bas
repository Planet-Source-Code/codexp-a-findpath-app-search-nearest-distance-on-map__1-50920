Attribute VB_Name = "SHORTPAT"
Option Explicit

Public Const defTabSpc = 12

'[Type Declarations Section]    '
Public Type StreetType
  From  As Integer
  To    As Integer
  Wgt   As Integer
End Type

Public Type POINTAPI
  X     As Long
  Y     As Long
End Type

Public Type BlinkingPoint
  Pos     As POINTAPI
  Color   As Long
  Color2  As Long
  BState  As Integer
  Blink   As Boolean
End Type

'[Private Declarations Section] '
Private Street    As StreetType

'[Public Declarations Section]  '
Public GetCity    As String * 20
Public FromCity   As String * 20
Public ToCity     As String * 20
Public TString    As String * 20
Public Point      As POINTAPI
Public PickOnMap  As Boolean
Public LastPath   As String
Public BlinkPt()  As BlinkingPoint
Public NumBPts    As Integer
Public tmrBPts    As Long


Public Sub MainInit()
  Dim RecSize As Integer
  RecSize = Len(Street)
  Open DataPath & "street.dat" _
       For Random Access Read Write Lock Read Write _
       As #1 Len = RecSize
  GetCity = String(20, " ")
  RecSize = Len(GetCity)
  Open DataPath & "City.dat" _
       For Random Access Read Write Lock Read Write _
       As #2 Len = RecSize
  RecSize = Len(Point)
  Open DataPath & "Coords.dat" _
       For Random Access Read Write Lock Read Write _
       As #3 Len = RecSize
End Sub


Public Function DataPath() As String
  DataPath = AppPath & "Data"
  If UCase(Dir(DataPath, vbDirectory)) <> "DATA" Then
    MkDir DataPath
  End If
  DataPath = DataPath & "\"
End Function


Public Function AppPath() As String
  AppPath = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\")
End Function


Public Function NumStreets() As Integer
  NumStreets = LOF(1) / Len(Street)
End Function


Public Function NumCities() As Integer
  NumCities = LOF(2) / Len(GetCity)
End Function


Public Function NumCoords() As Integer
  NumCoords = LOF(3) / Len(Point)
End Function


Public Sub RePlaceCentered(fWin As Form)
  Dim w As Single
  If Form4.Visible Then
    w = Form4.Width + fWin.Width
    Form4.Left = (Screen.Width - w) / 2 + fWin.Width
  Else
    w = fWin.Width
  End If
  w = (Screen.Width - w) / 2
  fWin.Left = IIf(w < 0, 0, w)
  fWin.Top = (Screen.Height - fWin.Height) / 2
End Sub


Public Function Tabs(ByVal sText As String, Optional ByVal NumTabs As Integer = 1, Optional ByVal TabSpacing As Integer = defTabSpc) As String
  If (NumTabs < 1) Or (Len(sText) > NumTabs * TabSpacing) Then Exit Function
  If TabSpacing < 1 Then TabSpacing = 1
  If Len(sText) > 0 Then
    NumTabs = NumTabs - Int(Len(sText) / TabSpacing)
  End If
  Tabs = String(NumTabs, Chr(9))
End Function


Public Function AddTabs(ByVal sText As String, Optional ByVal NumTabs As Integer = 1, Optional ByVal TabSpacing As Integer = defTabSpc) As String
  AddTabs = sText & Tabs(sText, NumTabs, TabSpacing)
End Function


Public Function ReadCoords(ByVal Rec As Integer) As POINTAPI
  If (Rec < 1) Or (Rec > NumCoords()) Then Exit Function
  Get #3, Rec, ReadCoords
End Function


Public Function WriteCoords(pCrd As POINTAPI, Optional ByVal Rec As Integer) As Integer
  If Rec < 1 Then Rec = NumCoords() + 1
  Put #3, Rec, pCrd
  WriteCoords = Rec
End Function


Public Function ReadCity(ByVal Rec As Integer) As String
  Dim Tmp As String * 20
  If (Rec < 1) Or (Rec > NumCities()) Then Exit Function
  Tmp = String(20, " ")
  Get #2, Rec, Tmp
  ReadCity = Trim(Tmp)
End Function


Public Function WriteCity(sCity As String, Optional ByVal Rec As Integer) As Integer
  Dim Tmp As String * 20
  If (Rec < 1) Or (Rec > NumCities()) Then Rec = NumCities() + 1
  Tmp = Mid(sCity, 1, 20)
  Put #2, Rec, Tmp
  WriteCity = Rec
End Function


Public Sub AddBlinkingPoint(ptPos As POINTAPI, Optional ByVal ptColor As Long = vbGreen, Optional ByVal ptColor2 As Long = vbBlack, Optional ByVal DoBlink As Boolean)
  NumBPts = NumBPts + 1
  ReDim Preserve BlinkPt(0 To NumBPts)
  BlinkPt(NumBPts).Pos = ptPos
  BlinkPt(NumBPts).Color = ptColor
  BlinkPt(NumBPts).Color2 = ptColor2
  BlinkPt(NumBPts).Blink = DoBlink
  tmrBPts = Timer
End Sub


Public Sub ClearBlinkingPoints()
  NumBPts = 0
  ReDim BlinkPt(0)
  tmrBPts = Timer
End Sub


Public Sub ClearMap()
  LastPath = ""
  ClearBlinkingPoints
End Sub


Public Function GetCityIndex(ByVal sCity As String, Optional ByVal bAddNew As Boolean, Optional ByVal bSilent As Boolean) As Integer
  Dim RetVal  As Integer
  Dim i       As Integer
  sCity = Trim(sCity)
  If Len(sCity) = 0 Then Exit Function
  For i = 1 To NumCities()
    If UCase(ReadCity(i)) = UCase(sCity) Then
      RetVal = i
      Exit For
    End If
  Next i
  If bAddNew And (RetVal = 0) Then
    RetVal = WriteCity(sCity)
    If (Not bSilent) And (RetVal > 0) Then
      MsgBox "Adding City: """ & sCity & """!", vbInformation
    End If
  End If
  GetCityIndex = RetVal
End Function
