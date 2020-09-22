VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "MAP"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9135
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   9135
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   120
   End
   Begin VB.PictureBox picMap 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   5850
      Left            =   0
      Picture         =   "Form4.frx":0000
      ScaleHeight     =   5790
      ScaleWidth      =   9105
      TabIndex        =   0
      Top             =   0
      Width           =   9165
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
  Me.Left = (Screen.Width - Me.Width) / 2
  Me.Top = (Screen.Height - Me.Height) / 2
End Sub


Private Sub picMap_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Select Case Button
  Case vbLeftButton
    If PickOnMap Then
      If Form5.Visible Then
        Form5.txtX = X / Screen.TwipsPerPixelX
        Form5.txtY = Y / Screen.TwipsPerPixelY
      End If
      PickOnMap = False
    End If
  End Select
End Sub


Private Sub Timer1_Timer()
  Static oPickOnMap As Boolean
  Static oPath      As String
  Static iBlink     As Integer
  Static oTmrBPts   As Long
  If Not Me.Visible Then Exit Sub
  ' Pick a Point on the Map '
  If PickOnMap <> oPickOnMap Then
    If PickOnMap Then
      picMap.MousePointer = 2
    Else
      picMap.MousePointer = 0
    End If
    oPickOnMap = PickOnMap
  End If
  ' Redraw last Path        '
  If oPath <> LastPath Then
    DrawLastPath
    oPath = LastPath
  End If
  ' Redraw Location Points  '
  If iBlink = 0 Then
    DrawBlinkingPoints
  End If
  If oTmrBPts <> tmrBPts Then
    picMap.Cls
    DrawLastPath
    oTmrBPts = tmrBPts
  End If
  iBlink = (iBlink + 1) Mod 5
End Sub


Private Sub DrawLastPath()
  Static iRecur As Integer
  Dim Location  As Integer
  Dim sPath     As String
  Dim P1        As POINTAPI
  Dim P2        As POINTAPI
  Dim t         As Integer
  Dim n         As Integer
  Dim c         As Long
  
  If iRecur = 0 Then picMap.Cls
  If (Len(Trim(LastPath)) = 0) Then Exit Sub
  sPath = LastPath
  
  iRecur = iRecur + 1
  c = IIf(iRecur = 2, RGB(255, 150, 0), vbBlack)
  Location = InStr(sPath, ",")
  Do While Location > 0
    t = Val(Mid(sPath, 1, Location - 1))
    sPath = Mid(sPath, Location + 1)
    Location = InStr(sPath, ",")
    If n > 0 Then
      P2 = ReadCoords(t)
      picMap.DrawWidth = IIf(iRecur < 2, 4, 2)
      picMap.Line (P1.X * Screen.TwipsPerPixelX, P1.Y * Screen.TwipsPerPixelY)- _
                  (P2.X * Screen.TwipsPerPixelX, P2.Y * Screen.TwipsPerPixelY), c
      P1 = P2
    Else
      P1 = ReadCoords(t)
    End If
    n = 1
  Loop
  t = Val(sPath)
  If n > 0 Then
    P2 = ReadCoords(t)
    picMap.DrawWidth = IIf(iRecur < 2, 4, 2)
    picMap.Line (P1.X * Screen.TwipsPerPixelX, P1.Y * Screen.TwipsPerPixelY)- _
                (P2.X * Screen.TwipsPerPixelX, P2.Y * Screen.TwipsPerPixelY), c
  End If
  If iRecur < 2 Then DrawLastPath
  iRecur = iRecur - 1
End Sub


Private Sub DrawBlinkingPoints()
  Dim pt  As BlinkingPoint
  Dim i   As Integer
  Dim X   As Single
  Dim Y   As Single
  Dim c   As Long
  i = 1
  While i <= NumBPts
    pt = BlinkPt(i)
    X = (pt.Pos.X) * Screen.TwipsPerPixelX
    Y = (pt.Pos.Y) * Screen.TwipsPerPixelY
    picMap.DrawWidth = 8
    picMap.Circle (X, Y), 1, vbBlack
    c = IIf(pt.BState > 0, pt.Color2, pt.Color)
    picMap.DrawWidth = 6
    picMap.Circle (X, Y), 1, c
    If pt.Blink Then pt.BState = (pt.BState + 1) Mod 2
    BlinkPt(i) = pt
    i = i + 1
  Wend
End Sub
