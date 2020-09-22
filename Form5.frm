VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Edit City Coordinates"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   3255
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   3255
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2520
      Top             =   1800
   End
   Begin VB.CommandButton cmdPick 
      Caption         =   "Show Map"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Top             =   720
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   3960
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   3015
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Update"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.Frame fraCityCoords 
      Caption         =   "City Coordinates:"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
      Begin VB.TextBox txtCity 
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         MaxLength       =   20
         TabIndex        =   10
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtY 
         Enabled         =   0   'False
         Height          =   285
         Left            =   960
         MaxLength       =   4
         TabIndex        =   7
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox txtX 
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         MaxLength       =   4
         TabIndex        =   6
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         Caption         =   "Y:"
         Height          =   195
         Index           =   2
         Left            =   1080
         TabIndex        =   5
         Top             =   840
         Width           =   150
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         Caption         =   "X:"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   4
         Top             =   840
         Width           =   150
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         Caption         =   "City:"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   345
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private SelCityIndex As Integer


Private Sub cmdClose_Click()
  ClearBlinkingPoints
  Me.Hide
End Sub


Private Sub cmdPick_Click()
  If Form4.Visible Then
    PickOnMap = Not PickOnMap
  Else
    Form4.Show
  End If
  Timer1_Timer
End Sub


Private Sub cmdSave_Click()
  Dim Point As POINTAPI
  Dim Tmp   As String
  If SelCityIndex Then
    Point.X = Val(txtX)
    Point.Y = Val(txtY)
    WriteCoords Point, SelCityIndex
    List
  End If
End Sub


Private Sub Form_Load()
  List
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Me.Hide
  Cancel = True
End Sub


Private Sub List1_Click()
  Dim Point As POINTAPI
  Dim i     As Integer
  i = List1.ListIndex
  EnableControls (i >= 0)
  ClearBlinkingPoints
  If i >= 0 Then
    GetCity = String(20, " ")
    Get #2, i + 1, GetCity
    Get #3, i + 1, Point
    txtCity = Trim(GetCity)
    txtX = Point.X
    txtY = Point.Y
    AddBlinkingPoint Point, , , True
    SelCityIndex = i + 1
  Else
    SelCityIndex = 0
    txtCity = ""
  End If
  GetCity = String(20, " ")
End Sub


Private Sub EnableControls(ByVal bEnable As Boolean)
  txtY.Enabled = bEnable
  txtX.Enabled = bEnable
  cmdSave.Enabled = bEnable
  cmdPick.Enabled = bEnable
  txtCity.Enabled = bEnable
End Sub


Private Sub Timer1_Timer()
  Dim bMap As Boolean
  If Not Me.Visible Then Exit Sub
  If bMap <> Form4.Visible Then
    If Form4.Visible Then
      cmdPick.Caption = "Pick on Map"
    Else
      cmdPick.Caption = "Show Map"
    End If
    RePlaceCentered Me
    bMap = Form4.Visible
  End If
End Sub


Private Sub txtX_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
  Case 8, Asc("0") To Asc("9")
  Case 13
    txtY.SetFocus
  Case Else
    KeyAscii = 0
  End Select
End Sub


Private Sub List()
  Dim Point As POINTAPI
  Dim sCity As String * 20
  Dim i     As Integer
  
  List1.Clear
  For i = 1 To NumCities()
    sCity = String(20, " ")
    Get #2, i, sCity
    Point = ReadCoords(i)
    List1.AddItem AddTabs(Left(Trim(sCity), 19), 2) & _
                  AddTabs("X: " & Point.X, 1) & _
                  "Y: " & Point.Y
  Next i
  If SelCityIndex < 1 Then SelCityIndex = 1
  If SelCityIndex > List1.ListCount Then SelCityIndex = List1.ListCount
  If List1.ListCount > 0 Then
    List1.ListIndex = SelCityIndex - 1
  End If
End Sub


Private Sub txtY_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
  Case 8, Asc("0") To Asc("9")
  Case 13
    txtY.SetFocus
  Case Else
    KeyAscii = 0
  End Select
End Sub
