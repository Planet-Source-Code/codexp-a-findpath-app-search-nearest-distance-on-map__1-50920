VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Add City and Distance to File"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   3975
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   3975
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3240
      Top             =   2160
   End
   Begin VB.Frame fraEditStreet 
      Caption         =   "Edit Street:"
      Height          =   1575
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   2415
      Begin VB.CommandButton cmdNew 
         Caption         =   "new"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1680
         TabIndex        =   14
         Top             =   0
         Width           =   495
      End
      Begin VB.TextBox FromCity 
         Height          =   285
         Left            =   1080
         MaxLength       =   20
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox ToCity 
         Height          =   285
         Left            =   1080
         MaxLength       =   20
         TabIndex        =   4
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox Distance 
         Height          =   285
         Left            =   1080
         MaxLength       =   5
         TabIndex        =   5
         Text            =   "0"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         Caption         =   "From City:"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   750
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         Caption         =   "To City:"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   570
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         Caption         =   "Distance:"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   675
      End
   End
   Begin VB.ListBox List1 
      Height          =   3570
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   3735
   End
   Begin VB.CommandButton cmdShowMap 
      Caption         =   "Show Map"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add Street"
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      Caption         =   "Distance:"
      Height          =   195
      Index           =   5
      Left            =   3000
      TabIndex        =   13
      Top             =   1800
      Width           =   675
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      Caption         =   "To City:"
      Height          =   195
      Index           =   4
      Left            =   1560
      TabIndex        =   12
      Top             =   1800
      Width           =   570
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      Caption         =   "From City:"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   11
      Top             =   1800
      Width           =   750
   End
   Begin VB.Menu menuCities 
      Caption         =   "Cities"
      Visible         =   0   'False
      Begin VB.Menu mnuCity 
         Caption         =   "City"
         Index           =   0
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'[Declarations Section]'
Private Street  As StreetType
Private TString As String * 20
Private SelStreetIndex As Integer


Private Sub cmdClose_Click()
  MsgBox "There should be " & NumCities() & " cities on file."
  Me.Hide
End Sub


Private Sub cmdAdd_Click()
  Dim oCity As String
  Dim Test  As String
  Dim i     As Integer
  
  FromCity = Trim(FromCity)
  ToCity = Trim(ToCity)
  Distance = Abs(Val(Distance))
  
  If Len(FromCity) = 0 Or Len(ToCity) = 0 Then
    MsgBox "Input Data into From and To Fields!"
    FromCity.SetFocus
    Exit Sub
  End If
  
  If SelStreetIndex < 1 Then SelStreetIndex = NumStreets() + 1
  Street.From = GetCityIndex(FromCity, True)
  Street.To = GetCityIndex(ToCity, True)
  Street.Wgt = Abs(Val(Distance))
  If (Street.From > 0) And (Street.To > 0) Then
    Put #1, SelStreetIndex, Street
  Else
    MsgBox "Error adding new Street!", vbCritical
  End If
  
  ListCity
  FromCity.SetFocus
End Sub


Private Sub cmdNew_Click()
  SelStreetIndex = 0
  List1.ListIndex = -1
  List1_Click
  Timer1_Timer
  FromCity.SetFocus
End Sub


Private Sub cmdShowMap_Click()
  If Form4.Visible Then
    Form4.Hide
  Else
    Form4.Show
  End If
  Timer1_Timer
End Sub

Private Sub Distance_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
  Case 8, Asc("0") To Asc("9")
  Case 13
    cmdAdd.SetFocus
  Case Else
    KeyAscii = 0
  End Select
End Sub


Private Sub Form_Load()
  ListCity
End Sub


Private Sub ListCity()
  Dim fCity As String * 20
  Dim tCity As String * 20
  Dim i     As Integer
  List1.Clear
  For i = 1 To NumStreets()
    fCity = String(20, " ")
    tCity = String(20, " ")
    Get #1, i, Street
    Get #2, Street.From, fCity
    Get #2, Street.To, tCity
    List1.AddItem AddTabs(Left(Trim(fCity), 19), 2) & _
                  AddTabs(Left(Trim(tCity), 19), 2) & _
                  Street.Wgt
  Next i
  GetCity = String(20, " ")
  If SelStreetIndex Then
    If SelStreetIndex > List1.ListCount Then
      List1_Click
    Else
      List1.ListIndex = SelStreetIndex - 1
    End If
  End If
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Me.Hide
  Cancel = True
End Sub


Private Sub FromCity_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
  Case 13
    ToCity.SetFocus
    KeyAscii = 0
  End Select
End Sub


Private Sub List1_Click()
  Dim Street  As StreetType
  Dim Point   As POINTAPI
  Dim i       As Integer
  ClearMap
  i = List1.ListIndex
  If i >= 0 Then
    Get #1, i + 1, Street
    FromCity = ReadCity(Street.From)
    ToCity = ReadCity(Street.To)
    LastPath = Street.From & ", " & Street.To
    AddBlinkingPoint ReadCoords(Street.From), RGB(0, 255, 0), RGB(0, 200, 0), True
    AddBlinkingPoint ReadCoords(Street.To), RGB(255, 0, 0), RGB(150, 0, 0), True
    Distance = Street.Wgt
    SelStreetIndex = i + 1
  Else
    SelStreetIndex = 0
    FromCity = ""
    ToCity = ""
    Distance = "0"
  End If
  GetCity = String(20, " ")
  Timer1_Timer
End Sub


Private Sub Timer1_Timer()
  Static oSelStreetIndex  As Integer
  Static bMap             As Boolean
  
  If Not Me.Visible Then Exit Sub
  If SelStreetIndex <> oSelStreetIndex Then
    If SelStreetIndex > 0 Then
      cmdNew.Enabled = True
      cmdAdd.Caption = "Update Street"
    Else
      cmdNew.Enabled = False
      cmdAdd.Caption = "Add Street"
    End If
    oSelStreetIndex = SelStreetIndex
  End If

  If bMap <> Form4.Visible Then
    If Form4.Visible Then
      cmdShowMap.Caption = "Hide Map"
    Else
      cmdShowMap.Caption = "Show Map"
    End If
    RePlaceCentered Me
    bMap = Form4.Visible
  End If
End Sub


Private Sub ToCity_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
  Case 13
    Distance.SetFocus
    KeyAscii = 0
  End Select
End Sub
