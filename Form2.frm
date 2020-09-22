VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Find Shortest Path"
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
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   3255
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2520
      Top             =   2640
   End
   Begin VB.CommandButton cmdShowMap 
      Caption         =   "Show Map"
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   1200
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   3180
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   3015
   End
   Begin VB.CommandButton cmdShowArrays 
      Caption         =   "Show Arrays"
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdFindPath 
      Caption         =   "Find Path"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Location"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
      Begin VB.CommandButton cmdDstMenu 
         Caption         =   "Ú"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   5.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   13
         Top             =   1200
         Width           =   165
      End
      Begin VB.CommandButton cmdSrcMenu 
         Caption         =   "Ú"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   5.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   12
         Top             =   600
         Width           =   165
      End
      Begin VB.TextBox Target 
         Height          =   285
         Left            =   120
         MaxLength       =   20
         TabIndex        =   3
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox Source 
         Height          =   285
         Left            =   120
         MaxLength       =   20
         TabIndex        =   2
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Mileage 
         Caption         =   "0"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         Caption         =   "Distance:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   1680
         Width           =   675
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         Caption         =   "Destination:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   870
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         Caption         =   "Starting Point:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1035
      End
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
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'[Declarations Section]'
Private NumNodes    As Integer
Private Infinity    As Integer
Private Street()    As StreetType
Private Distance()  As Integer
Private Path()      As Integer
Private Included()  As Integer
Private A_Show      As Integer
Private SelCity     As String


Private Function AllIncluded() As Boolean
  Dim RetVal  As Boolean
  Dim i       As Integer
  RetVal = True             ' Assume all are included
  For i = 1 To NumNodes
    If Included(i) = False Then RetVal = False
  Next
  AllIncluded = RetVal
End Function


Private Sub cmdDstMenu_Click()
  Target_DblClick
End Sub


Private Sub cmdFindPath_Click()
  FindPath
End Sub


Private Sub cmdShowArrays_Click()
  A_Show = Not A_Show
  cmdShowArrays.Caption = IIf(A_Show, "Hide Arrays", "Show Arrays")
End Sub


Private Sub FindPath()
  Dim Location As Integer
  Dim bFirst As Boolean
  Dim sPath As String
  Dim Test  As String
  Dim tt    As String
  Dim ii    As Integer
  Dim i     As Integer
  Dim j     As Integer
  Dim s     As Integer
  Dim t     As Integer
  
  ClearMap
  
  ' Now, get data from file
  For i = 1 To NumStreets()
    Get #1, i, Street(i)
  Next i
  
  ' Convert city to number
  For i = 1 To NumCities()
    Get #2, i, GetCity
    Test = UCase(Trim(GetCity))
    If UCase(Trim(Source)) = Test Then s = i
    If UCase(Trim(Target)) = Test Then t = i
    If (s <> 0) And (t <> 0) Then Exit For
  Next
  If (s = 0) Or (t = 0) Then
    MsgBox "*****  Location(s) not found **** " + IIf(s = 0, Chr(13) _
           + Source, "") + IIf(t = 0, Chr(13) + Target, "")
    Exit Sub
  End If
  Initialize s, t
  
  '* Begin Shortest Path function
  Do
    If A_Show Then
      tt = ""
      For ii = 1 To NumNodes
        tt = tt + Format(ii, "00") & " " & Format(Distance(ii), _
             "00000") & "  " & Format(Path(ii), "00") & "  " & _
             Format(Included(ii), "00") + Chr(13)
      Next
      If MsgBox(tt, 1) = vbCancel Then Exit Sub
    End If
    j = MinNode()
    Included(j) = True
    For i = 1 To NumNodes
      If Included(i) = False Then
        ' Prevent overflowing on any system
        If (Weight(j, i) < Infinity) And _
           (Distance(j) < Infinity) Then
          If Distance(j) + Weight(j, i) < Distance(i) Then
            Distance(i) = Distance(j) + Weight(j, i)
              Path(i) = j
          End If
        End If
      End If
    Next
  Loop Until AllIncluded()
  j = t
  sPath = Trim(Str(j))
  
  Do
    i = Path(j)
    sPath = Trim(Str(i)) + ", " + sPath
    j = i
    If i = -1 Then
      MsgBox "No Path Possible"
      Exit Sub
    End If
  Loop Until (i = Path(i))
  LastPath = sPath
  
  List1.Clear
  List1.AddItem "The Shortest Path:"
  bFirst = True
  Location = InStr(sPath, ",")
  Do While Location > 0
    t = Val(Mid(sPath, 1, Location - 1))
    Get #2, t, GetCity
    If bFirst Then
      List1.AddItem "<< " & GetCity
      AddBlinkingPoint ReadCoords(t), RGB(0, 255, 0), RGB(0, 200, 0), True
    Else
      List1.AddItem "-o- " & GetCity
      AddBlinkingPoint ReadCoords(t), vbYellow
    End If
    sPath = Mid(sPath, Location + 1)
    Location = InStr(sPath, ",")
    bFirst = False
  Loop
  t = Val(sPath)
  Get #2, t, GetCity
  List1.AddItem ">> " & GetCity
  Mileage.Caption = Distance(t)
  AddBlinkingPoint ReadCoords(t), RGB(255, 0, 0), RGB(150, 0, 0), True
  Source.SetFocus
End Sub


Private Sub cmdClose_Click()
  Me.Hide
  LastPath = ""
End Sub


Private Sub cmdShowMap_Click()
  If Form4.Visible Then
    Form4.Hide
  Else
    Form4.Show
  End If
  Timer1_Timer
End Sub


Private Sub cmdSrcMenu_Click()
  Source_DblClick
End Sub


Private Sub Form_Load()
  ClearMap
  If (NumCities() < 1) Or (NumStreets < 1) Then
    Unload Me
    Exit Sub
  End If
  NumNodes = NumCities()
  Infinity = 32767
  ReDim Street(1 To NumStreets()) As StreetType
  ReDim Distance(1 To NumNodes) As Integer
  ReDim Path(1 To NumNodes) As Integer
  ReDim Included(1 To NumNodes) As Integer
End Sub


Private Sub Initialize(ByVal s As Integer, ByVal t As Integer)
  Dim Wgt As Integer
  Dim i   As Integer
  For i = 1 To NumNodes
    Wgt = Weight(s, i)
    Distance(i) = Wgt
    Included(i) = (s = i)
    If Wgt >= 0 And Wgt <> Infinity Then
      Path(i) = s       ' source
    Else
      Path(i) = -1
    End If
  Next
End Sub


Private Function MinNode() As Integer
  Dim Temp  As Integer
  Dim Node  As Integer
  Dim Min   As Integer
  Dim i     As Integer
  '* Searches the unincluded nodes for the shortest distance and returns
  '* that node's number. Requires global arrays Included() and Distance()
  Min = Infinity     ' Max number for Integer
  For i = 1 To NumNodes
    If Included(i) = False Then        ' Not included
      Temp = i
      If Distance(i) < Min Then
        Min = Distance(i)
        Node = i
      End If
    End If
  Next
  If Min = Infinity Then
    Node = Temp
  End If
  MinNode = Node
End Function


Private Function Weight(ByVal lSource As Integer, ByVal j As Integer) As Integer
  Dim RetVal  As Integer
  Dim i       As Integer
  '* Returns the weight, or distance, between the two nodes lSource and j.
  '* If lSource is the same as j, Weight returns 0. If the two nodes
  '* aren't connected by a path, Weight returns Infinity.
  RetVal = 0
  If j = lSource Then
    Weight = 0
    Exit Function
  End If
  For i = 1 To UBound(Street)
    If (Street(i).From = lSource And Street(i).To = j) Or _
       (Street(i).From = j And Street(i).To = lSource) Then
      RetVal = Street(i).Wgt
      Exit For
    End If
  Next
  If RetVal = 0 Then RetVal = Infinity
  Weight = RetVal
End Function


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  LastPath = ""
  Me.Hide
  Cancel = True
End Sub


Private Sub mnuCity_Click(Index As Integer)
  SelCity = mnuCity(Index).Caption
End Sub


Private Sub Source_DblClick()
  SelCity = ""
  MakeCitiesMenu
  PopupMenu menuCities
  If Len(SelCity) Then Source = SelCity
End Sub


Private Sub Target_DblClick()
  SelCity = ""
  MakeCitiesMenu
  PopupMenu menuCities
  If Len(SelCity) Then Target = SelCity
End Sub


Private Sub Source_GotFocus()
  Source.SelStart = 0
  Source.SelLength = Len(Source)
End Sub


Private Sub Source_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
  Case 10, 13
    Target.SetFocus
    KeyAscii = 0
  End Select
End Sub


Private Sub Target_GotFocus()
  Target.SelStart = 0
  Target.SelLength = Len(Target)
End Sub


Private Sub Target_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
  Case 10, 13
    cmdFindPath_Click
    KeyAscii = 0
  End Select
End Sub


Private Sub Timer1_Timer()
  If Not Me.Visible Then Exit Sub
  If cmdShowMap.Tag <> Str(Abs(Form4.Visible)) Then
    If Form4.Visible Then
      cmdShowMap.Caption = "Hide Map"
    Else
      cmdShowMap.Caption = "Show Map"
    End If
    RePlaceCentered Me
    cmdShowMap.Tag = Str(Abs(Form4.Visible))
  End If
End Sub


Private Sub MakeCitiesMenu()
  Dim GetCity As String * 20
  Dim i       As Integer
  
  For i = 0 To NumCities() - 1
    If i > mnuCity.UBound Then
      Load mnuCity(i)
      mnuCity(i).Visible = True
    End If
    GetCity = String(20, " ")
    Get #2, i + 1, GetCity
    mnuCity(i).Caption = Trim(GetCity)
  Next i
  For i = i To mnuCity.UBound
    If i > 0 Then
      mnuCity(i).Visible = False
    End If
  Next i
End Sub
