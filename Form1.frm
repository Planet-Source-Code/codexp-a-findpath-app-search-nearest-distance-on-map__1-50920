VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Shortest Path Demonstration"
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   3255
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdCoordsEdit 
      Caption         =   "City && Map Location Editor"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2520
      Top             =   3960
   End
   Begin VB.CommandButton cmdShowMap 
      Caption         =   "Show Map"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   2760
      Width           =   2775
   End
   Begin VB.CommandButton cmdQUIT 
      Caption         =   "QUIT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   5160
      Width           =   2775
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Streets && Distance Editor"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   2775
   End
   Begin VB.CommandButton cmdFindPath 
      Caption         =   "Find Shortest Path"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   2775
   End
   Begin VB.Line linBorder 
      BorderColor     =   &H00808080&
      Index           =   5
      X1              =   3000
      X2              =   240
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Find Shortest Path Demo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   600
      TabIndex        =   7
      Top             =   150
      Width           =   2100
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Copyright (C)2003 by CodeXP [extd.]"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   3720
      Width           =   2715
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      Caption         =   "Copyright (C)1995 by The Cobb Group"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   3480
      Width           =   2775
   End
   Begin VB.Line linBorder 
      BorderColor     =   &H00808080&
      Index           =   4
      X1              =   240
      X2              =   3000
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line linBorder 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   3000
      X2              =   240
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line linBorder 
      BorderColor     =   &H00808080&
      Index           =   2
      X1              =   240
      X2              =   3000
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line linBorder 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   240
      X2              =   3000
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line linBorder 
      BorderColor     =   &H00808080&
      Index           =   3
      X1              =   240
      X2              =   3000
      Y1              =   5040
      Y2              =   5040
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdCoordsEdit_Click()
  Dim bExit As Boolean
  Me.Hide
  Form5.Show
  RePlaceCentered Form5
  Do Until bExit
    bExit = Not Form5.Visible
    DoEvents
  Loop
  ClearMap
  Me.Show
  RePlaceCentered Me
End Sub


Private Sub cmdEdit_Click()
  Dim bExit As Boolean
  Me.Hide
  Form3.Show
  RePlaceCentered Form3
  Do Until bExit
    bExit = Not Form3.Visible
    DoEvents
  Loop
  ClearMap
  Me.Show
  RePlaceCentered Me
End Sub


Private Sub cmdFindPath_Click()
  Dim bExit As Boolean
  Me.Hide
  Form2.Show
  RePlaceCentered Form2
  Do Until bExit
    bExit = Not Form2.Visible
    DoEvents
  Loop
  ClearMap
  Me.Show
  RePlaceCentered Me
End Sub


Private Sub cmdQUIT_Click()
  Unload Me
End Sub


Private Sub cmdShowMap_Click()
  If Form4.Visible Then
    Form4.Hide
  Else
    Form4.Show
  End If
  RePlaceCentered Me
  Timer1_Timer
End Sub


Private Sub Form_Load()
  LoadBorders
  MainInit
  RePlaceCentered Me
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  End
End Sub


Private Sub Timer1_Timer()
  If Not Me.Visible Then Exit Sub
  If cmdShowMap.Tag <> Str(Abs(Form4.Visible)) Then
    If Form4.Visible Then
      cmdShowMap.Caption = "Hide Map"
    Else
      cmdShowMap.Caption = "Show Map"
      RePlaceCentered Me
    End If
    cmdShowMap.Tag = Str(Abs(Form4.Visible))
  End If
End Sub


Private Sub LoadBorders()
  Dim i As Integer
  Dim j As Integer
  Dim n As Integer
  n = linBorder.UBound
  For i = 0 To n
    j = i + n + 1
    Load linBorder(j)
    linBorder(j).BorderColor = vbWhite
    linBorder(j).X1 = linBorder(i).X1
    linBorder(j).X2 = linBorder(i).X2
    linBorder(j).Y1 = linBorder(i).Y1 + Screen.TwipsPerPixelY
    linBorder(j).Y2 = linBorder(i).Y2 + Screen.TwipsPerPixelY
    linBorder(j).Visible = True
  Next i
End Sub
