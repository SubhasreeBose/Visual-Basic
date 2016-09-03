VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form5 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form5"
   ClientHeight    =   11445
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19830
   ForeColor       =   &H00800000&
   LinkTopic       =   "Form5"
   ScaleHeight     =   11445
   ScaleWidth      =   19830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      DownPicture     =   "Form5.frx":0000
      Height          =   495
      Left            =   12480
      Picture         =   "Form5.frx":1F19
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   9000
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      DownPicture     =   "Form5.frx":277B
      Height          =   495
      Left            =   12480
      Picture         =   "Form5.frx":4694
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7920
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      DownPicture     =   "Form5.frx":4EF6
      Height          =   495
      Left            =   12480
      Picture         =   "Form5.frx":6E0F
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      DownPicture     =   "Form5.frx":7671
      Height          =   495
      Left            =   12480
      Picture         =   "Form5.frx":958A
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5880
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   7800
      TabIndex        =   10
      Top             =   10800
      Visible         =   0   'False
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Max             =   200
   End
   Begin MSComDlg.CommonDialog c 
      Left            =   1920
      Top             =   7080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   1920
      Top             =   4440
   End
   Begin VB.CommandButton Command1 
      DownPicture     =   "Form5.frx":9DEC
      Height          =   495
      Left            =   12480
      Picture         =   "Form5.frx":BD05
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   2580
      Left            =   2280
      Picture         =   "Form5.frx":C567
      Top             =   -120
      Width           =   2490
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Download  list of  selected interviewee."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   3120
      TabIndex        =   13
      Top             =   9000
      Width           =   8055
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      Caption         =   "All files finished downloading"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   1680
      TabIndex        =   12
      Top             =   10560
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Download Complete."
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   1680
      TabIndex        =   11
      Top             =   10200
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   2895
      Left            =   15720
      Picture         =   "Form5.frx":D25E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2655
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Download notice regarding vaccancies."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   3120
      TabIndex        =   8
      Top             =   5640
      Width           =   8055
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Download notice regarding internship programmes."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   3120
      TabIndex        =   7
      Top             =   6720
      Width           =   8055
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Download notice regarding bonus."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   3120
      TabIndex        =   6
      Top             =   7800
      Width           =   8055
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Download job / interview criteria. "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   3120
      TabIndex        =   5
      Top             =   4560
      Width           =   8055
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"Form5.frx":EB87
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   1215
      Left            =   2760
      TabIndex        =   4
      Top             =   2640
      Width           =   13095
   End
   Begin VB.Label Label4 
      BackColor       =   &H00004000&
      Height          =   21735
      Left            =   18480
      TabIndex        =   3
      Top             =   -120
      Width           =   1995
   End
   Begin VB.Label Label3 
      BackColor       =   &H00004000&
      Height          =   27255
      Left            =   -120
      TabIndex        =   2
      Top             =   -960
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Search and Find job offers at The Eagles"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   1335
      Left            =   5160
      TabIndex        =   1
      Top             =   120
      Width           =   8895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Home"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   1560
      MouseIcon       =   "Form5.frx":ECE6
      MousePointer    =   4  'Icon
      TabIndex        =   0
      Top             =   0
      Width           =   2295
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Timer1.Interval = 5
ProgressBar1.Value = 0
ProgressBar1.Max = 200
Label10.Visible = False
Label11.Visible = False
Timer1.Enabled = True
ProgressBar1.Visible = True
End Sub



Private Sub Command2_Click()
Timer1.Interval = 5
ProgressBar1.Value = 0
ProgressBar1.Max = 200
Label10.Visible = False
Label11.Visible = False
Timer1.Enabled = True
ProgressBar1.Visible = True
End Sub

Private Sub Command3_Click()
Timer1.Interval = 5
ProgressBar1.Value = 0
ProgressBar1.Max = 200
Label10.Visible = False
Label11.Visible = False
Timer1.Enabled = True
ProgressBar1.Visible = True

End Sub

Private Sub Command4_Click()
Timer1.Interval = 5
ProgressBar1.Value = 0
ProgressBar1.Max = 200
Label10.Visible = False
Label11.Visible = False
Timer1.Enabled = True
ProgressBar1.Visible = True
End Sub

Private Sub Command5_Click()
Timer1.Interval = 5
ProgressBar1.Value = 0
ProgressBar1.Max = 200
Label10.Visible = False
Label11.Visible = False
Timer1.Enabled = True
ProgressBar1.Visible = True
End Sub

Private Sub Form_Load()
''If Form1.List1.ListIndex = 0 Then

''ElseIf Form1.List1.ListIndex = 1 Then
''Text2.Visible = True
''Else
''Text3.Visible = True
''End If
Label10.Visible = False
Label11.Visible = False
ProgressBar1.Visible = False
Timer1.Enabled = False
End Sub

Private Sub Label1_Click()
Form1.Show
Me.Hide

End Sub

Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 1
If ProgressBar1.Value = ProgressBar1.Max Then
Label10.Visible = True
Label11.Visible = True
ProgressBar1.Visible = False
Timer1.Enabled = False
End If

End Sub
