VERSION 5.00
Begin VB.Form Form7 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form7"
   ClientHeight    =   9990
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20490
   LinkTopic       =   "Form7"
   PaletteMode     =   2  'Custom
   ScaleHeight     =   9990
   ScaleWidth      =   20490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Height          =   1095
      Left            =   12960
      MouseIcon       =   "Form7.frx":0000
      MousePointer    =   4  'Icon
      Picture         =   "Form7.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9360
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Height          =   1095
      Left            =   4080
      MouseIcon       =   "Form7.frx":0B82
      MousePointer    =   4  'Icon
      Picture         =   "Form7.frx":110C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9120
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   3135
      Index           =   0
      Left            =   2280
      Picture         =   "Form7.frx":1704
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   4860
   End
   Begin VB.Image Image3 
      Height          =   3375
      Index           =   1
      Left            =   10080
      Picture         =   "Form7.frx":3B5F
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   5895
   End
   Begin VB.Image Image4 
      Height          =   2580
      Left            =   1920
      Picture         =   "Form7.frx":552A
      Top             =   0
      Width           =   2490
   End
   Begin VB.Label Label3 
      BackColor       =   &H00404000&
      ForeColor       =   &H00800000&
      Height          =   27255
      Left            =   0
      TabIndex        =   6
      Top             =   -240
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   3135
      Index           =   1
      Left            =   2280
      Picture         =   "Form7.frx":6221
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   4860
   End
   Begin VB.Image Image2 
      Height          =   3105
      Index           =   2
      Left            =   2280
      Picture         =   "Form7.frx":7CE9
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   4845
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Top products"
      BeginProperty Font 
         Name            =   "Rosewood Std Regular"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   615
      Left            =   12000
      TabIndex        =   4
      Top             =   3600
      Width           =   3855
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Top Employee"
      BeginProperty Font 
         Name            =   "Rosewood Std Regular"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   735
      Left            =   3960
      TabIndex        =   3
      Top             =   3600
      Width           =   4215
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
      Left            =   1200
      MouseIcon       =   "Form7.frx":A5A0
      MousePointer    =   4  'Icon
      TabIndex        =   2
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Some glimpses of us at The Eagles..."
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
      Height          =   1215
      Left            =   5160
      TabIndex        =   1
      Top             =   360
      Width           =   8895
   End
   Begin VB.Label Label4 
      BackColor       =   &H00404000&
      Height          =   21735
      Left            =   19080
      TabIndex        =   0
      Top             =   0
      Width           =   1995
   End
   Begin VB.Image Image1 
      Height          =   2895
      Left            =   14640
      Picture         =   "Form7.frx":AB2A
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   3735
   End
   Begin VB.Image Image2 
      Height          =   3090
      Index           =   3
      Left            =   2280
      Picture         =   "Form7.frx":C91D
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   4875
   End
   Begin VB.Image Image3 
      Height          =   3375
      Index           =   2
      Left            =   10080
      Picture         =   "Form7.frx":FC2C
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   5895
   End
   Begin VB.Image Image3 
      Height          =   3375
      Index           =   3
      Left            =   10080
      Picture         =   "Form7.frx":11D4B
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   5865
   End
   Begin VB.Image Image3 
      Height          =   3375
      Index           =   4
      Left            =   10080
      Picture         =   "Form7.frx":13705
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   5820
   End
   Begin VB.Image Image3 
      Height          =   3195
      Index           =   5
      Left            =   10080
      Picture         =   "Form7.frx":15EFA
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   5820
   End
   Begin VB.Image Image5 
      Height          =   3135
      Left            =   2280
      Picture         =   "Form7.frx":18543
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   4845
   End
   Begin VB.Image Image6 
      Height          =   3255
      Left            =   10080
      Picture         =   "Form7.frx":1A836
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   5775
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
For i = 0 To 3
If Image2(i).Visible = True Then
Image2(i).Visible = False
Exit Sub
Else
flag = 1
End If
Next i
If flag = 1 Then
For i = 1 To 3
Image2(i).Visible = True
Next i
End If

End Sub

Private Sub Command2_Click()
For i = 1 To 5
If Image3(i).Visible = True Then
Image3(i).Visible = False
Exit Sub
Else
flag = 1
End If
Next i
If flag = 1 Then
For i = 1 To 5
Image3(i).Visible = True
Next i
End If
End Sub

Private Sub Label1_Click()
Form1.Show
Me.Hide

End Sub

