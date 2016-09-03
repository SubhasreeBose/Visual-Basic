VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Splashscreen 
   BackColor       =   &H00400040&
   BorderStyle     =   0  'None
   Caption         =   "Form7"
   ClientHeight    =   8505
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13860
   DrawStyle       =   5  'Transparent
   Icon            =   "Splashscreen.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   13860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   735
      Left            =   3480
      TabIndex        =   6
      Top             =   6360
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   1296
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   480
      Top             =   2400
   End
   Begin VB.Label Label10 
      BackColor       =   &H00400040&
      Caption         =   "All rights reserved.Copyright @ subhasree10.7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   10
      Top             =   7680
      Width           =   5295
   End
   Begin VB.Image Image1 
      Height          =   2580
      Left            =   6120
      Picture         =   "Splashscreen.frx":000C
      Top             =   3000
      Width           =   2490
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   3975
      Left            =   5280
      Shape           =   2  'Oval
      Top             =   2160
      Width           =   4095
   End
   Begin VB.Label Label9 
      BackColor       =   &H00400040&
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3600
      TabIndex        =   9
      Top             =   2520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label8 
      BackColor       =   &H00400040&
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2880
      TabIndex        =   8
      Top             =   3600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2520
      TabIndex        =   7
      Top             =   4920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackColor       =   &H00400040&
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   11400
      TabIndex        =   5
      Top             =   4920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label5 
      BackColor       =   &H00400040&
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   11040
      TabIndex        =   4
      Top             =   3480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H00400040&
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   9840
      TabIndex        =   3
      Top             =   2040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H00400040&
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   8520
      TabIndex        =   2
      Top             =   1200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00400040&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   6600
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00400040&
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   5040
      TabIndex        =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "Splashscreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer
Dim f As Integer

Private Sub Image2_Click()

End Sub

Private Sub Timer1_Timer()
a = a + 1
ProgressBar1.Value = ProgressBar1.Value + 1
If (a Mod 15 = 0 And f = 0) Then
 Label7.Visible = True
 
 f = 1
End If
If (a Mod 20 = 0 And f = 1) Then
 Label8.Visible = True
 f = 2
End If
If (a Mod 25 = 0 And f = 2) Then
 Label9.Visible = True
 f = 3
End If
If (a Mod 35 = 0 And f = 3) Then
 Label1.Visible = True
 
 f = 4
End If
If (a Mod 35 = 0 And f = 4) Then
 Label2.Visible = True
 
 f = 5
End If
If (a Mod 40 = 0 And f = 5) Then
 Label3.Visible = True
 
 f = 6
End If
If (a Mod 45 = 0 And f = 6) Then
 Label4.Visible = True
 
 f = 7
 End If
 If (a Mod 50 = 0 And f = 7) Then
 Label5.Visible = True
 
 f = 8
End If
If (a Mod 55 = 0 And f = 8) Then
 Label6.Visible = True
 
 f = 9
End If
If (a Mod 60 = 0 And f = 9) Then
 Label7.Visible = True
 
 f = 10
End If

If (ProgressBar1.Value = ProgressBar1.Max) Then
Form1.WindowState = vbMaximized
 Form1.Show
 Unload Me
End If
End Sub
