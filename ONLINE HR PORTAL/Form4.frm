VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   10950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19320
   LinkTopic       =   "Form4"
   ScaleHeight     =   10950
   ScaleWidth      =   19320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   6600
      MultiLine       =   -1  'True
      TabIndex        =   9
      Text            =   "Form4.frx":0000
      Top             =   3360
      Visible         =   0   'False
      Width           =   10335
   End
   Begin VB.TextBox Text3 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   6720
      MultiLine       =   -1  'True
      TabIndex        =   8
      Text            =   "Form4.frx":00E7
      Top             =   3480
      Visible         =   0   'False
      Width           =   10215
   End
   Begin VB.TextBox Text2 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8295
      Left            =   6720
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "Form4.frx":0578
      Top             =   3480
      Visible         =   0   'False
      Width           =   10335
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7935
      Left            =   6720
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "Form4.frx":09EB
      Top             =   3360
      Visible         =   0   'False
      Width           =   8655
   End
   Begin VB.Image Image4 
      Height          =   3210
      Left            =   1920
      Picture         =   "Form4.frx":0C68
      Top             =   8520
      Width           =   3540
   End
   Begin VB.Image Image3 
      Height          =   2580
      Left            =   1680
      Picture         =   "Form4.frx":2DF5
      Top             =   480
      Width           =   2490
   End
   Begin VB.Label Label10 
      BackColor       =   &H00000080&
      Height          =   13335
      Left            =   18960
      TabIndex        =   11
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label Label12 
      BackColor       =   &H00000080&
      Height          =   13575
      Left            =   0
      MousePointer    =   4  'Icon
      TabIndex        =   10
      Top             =   0
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   8760
      Left            =   6600
      Picture         =   "Form4.frx":3AEC
      Top             =   3360
      Width           =   13800
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Welcome to the  Support System"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   1215
      Left            =   6840
      TabIndex        =   6
      Top             =   960
      Width           =   9255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "About Us"
      BeginProperty Font 
         Name            =   "Rosewood Std Regular"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   3240
      MouseIcon       =   "Form4.frx":1A70F
      MousePointer    =   4  'Icon
      TabIndex        =   4
      Top             =   7320
      Width           =   2655
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Contact Us"
      BeginProperty Font 
         Name            =   "Rosewood Std Regular"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   3240
      MouseIcon       =   "Form4.frx":1AC99
      MousePointer    =   4  'Icon
      TabIndex        =   3
      Top             =   6240
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Common Questions"
      BeginProperty Font 
         Name            =   "Rosewood Std Regular"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   3240
      MouseIcon       =   "Form4.frx":1B223
      MousePointer    =   4  'Icon
      TabIndex        =   2
      Top             =   4920
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "News and Events"
      BeginProperty Font 
         Name            =   "Rosewood Std Regular"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   3240
      MouseIcon       =   "Form4.frx":1B7AD
      MousePointer    =   4  'Icon
      TabIndex        =   1
      Top             =   3600
      Width           =   2655
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
      MouseIcon       =   "Form4.frx":1BD37
      MousePointer    =   4  'Icon
      TabIndex        =   0
      Top             =   0
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   2400
      Left            =   16320
      Picture         =   "Form4.frx":1C2C1
      Top             =   360
      Width           =   2400
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()
Form1.Show
Me.Hide
End Sub

Private Sub Label2_Click()
Image2.Visible = False
Text2.Visible = False
Text3.Visible = False
Text4.Visible = False
If Text1.Visible = False Then
Text1.Visible = True
Else
Image2.Visible = False
End If
End Sub

Private Sub Label4_Click()
Image2.Visible = False
Text2.Visible = False
Text1.Visible = False
Text3.Visible = False
If Text4.Visible = False Then
Text4.Visible = True
Else
Image2.Visible = False
End If
End Sub

Private Sub Label3_Click()
Image2.Visible = False
Text2.Visible = False
Text4.Visible = False
Text1.Visible = False
If Text3.Visible = False Then
Text3.Visible = True
Else
Image2.Visible = False
End If
End Sub

Private Sub Label5_Click()
Image2.Visible = False
Text1.Visible = False
Text3.Visible = False
Text4.Visible = False
If Text2.Visible = False Then
Text2.Visible = True
Else
Image2.Visible = False
End If
End Sub

