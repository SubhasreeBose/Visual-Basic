VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Home"
   ClientHeight    =   10950
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   16395
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   16395
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1920
      Top             =   2160
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   8400
      Top             =   3840
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1296
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"Form1.frx":0000
      OLEDBString     =   $"Form1.frx":00C8
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   480
      Left            =   3720
      MouseIcon       =   "Form1.frx":0190
      MousePointer    =   4  'Icon
      TabIndex        =   11
      Top             =   3120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Log in"
      BeginProperty Font 
         Name            =   "Niagara Engraved"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   3975
      Left            =   1440
      TabIndex        =   2
      Top             =   4200
      Visible         =   0   'False
      Width           =   4575
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Caption         =   "X"
         Height          =   375
         Left            =   4200
         MaskColor       =   &H000000FF&
         MouseIcon       =   "Form1.frx":071A
         TabIndex        =   19
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00000000&
         Caption         =   "FORGOT PASSWORD?"
         BeginProperty Font 
            Name            =   "Modern No. 20"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2400
         MouseIcon       =   "Form1.frx":0CA4
         MousePointer    =   4  'Icon
         TabIndex        =   10
         Top             =   3000
         Width           =   1575
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00000000&
         Caption         =   "SUBMIT"
         BeginProperty Font 
            Name            =   "Modern No. 20"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2400
         MouseIcon       =   "Form1.frx":122E
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   2280
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "EMPLOYEE"
         ForeColor       =   &H00400040&
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   3000
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "JOB SEEKER"
         ForeColor       =   &H00400040&
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   2280
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   2160
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   1320
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   495
         Left            =   2160
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Password :-"
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Username :-"
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   495
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Type your full  name"
         Top             =   600
         Width           =   1815
      End
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFFFF&
      Caption         =   "All rights reserved.Copyright @ subhasree10.7"
      Height          =   495
      Left            =   9720
      TabIndex        =   24
      Top             =   10920
      Width           =   5295
   End
   Begin VB.Image Image4 
      Height          =   405
      Left            =   2400
      MouseIcon       =   "Form1.frx":17B8
      MousePointer    =   4  'Icon
      Picture         =   "Form1.frx":1D42
      Stretch         =   -1  'True
      Top             =   9720
      Width           =   540
   End
   Begin VB.Image Image3 
      Height          =   735
      Left            =   1320
      MouseIcon       =   "Form1.frx":2F15
      MousePointer    =   4  'Icon
      Picture         =   "Form1.frx":349F
      Stretch         =   -1  'True
      Top             =   9600
      Width           =   720
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Register here to apply for job"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   495
      Index           =   4
      Left            =   13440
      MouseIcon       =   "Form1.frx":3E0B
      MousePointer    =   4  'Icon
      TabIndex        =   23
      Top             =   2760
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Find job offers and important notices at employee corner"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   495
      Index           =   2
      Left            =   13440
      TabIndex        =   22
      Top             =   2880
      Visible         =   0   'False
      Width           =   5655
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Support system will guide you "
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   495
      Index           =   1
      Left            =   13440
      TabIndex        =   21
      Top             =   2880
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Log in if you are already registered"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   495
      Index           =   0
      Left            =   13440
      TabIndex        =   20
      Top             =   2880
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Label Label12 
      BackColor       =   &H00400040&
      Height          =   13575
      Left            =   0
      MousePointer    =   4  'Icon
      TabIndex        =   18
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label11 
      Caption         =   "Label11"
      Height          =   15
      Left            =   0
      TabIndex        =   17
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label10 
      BackColor       =   &H00400040&
      Height          =   13335
      Left            =   19320
      TabIndex        =   16
      Top             =   -600
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   2580
      Left            =   1800
      Picture         =   "Form1.frx":4395
      Top             =   -120
      Width           =   2490
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Support"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   5760
      MouseIcon       =   "Form1.frx":508C
      MousePointer    =   4  'Icon
      TabIndex        =   15
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Employee Corner"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   3720
      MouseIcon       =   "Form1.frx":5616
      MousePointer    =   4  'Icon
      TabIndex        =   14
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Register"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   2720
      MouseIcon       =   "Form1.frx":5BA0
      MousePointer    =   4  'Icon
      TabIndex        =   13
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Log in"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   1800
      MouseIcon       =   "Form1.frx":612A
      MousePointer    =   4  'Icon
      TabIndex        =   12
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Welcome to the HR Department"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8760
      TabIndex        =   1
      Top             =   1080
      Width           =   12615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "The Eagles"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9000
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
   Begin VB.Image Image2 
      Height          =   6960
      Left            =   3600
      Picture         =   "Form1.frx":66B4
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   11850
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As Connection
Dim rs1 As Recordset
Dim a As Integer
Dim f As Integer

Private Sub Command1_Click()
Text1.Text = ""
Text2.Text = ""
Option1.Value = False
Option2.Value = False
Frame1.Visible = False
End Sub






Private Sub Command5_Click()
Dim s1, s2 As String
Dim temp As Recordset
Dim temp1 As Recordset
Set temp = New ADODB.Recordset
Set temp1 = New ADODB.Recordset


s1 = "select * from Details where ID=1"
temp1.Open s1, con, adOpenDynamic, adLockOptimistic
If Text1.Text = "" Or Text2.Text = "" Then
MsgBox "Fill up all the fields"
Else
If temp1![sname] = Trim(Text1.Text) And temp1![Password] = Trim(Text2.Text) Then
    Form3.Show
Else
    If Option1.Value = False And Option2.Value = False Then
    MsgBox "Fill up all the fields"
    Else
    s2 = "select * from Details where Uname ='" + (Trim(Text1.Text)) + "'"
    temp.Open s2, con, adOpenDynamic, adLockOptimistic
    If Not temp.EOF Then
    If Option1.Value = False Then
    l = 0
    Else
    l = 1
    End If
    
        If temp![Password] = Trim(Text2.Text) And l = Abs(temp![Type]) Then
        Form6.Show
        Else
        MsgBox "Wrong password"
        Text1.Text = ""
Text2.Text = ""
Option1.Value = False
Option2.Value = False
        End If
    Else
    MsgBox "Invalid Username"
    Text1.Text = ""
Text2.Text = ""
Option1.Value = False
Option2.Value = False
    End If
End If
End If
End If

End Sub

Private Sub Command6_Click()
x = MsgBox("Contact us.Go to Support page.", vbOKOnly, "Message")
End Sub

Private Sub Form_Load()
Set con = New ADODB.Connection
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:Hr.mdb;Persist Security Info=False"
Set rs1 = New ADODB.Recordset
rs1.Open "Details", con, adOpenDynamic, adLockOptimistic
List1.AddItem "Offers/Notices"
List1.AddItem "Photo Gallery"


End Sub

Private Sub Image3_Click()
MsgBox ("Log in with Facebook to like our page.")
End Sub

Private Sub Image4_Click()
MsgBox ("Log in with g+ to like our page.")
End Sub

Private Sub Label5_Click()
List1.Visible = False
Frame1.Visible = True
End Sub

Private Sub Label6_Click()
List1.Visible = False
Form2.Show
Form2.WindowState = vbMaximized

Me.Hide
End Sub



Private Sub Label7_Click(Index As Integer)
If Index = 4 Then
Form2.Show
Form2.WindowState = vbMaximized
Me.Hide
ElseIf Index = 2 Then
Form5.Show
Form5.WindowState = vbMaximized
Me.Hide
ElseIf Index = 1 Then
Form4.Show
Form4.WindowState = vbMaximized
Me.Hide
Else
Frame1.Visible = True
End If


End Sub

Private Sub Label8_Click()
If List1.Visible = False Then
List1.Visible = True
Else
List1.Visible = False
End If

End Sub

Private Sub Label9_Click()
List1.Visible = False
Form4.Show
Form4.WindowState = vbMaximized
Me.Hide
End Sub

Private Sub List1_Click()
If List1.ListIndex = 0 Then
Form5.Show
Form5.WindowState = vbMaximized
Me.Hide
End If
If List1.ListIndex = 1 Then
Form7.Show
Form7.WindowState = vbMaximized
Me.Hide
End If
End Sub

Private Sub Text1_Click()
Text1.Text = ""
End Sub

Private Sub Timer1_Timer()
a = a + 1

If (a Mod 25 = 0) Then
 Label7(0).Visible = True
 Label7(1).Visible = False
 Label7(2).Visible = False
 Label7(4).Visible = False

End If
If (a Mod 50 = 0) Then
 Label7(1).Visible = True
 Label7(0).Visible = False
 Label7(2).Visible = False
 Label7(4).Visible = False

End If
If (a Mod 75 = 0) Then
 Label7(2).Visible = True
 Label7(1).Visible = False
 Label7(0).Visible = False
 Label7(4).Visible = False

End If
If (a Mod 100 = 0) Then
 Label7(4).Visible = True
 Label7(1).Visible = False
 Label7(2).Visible = False
 Label7(0).Visible = False

End If
End Sub
