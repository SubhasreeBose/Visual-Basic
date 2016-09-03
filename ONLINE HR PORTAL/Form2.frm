VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REGISTRATION"
   ClientHeight    =   10800
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   19560
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10800
   ScaleWidth      =   19560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "d/M/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13440
      TabIndex        =   33
      Top             =   3240
      Width           =   2295
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Job Seeker"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   3960
      TabIndex        =   32
      ToolTipText     =   "Employees donot check it"
      Top             =   5760
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "General"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3960
      TabIndex        =   31
      Top             =   3960
      Width           =   2295
   End
   Begin VB.TextBox Text11 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   30
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Educational Qualification"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3615
      Left            =   10800
      TabIndex        =   22
      Top             =   6240
      Width           =   6375
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2880
         TabIndex        =   28
         Top             =   2520
         Width           =   3375
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         TabIndex        =   26
         Top             =   1560
         Width           =   3375
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2880
         TabIndex        =   24
         Top             =   720
         Width           =   3375
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFFF&
         Caption         =   "University"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   120
         TabIndex        =   27
         Top             =   2520
         Width           =   2655
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFFF&
         Caption         =   "12th"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   120
         TabIndex        =   25
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         Caption         =   "10th"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Width           =   3375
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   8160
      Top             =   480
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   873
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
      Connect         =   $"Form2.frx":0000
      OLEDBString     =   $"Form2.frx":00C8
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
   Begin MSComDlg.CommonDialog cd 
      Left            =   10920
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Height          =   2055
      Left            =   12120
      ScaleHeight     =   1995
      ScaleWidth      =   2115
      TabIndex        =   18
      Top             =   360
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Submit"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      MouseIcon       =   "Form2.frx":0190
      MousePointer    =   4  'Icon
      TabIndex        =   16
      Top             =   9840
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8280
      MouseIcon       =   "Form2.frx":071A
      MousePointer    =   4  'Icon
      TabIndex        =   15
      Top             =   9840
      Width           =   1695
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   7200
      PasswordChar    =   "*"
      TabIndex        =   14
      Top             =   7680
      Width           =   2295
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   7200
      PasswordChar    =   "*"
      TabIndex        =   13
      Top             =   6840
      Width           =   2295
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13440
      TabIndex        =   10
      Top             =   4800
      Width           =   2295
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Others"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   6240
      TabIndex        =   9
      Top             =   3960
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13440
      TabIndex        =   3
      Top             =   3960
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   4920
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   1
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   3960
      TabIndex        =   29
      Top             =   3120
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   2580
      Left            =   1200
      Picture         =   "Form2.frx":0CA4
      Top             =   360
      Width           =   2490
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C00000&
      Caption         =   "Label12"
      Height          =   12735
      Left            =   19320
      TabIndex        =   21
      Top             =   -240
      Width           =   1500
   End
   Begin VB.Label Label11 
      BackColor       =   &H00800000&
      Height          =   11775
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Upload Photo of yours"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   12120
      TabIndex        =   19
      ToolTipText     =   "Upload a stamp size color photo"
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Confirm Password"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3960
      TabIndex        =   17
      Top             =   7680
      Width           =   2295
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3960
      TabIndex        =   12
      Top             =   6840
      Width           =   2295
   End
   Begin VB.Image Image2 
      Height          =   3300
      Left            =   15360
      Picture         =   "Form2.frx":199B
      Top             =   0
      Width           =   3435
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "New User Registration Form"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   6240
      TabIndex        =   11
      Top             =   240
      Width           =   4695
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "E-mail id"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   10560
      TabIndex        =   8
      Top             =   4920
      Width           =   2295
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Date Of Birth(d/m/yyyy)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   10560
      TabIndex        =   7
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Contact No."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   10560
      TabIndex        =   6
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   5040
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   2280
      Width           =   2295
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
      Left            =   1080
      MouseIcon       =   "Form2.frx":4D0B
      MousePointer    =   4  'Icon
      TabIndex        =   0
      Top             =   0
      Width           =   2295
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim con As Connection
Dim rs1 As Recordset

Private Sub Command1_Click()
Dim temp As Recordset
Set temp = New Recordset
Dim s1 As String
s1 = "select * from Details where Uname='" + Trim(Text11.Text) + "'"
If Trim(Text1.Text) = "" Or Trim(Text2.Text) = "" Or Trim(Text3.Text) = "" Or Trim(Text4.Text) = "" Or Trim(Text5.Text) = "" Or Trim(Text6.Text) = "" Or Trim(Text7.Text) = "" Or Trim(Text8.Text) = "" Or Trim(Text9.Text) = "" Or Trim(Text10.Text) = "" Or Trim(Text11.Text) = "" Then
MsgBox "Please fill up the blank fields"
ElseIf Option1.Value = False And Option2.Value = False Then
MsgBox "Please select option."
ElseIf cd.FileName = "" Then
MsgBox "Please upload picture"
Else
If Trim(Text5.Text) = Trim(Text7.Text) Then
temp.Open s1, con, adOpenDynamic, adLockReadOnly
If Not temp.EOF Then
MsgBox "Type any other username"
Text11.Text = ""
Exit Sub
End If
rs1.AddNew
rs1![sname] = Trim(Text1.Text)
rs1![Dob] = Format(CDate(Trim(Text4.Text)), "dd/mm/yyyy")
rs1![Caste] = Option1.Value
rs1![Type] = Check1.Value
rs1![Password] = (Trim(Text5.Text))
rs1![Add] = Trim(Text2.Text)
rs1![Uname] = Trim(Text11.Text)
rs1![ten] = Val(Trim(Text8.Text))
rs1![twelve] = Val(Trim(Text9.Text))
rs1![jdate] = "Not updated"
rs1![uni] = Val(Trim(Text10.Text))
rs1![Ph] = Val(Trim(Text3.Text))
rs1![email] = Trim(Text6.Text)
rs1![Image] = cd.FileName
rs1.Update
MsgBox "Records Updated", vbInformation + vbOKOnly
Else
MsgBox "Re enter your password"
End If
End If
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
End Sub

Private Sub Form_Load()
Set con = New ADODB.Connection
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:Hr.mdb;Persist Security Info=False"
Set rs1 = New ADODB.Recordset
rs1.Open "Details", con, adOpenDynamic, adLockOptimistic
Option1.Value = 0
Option2.Value = 0
End Sub

Private Sub Label1_Click()
Unload Me
Form1.Show

End Sub

Private Sub Label10_Click()
cd.ShowOpen
Picture1.Picture = LoadPicture(cd.FileName)


End Sub


