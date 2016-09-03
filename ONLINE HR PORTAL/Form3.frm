VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form3 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   9750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19635
   LinkTopic       =   "Form3"
   ScaleHeight     =   9750
   ScaleWidth      =   19635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   13680
      TabIndex        =   45
      Text            =   "Text2"
      Top             =   9360
      Width           =   2535
   End
   Begin VB.TextBox Text12 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   40
      Top             =   3960
      Width           =   1695
   End
   Begin VB.TextBox Text13 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   39
      Top             =   4680
      Width           =   1695
   End
   Begin VB.TextBox Text14 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   38
      Top             =   5640
      Width           =   1695
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   13560
      TabIndex        =   32
      Text            =   "--(Select)--"
      Top             =   3960
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Update"
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
      Left            =   7440
      MouseIcon       =   "Form3.frx":0000
      MousePointer    =   4  'Icon
      TabIndex        =   31
      Top             =   9240
      Width           =   1455
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3600
      TabIndex        =   24
      Top             =   4200
      Width           =   1695
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3600
      TabIndex        =   23
      Top             =   4920
      Width           =   1695
   End
   Begin VB.TextBox Text9 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   22
      Top             =   7320
      Width           =   1695
   End
   Begin VB.TextBox Text10 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      MultiLine       =   -1  'True
      TabIndex        =   21
      Top             =   8040
      Width           =   2895
   End
   Begin VB.TextBox Text11 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      MultiLine       =   -1  'True
      TabIndex        =   20
      Top             =   8640
      Width           =   2895
   End
   Begin VB.TextBox Text8 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   19
      Top             =   5640
      Width           =   1815
   End
   Begin VB.TextBox Text7 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
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
      Index           =   0
      Left            =   13560
      TabIndex        =   18
      Top             =   8400
      Width           =   2535
   End
   Begin VB.TextBox Text6 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
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
      Index           =   0
      Left            =   13560
      TabIndex        =   15
      Top             =   7680
      Width           =   2655
   End
   Begin VB.TextBox Text5 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
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
      Left            =   13560
      TabIndex        =   14
      Top             =   7080
      Width           =   2655
   End
   Begin VB.TextBox Text4 
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
      Left            =   13560
      TabIndex        =   13
      Top             =   6480
      Width           =   2655
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   120
      Top             =   8040
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
      Connect         =   $"Form3.frx":058A
      OLEDBString     =   $"Form3.frx":0630
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
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   13560
      TabIndex        =   11
      Text            =   "--(Select)--"
      Top             =   5640
      Width           =   2655
   End
   Begin VB.TextBox Text3 
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
      Left            =   13560
      TabIndex        =   9
      Top             =   4800
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
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
      Height          =   420
      Left            =   9720
      TabIndex        =   1
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Joining Date"
      Height          =   375
      Left            =   11160
      TabIndex        =   44
      Top             =   9360
      Width           =   1575
   End
   Begin VB.Label Label18 
      BackColor       =   &H00FFFFFF&
      Caption         =   "10th"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   5880
      TabIndex        =   43
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Label Label19 
      BackColor       =   &H00FFFFFF&
      Caption         =   "12th"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Left            =   5880
      TabIndex        =   42
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label Label20 
      BackColor       =   &H00FFFFFF&
      Caption         =   "University"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Left            =   5880
      TabIndex        =   41
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Label Label28 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Additional Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   375
      Left            =   11160
      TabIndex        =   37
      Top             =   3240
      Width           =   4935
   End
   Begin VB.Label Label22 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Contact Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   615
      Left            =   1440
      TabIndex        =   36
      Top             =   6360
      Width           =   2775
   End
   Begin VB.Label Label23 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   615
      Left            =   1440
      TabIndex        =   35
      Top             =   3360
      Width           =   3495
   End
   Begin VB.Label Label24 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Educational Qualification"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   615
      Left            =   5880
      TabIndex        =   34
      Top             =   3120
      Width           =   3615
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Welcome Admin"
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
      Left            =   7320
      TabIndex        =   33
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Left            =   1440
      TabIndex        =   30
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Date of Birth"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Left            =   1440
      TabIndex        =   29
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Caste"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Index           =   1
      Left            =   1440
      TabIndex        =   28
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Phone no"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   27
      Top             =   7320
      Width           =   1815
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Email id"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Index           =   1
      Left            =   1440
      TabIndex        =   26
      Top             =   7920
      Width           =   1815
   End
   Begin VB.Label Label17 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Left            =   1440
      TabIndex        =   25
      Top             =   8640
      Width           =   1935
   End
   Begin VB.Label Label16 
      BackColor       =   &H00004080&
      Height          =   15135
      Index           =   0
      Left            =   19320
      TabIndex        =   17
      Top             =   0
      Width           =   2100
   End
   Begin VB.Label Label15 
      BackColor       =   &H00004080&
      Height          =   19575
      Index           =   0
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   3000
      Left            =   1080
      Picture         =   "Form3.frx":06D6
      Top             =   360
      Width           =   3000
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Designation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Index           =   0
      Left            =   11160
      TabIndex        =   12
      Top             =   6360
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Earn leave"
      ForeColor       =   &H00000040&
      Height          =   615
      Left            =   11160
      TabIndex        =   10
      Top             =   8400
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Appeared in Interview?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Left            =   11160
      TabIndex        =   8
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Result"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   11160
      TabIndex        =   7
      Top             =   4920
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Department"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Left            =   11160
      TabIndex        =   6
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pay Scale"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   11160
      TabIndex        =   5
      Top             =   7080
      Width           =   1695
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Casual Leave"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Left            =   11160
      TabIndex        =   4
      Top             =   7680
      Width           =   1815
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Delete Account"
      Height          =   375
      Left            =   6360
      TabIndex        =   3
      Top             =   -240
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Log out"
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
      MouseIcon       =   "Form3.frx":216B
      MousePointer    =   4  'Icon
      TabIndex        =   2
      Top             =   0
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   450
      Left            =   11520
      MouseIcon       =   "Form3.frx":26F5
      MousePointer    =   4  'Icon
      Picture         =   "Form3.frx":2C7F
      Top             =   2040
      Width           =   315
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter the ID of the emplyee/job seeker"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   5040
      TabIndex        =   0
      Top             =   1920
      Width           =   4335
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As Connection
Dim rs1 As Recordset

Private Sub Command1_Click()
Dim temp As Recordset
Dim s1 As String
s1 = "select * from Details where id =" & Val(Trim(Text1.Text)) & ""
Set temp = New ADODB.Recordset
temp.Open s1, con, adOpenDynamic, adLockOptimistic

If Combo2.Text <> "--(Select)--" Then
temp![inter] = Combo2.Text
End If
If Trim(Text3.Text) <> "" Then
temp![res] = Text3.Text
End If
If Combo1.Text <> "--(Select)--" Then
temp![department] = Combo1.Text
End If
If Trim(Text6(0).Text) <> "" Then
temp![cl] = Val(Trim(Text6(0).Text))
End If
If Trim(Text7(0).Text) <> "" Then
temp![el] = Val(Trim(Text7(0).Text))
End If
If Trim(Text5.Text) <> "" Then
temp![payscale] = Val(Trim(Text5.Text))
End If
If Trim(Text4.Text) <> "" Then
temp![designation] = Trim(Text4.Text)
End If
If Trim(Text2.Text) <> "" Then
temp![jdate] = Trim(Text2.Text)
End If
temp.Update
temp.Close
MsgBox "Updated successfully"
End Sub

Private Sub Form_Load()
Set con = New ADODB.Connection
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:Hr.mdb;Persist Security Info=False"
Set rs1 = New ADODB.Recordset
rs1.Open "Details", con, adOpenDynamic, adLockOptimistic
Combo2.AddItem "Yes"
Combo2.AddItem "No"
Combo1.AddItem "Human Resource"
Combo1.AddItem "Bill and Establishment"
Combo1.AddItem "Finanace"
Combo1.AddItem "General Section"
Combo1.AddItem "Law Cell"
Combo1.AddItem "Accouns"
End Sub

Private Sub Image1_Click()
Dim temp As Recordset
Dim s1 As String
Dim l As Long
l = CLng(Text1.Text)
s1 = "select * from Details where id =  " & l & " "
Set temp = New ADODB.Recordset
temp.Open s1, con, adOpenDynamic, adLockReadOnly
If Not temp.EOF Then
If temp![Type] = 1 Then
Label23.Caption = "Job Seeker"
Else
Label23.Caption = "Employee"
End If
Text6(1).Text = temp![sname]
Text7(1).Text = temp![Dob]
If temp![Caste] = 0 Then
Text8.Text = "General"
Else
Text8.Text = "Others"
End If
Text9.Text = temp![Ph]
Text10.Text = temp![email]
Text11.Text = temp![Add]
Text12.Text = temp![ten]
Text13.Text = temp![twelve]
Text14.Text = temp![uni]
Else
MsgBox "No such record"
End If

End Sub

Private Sub Label2_Click()
Form1.Show
Me.Hide
End Sub

