VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form6 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Profile"
   ClientHeight    =   10095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20250
   DrawStyle       =   2  'Dot
   LinkTopic       =   "Form6"
   ScaleHeight     =   10095
   ScaleWidth      =   20250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog c 
      Left            =   8040
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Post resume"
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
      Left            =   9600
      MouseIcon       =   "Form6.frx":0000
      MousePointer    =   4  'Icon
      TabIndex        =   51
      Top             =   9720
      Width           =   1575
   End
   Begin VB.TextBox Text18 
      Appearance      =   0  'Flat
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
      Left            =   14400
      TabIndex        =   50
      Text            =   "Text18"
      Top             =   8520
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Submit"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      MouseIcon       =   "Form6.frx":058A
      MousePointer    =   4  'Icon
      TabIndex        =   47
      Top             =   9720
      Width           =   1935
   End
   Begin VB.TextBox Text17 
      Appearance      =   0  'Flat
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
      Left            =   14400
      TabIndex        =   45
      Text            =   "Text17"
      Top             =   6360
      Width           =   2775
   End
   Begin VB.TextBox Text16 
      Appearance      =   0  'Flat
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
      Left            =   14400
      TabIndex        =   43
      Text            =   "Text16"
      Top             =   7680
      Width           =   2655
   End
   Begin VB.TextBox Text15 
      Appearance      =   0  'Flat
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
      Left            =   14400
      TabIndex        =   42
      Text            =   "Text15"
      Top             =   4920
      Width           =   2535
   End
   Begin VB.TextBox Text14 
      Appearance      =   0  'Flat
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
      Left            =   8880
      TabIndex        =   38
      Text            =   "Text14"
      Top             =   5640
      Width           =   1695
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
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
      TabIndex        =   33
      Text            =   "Text8"
      Top             =   5640
      Width           =   1815
   End
   Begin VB.TextBox Text13 
      Appearance      =   0  'Flat
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
      Left            =   8880
      TabIndex        =   32
      Text            =   "Text13"
      Top             =   4680
      Width           =   1695
   End
   Begin VB.TextBox Text12 
      Appearance      =   0  'Flat
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
      Left            =   8880
      TabIndex        =   31
      Text            =   "Text12"
      Top             =   3960
      Width           =   1695
   End
   Begin VB.TextBox Text11 
      Appearance      =   0  'Flat
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
      Height          =   495
      Left            =   3720
      TabIndex        =   30
      Text            =   "Text11"
      Top             =   8640
      Width           =   3015
   End
   Begin VB.TextBox Text10 
      Appearance      =   0  'Flat
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
      Height          =   495
      Left            =   3720
      TabIndex        =   29
      Text            =   "Text10"
      Top             =   7800
      Width           =   3015
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
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
      TabIndex        =   28
      Text            =   "Text9"
      Top             =   7200
      Width           =   1695
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
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
      TabIndex        =   27
      Text            =   "Text7"
      Top             =   4800
      Width           =   1695
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
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
      TabIndex        =   26
      Text            =   "Text6"
      Top             =   4080
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Height          =   2415
      Left            =   13440
      ScaleHeight     =   2355
      ScaleWidth      =   2475
      TabIndex        =   15
      Top             =   360
      Width           =   2535
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   14280
      Top             =   1440
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      Connect         =   $"Form6.frx":0B14
      OLEDBString     =   $"Form6.frx":0BDC
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
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
      Left            =   14400
      TabIndex        =   4
      Text            =   "Text5"
      Top             =   6960
      Width           =   2535
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
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
      Left            =   14400
      TabIndex        =   3
      Text            =   "Text4"
      Top             =   4080
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
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
      Left            =   14400
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   5520
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
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
      Left            =   14400
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   4920
      Width           =   2535
   End
   Begin VB.TextBox Text1 
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
      Left            =   14400
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   4080
      Width           =   2535
   End
   Begin VB.Label Label30 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Joining Date"
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
      Left            =   12000
      TabIndex        =   49
      Top             =   8520
      Width           =   1935
   End
   Begin VB.Label Label29 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Details of your interview is e-mailed to you.Please feel free to contact us .Thank you..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   1095
      Left            =   12000
      TabIndex        =   48
      Top             =   6240
      Width           =   4815
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
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   11880
      TabIndex        =   46
      Top             =   3120
      Width           =   4935
   End
   Begin VB.Label Label27 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Joining Date"
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
      Left            =   12120
      TabIndex        =   44
      Top             =   6360
      Width           =   1695
   End
   Begin VB.Label Label26 
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
      Left            =   12000
      TabIndex        =   41
      Top             =   7680
      Width           =   1935
   End
   Begin VB.Label Label25 
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
      Height          =   495
      Left            =   12000
      TabIndex        =   40
      Top             =   4800
      Width           =   1935
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
      Left            =   6600
      TabIndex        =   39
      Top             =   3000
      Width           =   3615
   End
   Begin VB.Label Label23 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label23"
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
      Left            =   1560
      TabIndex        =   37
      Top             =   3000
      Width           =   3495
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
      Left            =   1560
      TabIndex        =   36
      Top             =   6360
      Width           =   4215
   End
   Begin VB.Image Image1 
      Height          =   2460
      Left            =   1200
      Picture         =   "Form6.frx":0CA4
      Stretch         =   -1  'True
      Top             =   360
      Width           =   2955
   End
   Begin VB.Label Label21 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label21"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   9240
      TabIndex        =   35
      Top             =   600
      Width           =   3015
   End
   Begin VB.Label Label20 
      BackColor       =   &H00FFFFFF&
      Caption         =   "University"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Left            =   6600
      TabIndex        =   34
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Label Label19 
      BackColor       =   &H00FFFFFF&
      Caption         =   "12th"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Left            =   6600
      TabIndex        =   25
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label Label18 
      BackColor       =   &H00FFFFFF&
      Caption         =   "10th"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   6600
      TabIndex        =   24
      Top             =   4080
      Width           =   1935
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
      Left            =   1560
      TabIndex        =   23
      Top             =   8640
      Width           =   1935
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Email Id "
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
      Left            =   1560
      TabIndex        =   22
      Top             =   7800
      Width           =   1815
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Phone no."
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
      Left            =   1560
      TabIndex        =   21
      Top             =   7200
      Width           =   1815
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
      Left            =   1560
      TabIndex        =   20
      Top             =   5640
      Width           =   1815
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
      Left            =   1560
      TabIndex        =   19
      Top             =   4800
      Width           =   1695
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
      Left            =   1560
      TabIndex        =   18
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Welcome "
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6720
      TabIndex        =   17
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label Label10 
      BackColor       =   &H00004080&
      Height          =   25095
      Left            =   19200
      TabIndex        =   16
      Top             =   360
      Width           =   1860
   End
   Begin VB.Label Label1 
      BackColor       =   &H00004080&
      Height          =   16695
      Index           =   1
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Delete Account"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   5040
      MouseIcon       =   "Form6.frx":2A0D
      MousePointer    =   4  'Icon
      TabIndex        =   13
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Change your account information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2040
      MouseIcon       =   "Form6.frx":2F97
      MousePointer    =   4  'Icon
      TabIndex        =   12
      Top             =   0
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Log out"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1200
      MouseIcon       =   "Form6.frx":3521
      MousePointer    =   4  'Icon
      TabIndex        =   11
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Earn Leave"
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
      Left            =   12240
      TabIndex        =   10
      Top             =   6960
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
      Height          =   495
      Left            =   12000
      TabIndex        =   9
      Top             =   4080
      Width           =   1695
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
      Left            =   12000
      TabIndex        =   8
      Top             =   5520
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
      Left            =   12000
      TabIndex        =   7
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Appeared in interview?"
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
      Left            =   12000
      TabIndex        =   6
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Height          =   16695
      Index           =   0
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As Connection
Dim rs1 As Recordset

Private Sub Command1_Click()
Dim temp As Recordset
Dim s1 As String
s1 = "select * from Details where Uname ='" + Trim(Form1.Text1.Text) + "'"
Set temp = New ADODB.Recordset
temp.Open s1, con, adOpenDynamic, adLockOptimistic
temp![sname] = Text6.Text
temp![Dob] = Format(CDate(Text7.Text), "dd/mm/yyyy")
If Text8.Text = "General" Then
temp![Caste] = 0
Else
temp![Caste] = -1
End If
 temp![Ph] = Text9.Text
 temp![email] = Text10.Text
 temp![Add] = Text11.Text
 temp![ten] = Text12.Text
 temp![twelve] = Text13.Text
 temp![uni] = Text14.Text

temp.Update
temp.Close
MsgBox "Record updated"
Text6.Enabled = False
Text7.Enabled = False
Text8.Enabled = False
Text9.Enabled = False
Text10.Enabled = False
Text11.Enabled = False
Text12.Enabled = False
Text13.Enabled = False
Text14.Enabled = False
Command1.Enabled = False
End Sub

Private Sub Command2_Click()
Dim s1 As String

c.ShowOpen

MsgBox (c.FileName)
End Sub

Private Sub Form_Activate()
Dim temp As Recordset
Dim s1 As String

s1 = "select * from Details where Uname ='" + Trim(Form1.Text1.Text) + "'"
Set temp = New ADODB.Recordset
temp.Open s1, con, adOpenDynamic, adLockReadOnly
Text6.Text = temp![sname]
Text7.Text = temp![Dob]
If temp![Caste] = 0 Then
Text8.Text = "General"
Else
Text8.Text = "Others"
End If
Text9.Text = temp![Ph]
Text10.Text = temp![email]
Text11.Text = temp![Add]
Text1.Text = temp![inter]
Text2.Text = temp![res]
Text3.Text = temp![department]
Text4.Text = temp![payscale]
Text5.Text = temp![el]
Text16.Text = temp![cl]
Text15.Text = temp![designation]
Text17.Text = temp![jdate]
Text12.Text = temp![ten]
Text13.Text = temp![twelve]
Text14.Text = temp![uni]
Text18.Text = temp![jdate]
If temp![Image] <> "" Then
Picture1.Picture = LoadPicture(temp![Image])
End If
Label21.Caption = Form1.Text1.Text

If Form1.Option1.Value = True Then
Label23.Caption = "Job Seeker"
Label8.Visible = False
Text4.Visible = False
Label25.Visible = False
Text15.Visible = False
Label27.Visible = False
Text3.Visible = False
Label9.Visible = False
Text17.Visible = False
Label7.Visible = False
Text5.Visible = False
Label26.Visible = False
Text16.Visible = False
Label30.Visible = False
Text18.Visible = False
Else
Label23.Caption = "Employee"
Text1.Visible = False
Text2.Visible = False
Label5.Visible = False
Label6.Visible = False
Label29.Visible = False
End If
temp.Close
End Sub

Private Sub Image2_Click()

End Sub

Private Sub Form_Load()
Set con = New ADODB.Connection
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:Hr.mdb;Persist Security Info=False"
Set rs1 = New ADODB.Recordset
rs1.Open "Details", con, adOpenDynamic, adLockReadOnly
End Sub

Private Sub Label11_Click()

Dim flag As Boolean
Dim st1 As String
Dim st2 As String
Dim temp As ADODB.Recordset
Dim reply




st1 = "select * from Details where Uname='" + Trim(Form1.Text1.Text) + "'"
Set temp = New ADODB.Recordset
temp.Open st1, con, adOpenDynamic, adLockOptimistic
If Not temp.EOF Then
    st2 = "delete * from Details where Uname='" + Trim(Form1.Text1.Text) + "'"
    Set temp = New ADODB.Recordset
    reply = MsgBox("Are you sure?", vbYesNo + vbExclamation, "Delete")
    If reply = vbYes Then
        temp.Open st2, con, adOpenDynamic, adLockOptimistic
        x = MsgBox("Record deleted", vbOKOnly + vbInformation, "Message")
        Form1.Show
        Me.Hide
    Else
        MsgBox "Account not deleted"
        
    End If
Else
    MsgBox "No such account exists."
    
    Exit Sub
 End If
Form1.Text2.Text = ""
Form1.Text1.Text = ""
Form1.Option1.Value = False
Form1.Option2.Value = False
End Sub

Private Sub Label2_Click()
Form1.Show
MsgBox "Successfully logged out"
Form1.Text2.Text = ""
Form1.Text1.Text = ""
Form1.Option1.Value = False
Form1.Option2.Value = False
Unload Me

End Sub

Private Sub Label3_Click()
Text6.Enabled = True
Text7.Enabled = True
Text8.Enabled = True
Text9.Enabled = True
Text10.Enabled = True
Text11.Enabled = True
Text12.Enabled = True
Text13.Enabled = True
Text14.Enabled = True

Text6.SetFocus
Command1.Enabled = True
End Sub

Private Sub Label31_Click()

End Sub
