VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Form Template"
   ClientHeight    =   1590
   ClientLeft      =   7290
   ClientTop       =   3375
   ClientWidth     =   5745
   FillColor       =   &H00000040&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   5745
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FFFF&
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5445
      Width           =   1095
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Guest Lecturer"
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   2400
      TabIndex        =   10
      Top             =   840
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Lecturer"
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   840
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   7440
      Top             =   4920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Time.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Time.mdb;Persist Security Info=False"
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
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Register"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5445
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF0000&
      Caption         =   "Guest Lecturer"
      ForeColor       =   &H0000FFFF&
      Height          =   3375
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Visible         =   0   'False
      Width           =   5175
      Begin VB.ComboBox Combo1 
         Height          =   360
         Left            =   2040
         TabIndex        =   14
         Text            =   "Select"
         Top             =   1125
         Width           =   2655
      End
      Begin VB.ComboBox Combo3 
         Height          =   360
         Left            =   2040
         TabIndex        =   12
         Text            =   "Select"
         Top             =   2640
         Width           =   2655
      End
      Begin VB.ComboBox Combo2 
         Height          =   360
         Left            =   2040
         TabIndex        =   11
         Text            =   "Select"
         Top             =   1890
         Width           =   2655
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   8
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FF0000&
         Caption         =   "On Days"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FF0000&
         Caption         =   "Name"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF0000&
      Caption         =   "Teacher"
      ForeColor       =   &H0000FFFF&
      Height          =   3375
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Visible         =   0   'False
      Width           =   5175
      Begin VB.ComboBox Combo4 
         Height          =   360
         Left            =   2040
         TabIndex        =   13
         Text            =   "Select"
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   4
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF0000&
         Caption         =   "Off Day"
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Left            =   480
         TabIndex        =   3
         Top             =   1380
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF0000&
         Caption         =   "Name"
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   480
         TabIndex        =   2
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select Category:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   240
      TabIndex        =   16
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Dim res As ADODB.Recordset
Dim res1 As ADODB.Recordset


Private Sub Combo2_DropDown()
If Combo1.ListIndex <> -1 Then
Combo2.RemoveItem Combo1.ListIndex
Combo3.RemoveItem Combo1.ListIndex
End If
End Sub



Private Sub Combo3_DropDown()
If Combo2.ListIndex <> -1 Then
Combo3.RemoveItem Combo2.ListIndex
End If
End Sub


Private Sub Command1_Click()
Dim temp As ADODB.Recordset
If Frame1.Visible = True Then
    If Text1.Text = "" Or Combo4.Text = "Select" Then
        MsgBox "Fill Empty Fields!", vbCritical, "Field Empty"
    Else
        s = "select * from teachers where tname='" + Trim(Text1.Text) + "'"
        Set temp = New ADODB.Recordset
        temp.Open s, con, adOpenDynamic, adLockOptimistic
        If temp.EOF = True Then
             res.AddNew
             res![tname] = Trim(Text1.Text)
             res![off1] = Trim(Combo4.Text)

             res.Update
             MsgBox "Registered Successfully", vbInformation, "Registration"
        Else
            MsgBox "Already exists", vbCritical, "Lecturer"
        End If
     
    End If
 'con.Execute ("create table " & Text1.Text & "(Room varchar(255), xdate Date/Time varchar(255))")
End If
If Frame2.Visible = True Then
    If Text4.Text = "" Then
        MsgBox "Fill Empty Fields!", vbCritical, "Field Empty"
    Else
         s = "select * from invigilator where iname='" + Trim(Text4.Text) + "'"
        Set temp = New ADODB.Recordset
        temp.Open s, con, adOpenDynamic, adLockOptimistic
        If temp.EOF = True Then
            res1.AddNew
            res1![iname] = Trim(Text4.Text)
            res1![on1] = Trim(Combo1.Text)
            res1![on2] = Trim(Combo2.Text)
            res1![on3] = Trim(Combo3.Text)
            res1.Update
            MsgBox "Registered Successfully", vbInformation, "Registration"
        Else
            MsgBox "Already exists", vbCritical, "Guest Lecturer"
        End If
    End If
'con.Execute ("create table " & Text4.Text & "(Room varchar(255), xdate  varchar(255))")
End If


Combo1.Clear
Combo1.AddItem "Monday"
Combo1.AddItem "Tuesday"
Combo1.AddItem "Wednesday"
Combo1.AddItem "Thursday"
Combo1.AddItem "Friday"
Combo1.AddItem "Saturday"
Combo1.Text = "Select"

Combo2.Clear
Combo2.AddItem "Monday"
Combo2.AddItem "Tuesday"
Combo2.AddItem "Wednesday"
Combo2.AddItem "Thursday"
Combo2.AddItem "Friday"
Combo2.AddItem "Saturday"
Combo2.Text = "Select"

Combo3.Clear
Combo3.AddItem "Monday"
Combo3.AddItem "Tuesday"
Combo3.AddItem "Wednesday"
Combo3.AddItem "Thursday"
Combo3.AddItem "Friday"
Combo3.AddItem "Saturday"
Combo3.Text = "Select"

Combo4.Clear
Combo4.AddItem "Monday"
Combo4.AddItem "Tuesday"
Combo4.AddItem "Wednesday"
Combo4.AddItem "Thursday"
Combo4.AddItem "Friday"
Combo4.AddItem "Saturday"
Combo4.Text = "Select"

Text1.Text = ""

Text4.Text = ""

End Sub


Private Sub Command2_Click()
Combo1.Clear
Combo1.AddItem "Monday"
Combo1.AddItem "Tuesday"
Combo1.AddItem "Wednesday"
Combo1.AddItem "Thursday"
Combo1.AddItem "Friday"
Combo1.AddItem "Saturday"
Combo1.Text = "Select"

Combo2.Clear
Combo2.AddItem "Monday"
Combo2.AddItem "Tuesday"
Combo2.AddItem "Wednesday"
Combo2.AddItem "Thursday"
Combo2.AddItem "Friday"
Combo2.AddItem "Saturday"
Combo2.Text = "Select"

Combo3.Clear
Combo3.AddItem "Monday"
Combo3.AddItem "Tuesday"
Combo3.AddItem "Wednesday"
Combo3.AddItem "Thursday"
Combo3.AddItem "Friday"
Combo3.AddItem "Saturday"
Combo3.Text = "Select"

Combo4.Clear
Combo4.AddItem "Monday"
Combo4.AddItem "Tuesday"
Combo4.AddItem "Wednesday"
Combo4.AddItem "Thursday"
Combo4.AddItem "Friday"
Combo4.AddItem "Saturday"
Combo4.Text = "Select"

Text1.Text = ""
Text4.Text = ""
End Sub

Private Sub Form_Activate()
Option1.Value = False
Option2.Value = False


End Sub

Private Sub Form_Load()
Set con = New ADODB.Connection
Set res = New ADODB.Recordset
Set res1 = New ADODB.Recordset
con.Open " Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Time.mdb;Persist Security Info=False"
res.Open "teachers", con, adOpenDynamic, adLockOptimistic
res1.Open "invigilator", con, adOpenDynamic, adLockOptimistic
Combo1.AddItem "Monday"
Combo1.AddItem "Tuesday"
Combo1.AddItem "Wednesday"
Combo1.AddItem "Thursday"
Combo1.AddItem "Friday"
Combo1.AddItem "Saturday"

Combo2.AddItem "Monday"
Combo2.AddItem "Tuesday"
Combo2.AddItem "Wednesday"
Combo2.AddItem "Thursday"
Combo2.AddItem "Friday"
Combo2.AddItem "Saturday"

Combo3.AddItem "Monday"
Combo3.AddItem "Tuesday"
Combo3.AddItem "Wednesday"
Combo3.AddItem "Thursday"
Combo3.AddItem "Friday"
Combo3.AddItem "Saturday"

Combo4.AddItem "Monday"
Combo4.AddItem "Tuesday"
Combo4.AddItem "Wednesday"
Combo4.AddItem "Thursday"
Combo4.AddItem "Friday"
Combo4.AddItem "Saturday"

Command1.Enabled = True
End Sub

Private Sub Option1_Click()
Frame1.Visible = True
Frame2.Visible = False
Form1.Height = 6720
End Sub

Private Sub Option2_Click()
Frame2.Visible = True
Frame1.Visible = False
Form1.Height = 6720
End Sub





