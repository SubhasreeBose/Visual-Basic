VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form3 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View"
   ClientHeight    =   3870
   ClientLeft      =   7050
   ClientTop       =   2130
   ClientWidth     =   6345
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   6345
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "Form3.frx":0000
      Height          =   2175
      Left            =   3240
      TabIndex        =   21
      ToolTipText     =   "Double click to select name"
      Top             =   1440
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   3836
      _Version        =   393216
      BackColor       =   65535
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "ID"
         Caption         =   "ID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "tname"
         Caption         =   "Name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form3.frx":0015
      Height          =   2175
      Left            =   240
      TabIndex        =   20
      ToolTipText     =   "Double click to select name"
      Top             =   1440
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   3836
      _Version        =   393216
      BackColor       =   65535
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "IID"
         Caption         =   "ID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "iname"
         Caption         =   "Name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   615
      Left            =   9720
      Top             =   240
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      RecordSource    =   "teachers"
      Caption         =   "Adodc3"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   615
      Left            =   8040
      Top             =   240
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      RecordSource    =   "invigilator"
      Caption         =   "Adodc2"
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
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FFFF&
      Caption         =   "Delete"
      Height          =   375
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8160
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Update"
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8160
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   6360
      Top             =   240
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   5295
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF0000&
      Caption         =   "Teacher"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   3975
      Left            =   120
      TabIndex        =   2
      Top             =   3960
      Visible         =   0   'False
      Width           =   6015
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   495
         Left            =   2640
         TabIndex        =   19
         Top             =   2160
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   873
         _Version        =   393216
         Format          =   92667905
         CurrentDate     =   41805
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2640
         TabIndex        =   14
         Top             =   1305
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2640
         TabIndex        =   3
         Top             =   600
         Width           =   2775
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   495
         Left            =   2640
         TabIndex        =   18
         Top             =   3120
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   873
         _Version        =   393216
         Format          =   92667905
         CurrentDate     =   41805
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   495
         Left            =   2640
         TabIndex        =   25
         Top             =   3120
         Width           =   2415
      End
      Begin VB.Label Label9 
         Caption         =   "Label3"
         Height          =   495
         Left            =   2640
         TabIndex        =   24
         Top             =   2160
         Width           =   2415
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FF0000&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   5160
         TabIndex        =   23
         Top             =   3360
         Width           =   255
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FF0000&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   5160
         TabIndex        =   22
         Top             =   2280
         Width           =   255
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FF0000&
         Caption         =   "Ending Date:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   7
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FF0000&
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   5
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF0000&
         Caption         =   "Off Day:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   1290
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FF0000&
         Caption         =   "Starting Date:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   2280
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF0000&
      Caption         =   "Guest Lecturer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   3975
      Left            =   120
      TabIndex        =   8
      Top             =   3960
      Visible         =   0   'False
      Width           =   6015
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   2520
         TabIndex        =   17
         Text            =   "Combo4"
         Top             =   3000
         Width           =   2775
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   2520
         TabIndex        =   16
         Text            =   "Combo3"
         Top             =   2220
         Width           =   2775
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   2520
         TabIndex        =   15
         Text            =   "Combo2"
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   2520
         TabIndex        =   9
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FF0000&
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   11
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FF0000&
         Caption         =   "On Days:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   10
         Top             =   1440
         Width           =   1215
      End
   End
   Begin VB.Image Image1 
      Height          =   405
      Left            =   5760
      Picture         =   "Form3.frx":002A
      Stretch         =   -1  'True
      Top             =   720
      Width           =   315
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Search by name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Dim res As ADODB.Recordset
Dim res1 As ADODB.Recordset
Dim flag As Integer

Private Sub Command1_Click()

If Text1.Text = "" Then
 MsgBox "Please enter a name", vbCritical, "Error"
Else
 Dim temp As ADODB.Recordset
 Set temp = New ADODB.Recordset
 Dim s As String
 Dim dt1 As Date
 Dim dt2 As Date
 s = "select * from teachers where tname='" + Trim(Text1.Text) + "'"
 Dim s1 As String
 s1 = "select * from invigilator where iname='" + Trim(Text1.Text) + "'"
 Dim temp1 As ADODB.Recordset
 Set temp1 = New ADODB.Recordset
 If flag = 2 Then
    'dt1 = Format(Text5.Text, "mm/dd/yyyy")
    'dt2 = Format(Text6.Text, "mm/dd/yyyy")
    dt1 = DTPicker1.Value
    dt2 = DTPicker2.Value
    
        temp.Open s, con, adOpenDynamic, adLockOptimistic
        temp![tname] = Text2.Text
        temp![off1] = Combo1.Text

        'temp![earn] = Val(Text3.Text) - Val(a)

        temp![sdate] = dt1
        temp![edate] = dt2
        temp.Update
        MsgBox ("Updated successfully")
         If temp![sdate] <> "" Then
            Label9.Caption = temp![sdate]
        End If
        If temp![sdate] <> "" Then
            Label3.Caption = temp![edate]
        End If
   
    
    Combo2.Clear
    Combo2.AddItem "Monday"
    Combo2.AddItem "Tuesday"
    Combo2.AddItem "Wednesday"
    Combo2.AddItem "Thursday"
    Combo2.AddItem "Friday"
    Combo2.AddItem "Saturday"


    Combo3.Clear
    Combo3.AddItem "Monday"
    Combo3.AddItem "Tuesday"
    Combo3.AddItem "Wednesday"
    Combo3.AddItem "Thursday"
    Combo3.AddItem "Friday"
    Combo3.AddItem "Saturday"


    Combo4.Clear
    Combo4.AddItem "Monday"
    Combo4.AddItem "Tuesday"
    Combo4.AddItem "Wednesday"
    Combo4.AddItem "Thursday"
    Combo4.AddItem "Friday"
    Combo4.AddItem "Saturday"
 End If

 If flag = 3 Then
   temp1.Open s1, con, adOpenDynamic, adLockOptimistic
   temp1![iname] = Text7.Text
   temp1![on1] = Text8.Text
   temp1![on2] = Text9.Text
   temp1![on3] = Text10.Text
   temp1.Update
   MsgBox ("Updated successfully")
 End If
End If

End Sub

Private Sub Command2_Click()
Dim temp As ADODB.Recordset
Set temp = New ADODB.Recordset
Dim s As String
s = "delete * from teachers where tname='" + Trim(Text1.Text) + "'"

Dim s1 As String
s1 = "delete * from invigilator where iname='" + Trim(Text1.Text) + "'"
Dim temp1 As ADODB.Recordset
Set temp1 = New ADODB.Recordset

If flag = 2 Then
temp.Open s, con, adOpenDynamic, adLockOptimistic
MsgBox ("Deleted successfully")
Form3.Height = 1770
End If
If flag = 3 Then
temp1.Open s1, con, adOpenDynamic, adLockOptimistic
MsgBox ("Deleted successfully")
Form3.Height = 1770
End If
Unload Me
End Sub

Private Sub DataGrid1_Click()
Text1.Text = DataGrid1.Columns(1).Value
End Sub

Private Sub DataGrid2_Click()
Text1.Text = DataGrid2.Columns(1).Value
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

Command1.Enabled = False
Command2.Enabled = False

End Sub
Private Sub Combo3_DropDown()
If Combo2.ListIndex <> -1 Then
Combo3.RemoveItem Combo2.ListIndex
Combo4.RemoveItem Combo2.ListIndex
End If
End Sub



Private Sub Combo4_DropDown()
If Combo3.ListIndex <> -1 Then
Combo4.RemoveItem Combo3.ListIndex
End If
End Sub

Private Sub Image1_Click()
DTPicker1.Visible = False
DTPicker2.Visible = False
Label3.Caption = ""
Label9.Caption = ""
Label10.Visible = False
Label11.Visible = False

Dim temp As ADODB.Recordset
Set temp = New ADODB.Recordset
Dim s As String
s = "select * from teachers where tname='" + Trim(Text1.Text) + "'"

Dim s1 As String
s1 = "select * from invigilator where iname='" + Trim(Text1.Text) + "'"
Dim temp1 As ADODB.Recordset
Set temp1 = New ADODB.Recordset

temp.Open s, con, adOpenDynamic, adLockOptimistic
If Not temp.EOF = True Then
 Frame1.Visible = True
 Frame2.Visible = False
 flag = 2
 Text2.Text = temp![tname]
 Combo1.Text = temp![off1]
 If temp![sdate] <> "" Then
    Label9.Caption = temp![sdate]
 End If
 If temp![sdate] <> "" Then
   Label3.Caption = temp![edate]
 End If
 Form3.Height = 6720
 Command1.Enabled = True
 Command2.Enabled = True
End If

temp1.Open s1, con, adOpenDynamic, adLockOptimistic
If Not temp1.EOF = True Then
 Frame2.Visible = True
 Frame1.Visible = False
 flag = 3
 Text7.Text = temp1![iname]
 Combo2.Text = temp1![on1]
 Combo3.Text = temp1![on2]
 Combo4.Text = temp1![on3]
 Form3.Height = 6720
 Command1.Enabled = True
 Command2.Enabled = True
End If

If flag <> 2 And flag <> 3 Then
 MsgBox "No such record exists", vbCritical, "Error"
End If
Me.Height = 9150

End Sub

Private Sub Label10_Click()

Label9.Caption = DTPicker2.Value

Label9.Visible = True
DTPicker2.Visible = False
Label10.Visible = False
End Sub

Private Sub Label11_Click()

Label3.Caption = DTPicker1.Value

Label3.Visible = True
DTPicker1.Visible = False
Label11.Visible = False
End Sub

Private Sub Label9_Click()
DTPicker2.Visible = True
Label10.Visible = True

End Sub

Private Sub Label3_Click()
DTPicker1.Visible = True
Label11.Visible = True
End Sub
