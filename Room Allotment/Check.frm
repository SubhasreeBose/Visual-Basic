VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form4 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Allotment"
   ClientHeight    =   585
   ClientLeft      =   5040
   ClientTop       =   2880
   ClientWidth     =   2490
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   585
   ScaleWidth      =   2490
   Begin MSComDlg.CommonDialog c 
      Left            =   360
      Top             =   6840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   480
      TabIndex        =   7
      Top             =   4800
      Width           =   3015
   End
   Begin VB.CommandButton Command5 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Caption         =   "Delete all"
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
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Caption         =   "Delete room"
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
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Caption         =   "Add Room"
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
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4200
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   120
      Top             =   7560
      Width           =   1200
      _ExtentX        =   2117
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
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Caption         =   "Print"
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
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Caption         =   "Delete Allocation"
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
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4800
      Width           =   1935
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3375
      Left            =   480
      TabIndex        =   0
      ToolTipText     =   "Double click to select No."
      Top             =   600
      Visible         =   0   'False
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   5953
      _Version        =   393216
      BackColor       =   8454143
      HeadLines       =   1
      RowHeight       =   19
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
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "View Allocation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim res As ADODB.Recordset
Dim res1 As ADODB.Recordset
Dim res2 As ADODB.Recordset
Dim res3 As ADODB.Recordset
Dim con As ADODB.Connection

Private Sub Command1_Click()
Dim s As String
Dim temp As ADODB.Recordset

x = MsgBox("Are you sure?", vbYesNo, "Message")
If x = vbYes Then
s = "select * from room where number='" + Trim(Text1.Text) + "'"
    Set temp = New ADODB.Recordset
    temp.Open s, con, adOpenDynamic, adLockOptimistic
    If temp.EOF = True Then
        MsgBox ("No such room allocation exists")
    Else
        Set temp = New ADODB.Recordset
        s = "delete * from room where number='" + Trim(Text1.Text) + "'"
        temp.Open s, con, adOpenDynamic, adLockOptimistic
        DataGrid1.Visible = False
        Adodc1.RecordSource = "select * from room order by number desc"
        Set DataGrid1.DataSource = Adodc1
        Adodc1.Refresh
    End If
End If
End Sub

Private Sub Command2_Click()
c.ShowPrinter
End Sub

Private Sub Command3_Click()
Dim a As Integer
Dim b As String
b = InputBox("Enter Room No.:")
If b = "" Then
     MsgBox "Invalid Room No.!", vbCritical, "Room No."
Else
a = Val(InputBox("Enter Capacity:"))
If a <= 0 Then
     MsgBox "Invalid Capacity!", vbCritical, "Capacity"
Else
    Dim temp As ADODB.Recordset
    s = "select * from checkroom where rn='" + b + "'"
    Set temp = New ADODB.Recordset
    temp.Open s, con, adOpenDynamic, adLockOptimistic
     If temp.EOF = True Then
     res3.AddNew
     res3![rn] = b
     res3![capa] = a
     res3.Update
     MsgBox "Successfully Added!", vbInformation, "Added"
     Else
     MsgBox "Already exists", vbCritical, "Room Number"
     End If
     
     
End If
End If
End Sub

Private Sub Command4_Click()
Dim temp As ADODB.Recordset
x = MsgBox("Are you sure?", vbYesNo, "Message")
If x = vbYes Then

s = "select * from checkroom where rn='" + Trim(Text2.Text) + "'"
    Set temp = New ADODB.Recordset
    temp.Open s, con, adOpenDynamic, adLockOptimistic
    If temp.EOF = True Then
        MsgBox ("No such room exists")
    Else
        Set temp = New ADODB.Recordset
        s = "delete * from room where number='" + Trim(Text2.Text) + "'"
        temp.Open s, con, adOpenDynamic, adLockOptimistic

        Dim temp1 As ADODB.Recordset
        Set temp1 = New ADODB.Recordset
        s = "delete * from checkroom where rn='" + Trim(Text2.Text) + "'"
        temp1.Open s, con, adOpenDynamic, adLockOptimistic
        MsgBox ("Successfully deleted")
        DataGrid1.Visible = False
    End If
End If
End Sub

Private Sub Command5_Click()
Dim s As String
x = MsgBox("This will delete all allocations.You will no longer be able to retrieve it.Continue?", vbYesNo, "Warning")
If x = vbYes Then
 Dim temp As ADODB.Recordset
 Set temp = New ADODB.Recordset
 s = "delete * from room"
 temp.Open s, con, adOpenDynamic, adLockOptimistic
 DataGrid1.Visible = False
End If
End Sub

Private Sub DataGrid1_Click()
Text1.Text = DataGrid1.Columns(0).Value

End Sub



Private Sub Form_Load()
Adodc1.RecordSource = "select * from room order by number desc"
Set DataGrid1.DataSource = Adodc1
Set con = New ADODB.Connection
con.Open " Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Time.mdb;Persist Security Info=False"
Set res2 = New ADODB.Recordset
Set res3 = New ADODB.Recordset
res2.Open "room", con, adOpenDynamic, adLockOptimistic
res3.Open "checkroom", con, adOpenDynamic, adLockOptimistic

End Sub

Private Sub Label1_Click()
Adodc1.RecordSource = "select * from room order by number desc"
Set DataGrid1.DataSource = Adodc1
Adodc1.Refresh
DataGrid1.Visible = True
For i = 0 To 2
DataGrid1.Columns(i).Width = 1000
Next i
For i = 3 To 6
DataGrid1.Columns(i).Width = 1500
Next i
DataGrid1.Columns(0).Caption = "No."
DataGrid1.Columns(1).Caption = "Capacity"
DataGrid1.Columns(2).Caption = "Alloted"
DataGrid1.Columns(3).Caption = "Date"
DataGrid1.Columns(4).Caption = "Invigilator"
DataGrid1.Columns(5).Caption = "Invigilator"
DataGrid1.Columns(6).Caption = "Invigilator"
Me.Height = 5895
Me.Width = 10500

End Sub
