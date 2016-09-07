VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form room 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Room Allotment"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10635
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   10635
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1800
      TabIndex        =   24
      Top             =   360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   211943425
      CurrentDate     =   41822
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2520
      TabIndex        =   23
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Enter"
      Height          =   375
      Left            =   240
      TabIndex        =   21
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Remove"
      Height          =   375
      Left            =   9480
      TabIndex        =   19
      Top             =   5400
      Width           =   855
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   375
      Left            =   8520
      Top             =   3840
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   2775
      Left            =   3960
      TabIndex        =   14
      ToolTipText     =   "Double click to select name"
      Top             =   3480
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   4895
      _Version        =   393216
      AllowUpdate     =   0   'False
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   8520
      Top             =   2760
      Visible         =   0   'False
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   8520
      Top             =   3277
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
   Begin VB.CommandButton Command4 
      Caption         =   "Confirm List"
      Height          =   375
      Left            =   8640
      TabIndex        =   10
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear List"
      Height          =   375
      Left            =   7920
      TabIndex        =   9
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Go"
      Height          =   375
      Left            =   6000
      TabIndex        =   7
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6000
      TabIndex        =   6
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6000
      TabIndex        =   5
      Top             =   240
      Width           =   855
   End
   Begin VB.ListBox List2 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2985
      Left            =   7920
      TabIndex        =   1
      Top             =   1920
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   495
      Left            =   6000
      TabIndex        =   0
      Top             =   6480
      Width           =   855
   End
   Begin MSDataGridLib.DataGrid DataGrid3 
      Height          =   2775
      Left            =   240
      TabIndex        =   16
      ToolTipText     =   "Double click to select name"
      Top             =   3480
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   4895
      _Version        =   393216
      AllowUpdate     =   0   'False
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
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFFF&
      Caption         =   "View Allotment"
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
      Height          =   375
      Left            =   8640
      TabIndex        =   22
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label11 
      BackColor       =   &H8000000E&
      Height          =   255
      Left            =   2640
      TabIndex        =   20
      Top             =   6600
      Width           =   3015
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3960
      TabIndex        =   18
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Date:"
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
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Invigilators Alloted:"
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
      Height          =   255
      Left            =   7920
      TabIndex        =   15
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Line Line2 
      BorderWidth     =   5
      X1              =   7560
      X2              =   7560
      Y1              =   0
      Y2              =   7320
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Day of Examination:"
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
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3960
      TabIndex        =   12
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Lecturer"
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
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Number of Teachers Required:"
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
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2040
      Width           =   3255
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Capacity:"
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
      Height          =   255
      Left            =   4680
      TabIndex        =   4
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Room No.:"
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
      Height          =   255
      Left            =   4680
      TabIndex        =   3
      Top             =   240
      Width           =   1170
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Selected Invigilator:"
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
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   6600
      Width           =   2175
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderWidth     =   5
      X1              =   -1200
      X2              =   7560
      Y1              =   1800
      Y2              =   1800
   End
End
Attribute VB_Name = "room"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim res As ADODB.Recordset
Dim res1 As ADODB.Recordset
Dim res2 As ADODB.Recordset
Dim res3 As ADODB.Recordset
Dim con As ADODB.Connection
Dim m As Integer


Private Sub Command1_Click()
Dim s As String
Dim dt1 As Date
Dim temp As ADODB.Recordset

For i = 0 To List2.ListCount - 1
    If List2.List(i) = Label11.Caption Then
    MsgBox ("Already added")
    Exit Sub
    End If
Next i
dt1 = DTPicker1.Value
s = "select * from room where  xdate= # " & dt1 & " # and (t1='" + Trim(Label11.Caption) + "' or t2='" + Trim(Label11.Caption) + "' or t3='" + Trim(Label11.Caption) + "')"
Set temp = New ADODB.Recordset
temp.Open s, con, adOpenDynamic, adLockOptimistic
If Not temp.EOF Then
    Label11.Caption = ""
    Command1.Enabled = False
    s = temp![Number]
    MsgBox ("Already added to Room No." & s)
    Exit Sub
End If
List2.AddItem Label11.Caption
Command3.Enabled = True
Command4.Enabled = True

End Sub

Private Sub Command2_Click()
Dim s As String
Dim dt1 As Date
Dim x As Integer
DataGrid2.Refresh
DataGrid3.Refresh
Label10.Caption = ""

If Text1.Text = "Select" Or Text2.Text = "" Or Text3.Text = "" Then
MsgBox ("Please fill up")
Else

        x = DTPicker1.DayOfWeek
        If x = 1 Then
            Text1.Text = "Sunday"
        ElseIf x = 2 Then
            Text1.Text = "Monday"
        ElseIf x = 3 Then
            Text1.Text = "Tuesday"
        ElseIf x = 4 Then
            Text1.Text = "Wednesday"
        ElseIf x = 5 Then
            Text1.Text = "Thursday"
        ElseIf x = 6 Then
            Text1.Text = "Friday"
        ElseIf x = 7 Then
            Text1.Text = "Saturday"
        End If


    s = "select * from checkroom where rn='" + Trim(Text2.Text) + "'"
    Dim temp As ADODB.Recordset
    Set temp = New ADODB.Recordset
    temp.Open s, con, adOpenDynamic, adLockOptimistic
    If temp.EOF = True Then
        MsgBox ("No such room exists")
        Text2.Text = ""
    Else
         If Val(Text3.Text) > temp![capa] Then
            MsgBox ("Room capacity exceeded")
            Text3.Text = ""
         Else
            If Val(Text3.Text) <= 25 Then
                Label10.Caption = 1
            ElseIf Val(Text3.Text) < 45 And Val(Text3.Text) > 26 Then
                Label10.Caption = 2
            Else
                Label10.Caption = 3
            End If

            dt1 = DTPicker1.Value
            Adodc1.RecordSource = "select * from invigilator where on2 ='" + Text1.Text + "'or on1='" + Text1.Text + "'or on3='" + Text1.Text + "'"
            Set DataGrid2.DataSource = Adodc1
            Adodc1.Refresh

            Adodc3.RecordSource = "select * from teachers where # " & dt1 & " # not between sdate and edate and off1 <>'" + Text1.Text + "'"
            Set DataGrid3.DataSource = Adodc3
            Adodc3.Refresh
            DTPicker1.Enabled = False
            Text2.Enabled = False
            Text3.Enabled = False
        Label11.Caption = ""
        Command1.Enabled = False
        If m <> 2 Then
            List2.Clear
        End If
        Command3.Enabled = False
        Command6.Enabled = False
        Command4.Enabled = False
        DataGrid2.Enabled = True
        DataGrid3.Enabled = True
            
            s = "select * from room where number='" + Trim(Text2.Text) + "' and xdate= # " & dt1 & " #"
          
            Set temp = New ADODB.Recordset
            temp.Open s, con, adOpenDynamic, adLockOptimistic
            If Not temp.EOF Then
                List2.List(0) = temp![t1]
                If temp![t1] <> "" Then
                     List2.List(1) = temp![t2]
                End If
                If temp![t3] <> "" Then
                     List2.List(2) = temp![t3]
                End If
                x = MsgBox("Already added.Want to update?", vbInformation + vbYesNo, "Message")
                If x = vbYes Then
                    m = 2
                    Command4.Enabled = True
                    Command3.Enabled = True
                Else
                List2.Clear
                Command4.Enabled = False
                Command3.Enabled = False
                Command1.Enabled = False
                DataGrid2.Enabled = False
                DataGrid3.Enabled = False
                End If
            End If
        End If
  End If
   
End If

End Sub


Private Sub Command3_Click()
List2.Clear
Command3.Enabled = False
Command6.Enabled = False
Command4.Enabled = False
End Sub


Private Sub Command4_Click()

Dim s As String
 Dim dt1 As Date
 Dim temp As ADODB.Recordset
 Set temp = New ADODB.Recordset
 dt1 = DTPicker1.Value
If (List2.ListCount < Val(Label10.Caption)) Then
 x = MsgBox("More invigilators required.Continue?", vbInformation + vbYesNo, "Notice")
Else
 x = MsgBox("Are you sure?", vbInformation + vbYesNo, "Notice")
End If

If x = vbYes Then

 
 If m = 2 Then
     s = "select * from room where number='" + Trim(Text2.Text) + "' and xdate= # " & dt1 & " #"
     temp.Open s, con, adOpenDynamic, adLockOptimistic
     temp![alloted] = Val(Text3.Text)
     If List2.ListCount = 1 Then
           temp![t1] = List2.List(0)
           'temp![t2] = ""
           ''temp![t3] = ""
     ElseIf List2.ListCount = 2 Then
           temp![t1] = List2.List(0)
           temp![t2] = List2.List(1)
           ''temp![t3] = ""
     Else
           temp![t1] = List2.List(0)
           temp![t2] = List2.List(1)
           temp![t3] = List2.List(2)
     End If
     temp.Update
 Else
    s = "select * from checkroom where rn='" + Trim(Text2.Text) + "'"
    temp.Open s, con, adOpenDynamic, adLockOptimistic
    If temp.EOF = True Then
        MsgBox ("No such record exists")
        
    Else

        dt1 = DTPicker1.Value
        res2.AddNew
        res2![Number] = Trim(Text2.Text)
        res2![xdate] = dt1
        res2![alloted] = Val(Text3.Text)
        res2![capacity] = temp![capa]
        If List2.ListCount = 1 Then
            res2![t1] = List2.List(0)
            ''res2![t2] = ""
            ''res2![t3] = ""
        ElseIf List2.ListCount = 2 Then
            res2![t1] = List2.List(0)
            res2![t2] = List2.List(1)
            ''res2![t3] = ""
        Else
            res2![t1] = List2.List(0)
            res2![t2] = List2.List(1)
            res2![t3] = List2.List(2)
        End If
        res2.Update
    End If
    
  End If
   MsgBox ("Successfully submitted")
     
     Command6.Enabled = False
     Command3.Enabled = False
     Command1.Enabled = False
     Command4.Enabled = False
     List2.Clear
     Label11.Caption = ""
     Label10.Caption = ""
     DataGrid2.Enabled = False
     DataGrid3.Enabled = False
     Command2.Enabled = False
 End If

End Sub

Private Sub Command5_Click()
DTPicker1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Command2.Enabled = True

End Sub

Private Sub Command6_Click()
List2.RemoveItem List2.ListIndex
If (List2.ListCount = 0) Then
 Command3.Enabled = False
 Command4.Enabled = False
End If
Command6.Enabled = False
End Sub

Private Sub DataGrid2_Click()
Label11.Caption = DataGrid2.Columns(1).Value
Command1.Enabled = True
End Sub

Private Sub DataGrid3_Click()
Label11.Caption = DataGrid3.Columns(1).Value
Command1.Enabled = True
End Sub


Private Sub Form_Activate()

Set con = New ADODB.Connection
con.Open " Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Time.mdb;Persist Security Info=False"
Set res = New ADODB.Recordset
Set res1 = New ADODB.Recordset
Set res2 = New ADODB.Recordset
Set res3 = New ADODB.Recordset
res.Open "teachers", con, adOpenDynamic, adLockOptimistic
res1.Open "invigilator", con, adOpenDynamic, adLockOptimistic
res2.Open "room", con, adOpenDynamic, adLockOptimistic
res3.Open "checkroom", con, adOpenDynamic, adLockOptimistic



Command6.Enabled = False
Command3.Enabled = False
Command1.Enabled = False
Command4.Enabled = False
End Sub

Private Sub Label12_Click()
Form4.Show
End Sub

Private Sub List2_Click()
Command6.Enabled = True
End Sub




