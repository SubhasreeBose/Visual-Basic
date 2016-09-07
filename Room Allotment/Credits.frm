VERSION 5.00
Begin VB.Form Credits 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Credits"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3390
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   3390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   195
      TabIndex        =   0
      Top             =   248
      Width           =   3015
   End
End
Attribute VB_Name = "Credits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim str, str1, str2 As String
str = "This app has been developed by:" & vbCrLf & vbCrLf & "Subhasree Bose" & vbCrLf & "Subhayu Chakravorty"
str1 = vbCrLf & vbCrLf & "This is a personal and customized software." & vbCrLf & "Please Do Not reproduce this software for any monetory gains."
str2 = vbCrLf & vbCrLf & "(C) FordoX"
Label1.Caption = str & str1 & str2
End Sub
