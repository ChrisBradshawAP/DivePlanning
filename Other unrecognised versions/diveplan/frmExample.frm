VERSION 5.00
Begin VB.Form frmdisplay 
   Caption         =   "Display Format"
   ClientHeight    =   2130
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   5490
   Icon            =   "frmExample.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2130
   ScaleWidth      =   5490
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option2 
      Caption         =   "Meter"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   480
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Feet"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Main Menu"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdchange 
      Caption         =   "Change"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Current Setting on Display format :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "frmdisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
End Sub

Private Sub cmdcancel_Click()
    SQL = "SELECT * FROM Display "
    Set RS = DB.OpenRecordset(SQL)
    comfirmDisplay = RS("display")
    If comfirmDisplay = "Feet" Then
       Option1 = True
    Else
       Option2 = True
    End If
End Sub

Private Sub cmdchange_Click()
If Option1 = True Then
  comfirmDisplay = "Feet"
Else
  comfirmDisplay = "Meter"
End If
SQL = "SELECT * FROM Display "
Set RS = DB.OpenRecordset(SQL)
RS.Edit
RS!Display = comfirmDisplay
RS.Update
Unload Me
MsgBox " Record updated ! "
End Sub

Private Sub cmdclose_Click()
Unload Me
main.Show
End Sub

Private Sub Form_Load()
Me.Left = (Screen.Width - Me.Width) / 2
Set DB = OpenDatabase(App.Path & "/rb.mdb")
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = 800
SQL = "SELECT * FROM Display "
Set RS = DB.OpenRecordset(SQL)
comfirmDisplay = RS("display")
If comfirmDisplay = "Feet" Then
  Option1 = True
Else
  Option2 = True
End If
End Sub

Private Sub mnuSub_Click(Index As Integer)
End Sub
