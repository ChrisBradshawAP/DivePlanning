VERSION 5.00
Begin VB.Form frmtips 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tips Windows"
   ClientHeight    =   2505
   ClientLeft      =   255
   ClientTop       =   1800
   ClientWidth     =   6690
   Icon            =   "frmtips.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   590
      Left            =   2565
      Picture         =   "frmtips.frx":000C
      ScaleHeight     =   585
      ScaleWidth      =   630
      TabIndex        =   13
      Top             =   40
      Width           =   630
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2610
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   1905
      Begin VB.Image Image1 
         Height          =   1275
         Left            =   0
         Picture         =   "frmtips.frx":13CE
         Top             =   1200
         Width           =   1830
      End
      Begin VB.Label lblproduct 
         BackStyle       =   0  'Transparent
         Caption         =   "Professional  dive planning software ...."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   930
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.TextBox tipstext 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2160
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Choose you category to view the tips"
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
      Left            =   2280
      TabIndex        =   12
      Top             =   720
      Width           =   3615
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Next Tip"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   4560
      MouseIcon       =   "frmtips.frx":8E40
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   2160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3840
      MouseIcon       =   "frmtips.frx":914A
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Menu"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3120
      MouseIcon       =   "frmtips.frx":9454
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   2160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Gas Profile"
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
      Left            =   4440
      MouseIcon       =   "frmtips.frx":975E
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Dive Series"
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
      Left            =   4440
      MouseIcon       =   "frmtips.frx":9A68
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Library Dive"
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
      Height          =   495
      Left            =   2520
      MouseIcon       =   "frmtips.frx":9D72
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Main Menu"
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
      Left            =   2520
      MouseIcon       =   "frmtips.frx":A07C
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Did You Know....."
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmtips"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
   Unload Me
End Sub

 Private Sub Form_Load()
 Dim comptext As String
' Dim temptext3 As String
    ' "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = "ProPlanner" 'App.Title
    'lbllicense.Caption = App.li
End Sub

Private Sub Frame1_Click()
On Error Resume Next
'    Timer1.Enabled = False
'    frmGetS.Show
 '   Unload Me
    'frmGetS.Show 'Show 'main.Show 'frmGetS.Show
End Sub

Private Sub lblCompany_Click()

End Sub

Private Sub Label2_Click()
Dim comptext As String
Label3.Visible = True
Label4.Visible = True
Label5.Visible = True
Label6.Visible = True
Label9.Visible = True
tipstext.Visible = False
Label2.Visible = False
Label7.Visible = True
Label8.Visible = False
End Sub

Private Sub Label3_Click()
Dim comptext As String
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Label9.Visible = False
tipstext.Visible = True
Label2.Visible = True
Label7.Visible = True
Label8.Visible = True
comptext = "Main"
SQL = "select * FROM tips "
SQL = SQL & "WHERE tipsname = '" & Trim(comptext) & "'"
'SQL = SQL & "order by tipstext "
Set RS6 = DB.OpenRecordset(SQL)
RS6.MoveFirst
tipstext.Text = RS6("tipstext")
End Sub

Private Sub Label4_Click()
Dim comptext As String
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Label9.Visible = False
tipstext.Visible = True
Label2.Visible = True
Label7.Visible = True
Label8.Visible = True
SQL = "select * FROM tips "
SQL = SQL & "WHERE tipsname = 'Library' "
'SQL = SQL & "order by tipstext "
Set RS6 = DB.OpenRecordset(SQL)
RS6.MoveFirst
tipstext.Text = RS6("tipstext")
End Sub

Private Sub Label5_Click()
Dim comptext As String
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Label9.Visible = False
tipstext.Visible = True
Label2.Visible = True
Label7.Visible = True
Label8.Visible = True
SQL = "select * FROM tips "
SQL = SQL & "WHERE tipsname = 'Series' "
'SQL = SQL & "order by tipstext "
Set RS6 = DB.OpenRecordset(SQL)
RS6.MoveFirst
tipstext.Text = RS6("tipstext")
End Sub

Private Sub Label6_Click()
Dim comptext As String
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Label9.Visible = False
tipstext.Visible = True
Label2.Visible = True
Label7.Visible = True
Label8.Visible = True
SQL = "select * FROM tips "
SQL = SQL & "WHERE tipsname = 'Gas' "
'SQL = SQL & "order by tipstext "
Set RS6 = DB.OpenRecordset(SQL)
RS6.MoveFirst
tipstext.Text = RS6("tipstext")
End Sub

Private Sub Label7_Click()
Unload Me
End Sub

Private Sub Label8_Click()
RS6.MoveNext
If RS6.EOF = False Then
  
  tipstext.Text = RS6("tipstext")
Else
  RS6.MoveFirst
  tipstext.Text = RS6("tipstext")
End If
End Sub

Private Sub Timer1_Timer()
'    Timer1.Enabled = False
'    Frame1_Click
End Sub
