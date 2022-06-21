VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmdownload 
   BackColor       =   &H80000013&
   Caption         =   "Form1"
   ClientHeight    =   3525
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7545
   Icon            =   "frmdownload.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3525
   ScaleWidth      =   7545
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6960
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "mdb Files (*.mdb)|*.mdb|All Files (*.*)|*.*"
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   3
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton cmdseperate 
      Caption         =   "Seperate Database"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   2520
      Width           =   2055
   End
   Begin VB.CommandButton cmddefault 
      Caption         =   "Default Database"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "Download the data to default database or Seperated Database ?"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   6135
   End
End
Attribute VB_Name = "frmdownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdcancel_Click()
Unload Me
End Sub

Private Sub cmddefault_Click()
On Error GoTo ErrorHandler2
ErrorHandler2:
    If Err = 32755 Then
        MsgBox "Data is not saved to a file....!!"
        Exit Sub
    End If
End Sub

Private Sub cmdseperate_Click()
 On Error GoTo ErrorHandler2
    CommonDialog1.Action = 2
 Open CommonDialog1.FileName For Output As #1
ErrorHandler2:
    If Err = 32755 Then
        MsgBox "Data is not saved to a file....!!"
        Exit Sub
    End If
End Sub
