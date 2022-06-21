VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSplash 
   BackColor       =   &H00004000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4245
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4050
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   7080
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   3480
         Top             =   2520
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FFFF&
         Caption         =   "FIX DATABASE"
         Height          =   735
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Image Image1 
         Height          =   2265
         Left            =   120
         Picture         =   "frmSplash.frx":000C
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label lbllicense 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   6855
      End
      Begin VB.Label lblCopyright 
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3840
         TabIndex        =   3
         Top             =   3300
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label lblCompany 
         BackStyle       =   0  'Transparent
         Caption         =   "Company"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3840
         TabIndex        =   2
         Top             =   3510
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label lblWarning 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Warning"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   150
         TabIndex        =   1
         Top             =   3660
         Width           =   6855
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5880
         TabIndex        =   4
         Top             =   2880
         Width           =   885
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   765
         Left            =   3600
         TabIndex        =   5
         Top             =   1080
         Width           =   2430
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command1_Click()
On Error Resume Next
 Timer1.Enabled = False
 ans = MsgBox("Restore backup database file to fix database corruption. Note, most recent planned dives willbe lost! Backup2.mdb is most recent backup file" & vbCrLf & "Fix Database?", vbYesNo, "Restore backup database")
 If ans = vbYes Then
   CommonDialog1.Filter = "*.mdb | *.mdb"
   CommonDialog1.Action = 1
   If Right(CommonDialog1.FileName, 3) = "mdb" Then
     FileCopy CommonDialog1.FileName, "planmain.mdb"
     Command1.BackColor = vbGreen
   End If
 End If
 Timer1.Enabled = True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

 Private Sub Form_Load()
    lblVersion.Caption = "Copyright Nick Bushell 1992-2008" ' "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = "ProPlanner" 'App.Title
    'lbllicense.Caption = App.li
End Sub

Private Sub Frame1_Click()
On Error Resume Next
   Timer1.Enabled = False
   frmGetS.Show
   Unload Me
    'frmGetS.Show 'Show 'main.Show 'frmGetS.Show
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Frame1_Click
End Sub
