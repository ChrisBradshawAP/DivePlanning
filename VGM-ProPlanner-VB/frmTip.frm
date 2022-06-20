VERSION 5.00
Begin VB.Form frmTip 
   Caption         =   "VGM and VRx"
   ClientHeight    =   7140
   ClientLeft      =   2370
   ClientTop       =   2400
   ClientWidth     =   10095
   Icon            =   "frmTip.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   10095
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton Command1 
      Caption         =   "&Previous Page"
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CheckBox chkLoadTipsAtStartup 
      Caption         =   "&Show Tips at Startup"
      Height          =   315
      Left            =   4800
      TabIndex        =   3
      Top             =   6720
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdNextTip 
      Caption         =   "&Next Page"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   6720
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5475
      Left            =   120
      ScaleHeight     =   5475
      ScaleWidth      =   9855
      TabIndex        =   1
      Top             =   1080
      Width           =   9855
      Begin VB.Label lblTipText 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   5295
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   9615
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "VGM and VRx information"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   6
      Top             =   360
      Width           =   8895
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   1455
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   9855
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ShowAtStartup As Long
Dim CurrentTip As Long
' The in-memory database of tips.
Dim Tips As New Collection
Dim stemp As String

Option Explicit

' The in-memory database of tips.
'Dim Tips As New Collection

' Name of tips file
Const TIP_FILE = "TIPOFDAY.TXT"
Const TIP_FILE2 = "TIPOFDAY2.TXT"
Const TIP_FILE3 = "TIPOFDAY3.TXT"
Const TIP_FILE4 = "TIPOFDAY4.TXT"

' Index in collection of tip currently being displayed.


Public Sub DoNextTip()
Dim T As Integer
On Error Resume Next

    ' Select a tip at random.
    'CurrentTip = Int((Tips.Count * Rnd) + 1)
    
    ' Or, you could cycle through the Tips in order

    CurrentTip = CurrentTip + 1
    T = Tips.Count
    If Tips.Count < CurrentTip Then
      CurrentTip = 1
    End If
    'CurrentTip = Tips.Count
    ' Show it.
    frmTip.DisplayCurrentTip
    
End Sub

Function LoadTips(sFile As String) As Boolean
    Dim Nexttip As String   ' Each tip read in from file.
    Dim InFile As Integer   ' Descriptor for file.
    
    ' Obtain the next free file descriptor.
    InFile = FreeFile
    
    ' Make sure a file is specified.
    If sFile = "" Then
        LoadTips = False
        Exit Function
    End If
    
    ' Make sure the file exists before trying to open it.
    If Dir(sFile) = "" Then
        LoadTips = False
        Exit Function
    End If
    
    ' Read the collection from a text file.
    Open sFile For Input As InFile
    Line Input #InFile, Nexttip
    stemp = Nexttip
    While Not EOF(InFile)
        Line Input #InFile, Nexttip
        stemp = stemp & vbCrLf & Nexttip
    Wend
    Tips.Add stemp  'Nexttip
    Close InFile

    ' Display a tip at random.
    DoNextTip
    
    LoadTips = True
    
End Function

Private Sub chkLoadTipsAtStartup_Click()
    ' save whether or not this form should be displayed at startup
    SaveSetting App.EXEName, "Options", "Show Tips at Startup", chkLoadTipsAtStartup.Value
    ShowAtStartup = GetSetting(App.EXEName, "Options", "Show Tips at Startup", 1)
End Sub

Private Sub cmdNextTip_Click()
    DoNextTip
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
  DoPreviousTip
End Sub

Private Sub Form_Load()
    Dim ShowAtStartup As Long
    Dim i As Integer
    
    Do While Tips.Count > 0
      Tips.Remove (1)
    Loop
    ' See if we should be shown at startup
    ShowAtStartup = GetSetting(App.EXEName, "Options", "Show Tips at Startup", 1)
    ' Set the checkbox, this will force the value to be written back out to the registry
    Me.chkLoadTipsAtStartup.Value = vbChecked
    
    ' Seed Rnd
'    Randomize
    
    ' Read in the tips file and display a tip at random.
    If LoadTips(App.Path & "\" & TIP_FILE) = False Then
        lblTipText.Caption = "That the " & TIP_FILE & " file was not found? " & vbCrLf & vbCrLf & _
           "Create a text file named " & TIP_FILE & " using NotePad with 1 tip per line. " & _
           "Then place it in the same directory as the application. "
    End If

    If LoadTips(App.Path & "\" & TIP_FILE2) = False Then
        lblTipText.Caption = "That the " & TIP_FILE2 & " file was not found? " & vbCrLf & vbCrLf & _
           "Create a text file named " & TIP_FILE2 & " using NotePad with 1 tip per line. " & _
           "Then place it in the same directory as the application. "
    End If

    If LoadTips(App.Path & "\" & TIP_FILE3) = False Then
        lblTipText.Caption = "That the " & TIP_FILE3 & " file was not found? " & vbCrLf & vbCrLf & _
           "Create a text file named " & TIP_FILE3 & " using NotePad with 1 tip per line. " & _
           "Then place it in the same directory as the application. "
    End If

    CurrentTip = 0
    If LoadTips(App.Path & "\" & TIP_FILE4) = False Then
        lblTipText.Caption = "That the " & TIP_FILE4 & " file was not found? " & vbCrLf & vbCrLf & _
           "Create a text file named " & TIP_FILE4 & " using NotePad with 1 tip per line. " & _
           "Then place it in the same directory as the application. "
    End If

    If ShowAtStartup = 0 Then
        frmTip.Visible = False
        Exit Sub
    End If
        
    
End Sub

Public Sub DisplayCurrentTip()
'    If Tips.Count > 0 And frmTip.Visible = True Then
'        lblTipText.Caption = "Step" + CStr(CurrentTip) + ":   " + Tips.Item(CurrentTip)
'    End If
    If Tips.Count > 0 Then
'        lblTipText.Caption = "Tip " + CStr(CurrentTip) + ":   " + Tips.Item(CurrentTip)
        lblTipText.Caption = Tips.Item(CurrentTip)
    End If
End Sub

Private Sub Image1_Click()
  cmdNextTip_Click
End Sub

Public Sub DoPreviousTip()
Dim T As Integer

    ' Select a tip at random.
    'CurrentTip = Int((Tips.Count * Rnd) + 1)
    
    ' Or, you could cycle through the Tips in order

    CurrentTip = CurrentTip - 1
    T = Tips.Count
    If CurrentTip < 1 Then
      CurrentTip = T
    End If
    'CurrentTip = Tips.Count
    ' Show it.
    frmTip.DisplayCurrentTip
    
End Sub


Private Sub lblTipText_Click()
  cmdNextTip_Click
End Sub
