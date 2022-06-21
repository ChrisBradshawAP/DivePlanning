VERSION 5.00
Begin VB.Form frmGetS 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ProPlanner - Getting Started"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13140
   Icon            =   "frmGetStarted.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   13140
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   0
      TabIndex        =   8
      Top             =   8040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Show this screen at start up"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   3960
      TabIndex        =   5
      Top             =   8040
      Value           =   1  'Checked
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Finish"
      Height          =   375
      Left            =   11760
      TabIndex        =   4
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Next"
      Height          =   375
      Index           =   1
      Left            =   2280
      TabIndex        =   3
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Previous"
      Height          =   375
      Index           =   0
      Left            =   600
      TabIndex        =   2
      Top             =   8160
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   7680
      Width           =   3375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "ProPlanner"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   3615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Introduction and Getting Started"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   3375
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   7695
      Left            =   3960
      Picture         =   "frmGetStarted.frx":2CFA
      Stretch         =   -1  'True
      Top             =   240
      Width           =   9015
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5415
      Left            =   360
      TabIndex        =   0
      Top             =   2280
      Width           =   3375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   7695
      Left            =   240
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "frmGetS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tipnum As Integer

Private Sub Check1_Click()
  SQL = "SELECT * FROM dpSERIALNO"
  Set RS = DB.OpenRecordset(SQL)
  RS.Edit
  If Check1.Value = vbChecked Then
    RS!getstarted = "Open"
  Else
    RS!getstarted = "Close"
  End If
  RS.Update
    
End Sub

Private Sub Command1_Click(Index As Integer)
Dim gstr(10) As String
Dim gtitle(10) As String

gtitle(0) = "Introduction and Getting Started"
gtitle(1) = "Main Menu"
gtitle(2) = "Gas Blend"
gtitle(3) = "Dive Planning"
gtitle(4) = "Dive Series Planning"
gtitle(5) = "Decompression Result"
gtitle(6) = "Graphical plotting"


gstr(0) = "Thank you for purchasing the ProPlanner software." + vbCrLf + vbCrLf + vbCrLf + "This is a short introduction to getting started with the ProPlanner software." + vbCrLf + vbCrLf + vbCrLf + "Click Next to continue.."
gstr(1) = "This is the central point for you to access and view the plan list." + vbCrLf + vbCrLf + vbCrLf + "You can create, edit and delete dive plans and series from here." + vbCrLf + vbCrLf + vbCrLf + "Dive planning is the individual plan created for a single dive. " + vbCrLf + vbCrLf + "Series planning consists of multiple plan with the surface intervals in between."
gstr(2) = "This menu is for defining the default gas list for a new dive." + vbCrLf + vbCrLf + vbCrLf + "All infomation can be modified." + vbCrLf + vbCrLf + vbCrLf + "'Reset Factory' will reload all the factory default settings." + vbCrLf + "'Save as Default' will save the gas list as the default for future dive planning use."
gstr(3) = "This is the screen for planning a new single dive," + vbCrLf + vbCrLf + vbCrLf + "Picture 1 show the input screen for the setting." + vbCrLf + vbCrLf + "Picture 2 show the gas setting screen for the dive planning selection." + vbCrLf + vbCrLf + "For more details on how to plan a new dive, check on the walk through section."
gstr(4) = "This is the screen for planning a new series dive," + vbCrLf + vbCrLf + vbCrLf + "Picture 1 show the list od the available dive for series planning." + vbCrLf + vbCrLf + "Picture 2 show the sequence of the series planning." + vbCrLf + vbCrLf + "For more details on how to plan a series dive, check on the walk through section."
gstr(5) = "Picture 1 show the three different type of generate the decompression result. For more explaination, please read the manual." + vbCrLf + vbCrLf + vbCrLf + "Picture 2 shows the decompression result. Items in yellow are the calculated decompression result."
gstr(6) = "This section of the dive series planner displays the of a dive within a series of dives. The profile graph of the dive and the textThe depth point with the decompression result." + vbCrLf + vbCrLf + vbCrLf + "Simply click on any point of the graph to obtain the data of the position. "
 
  If Index = 1 Then tipnum = tipnum + 1
  If Index = 0 Then tipnum = tipnum - 1
  If tipnum < 0 Then tipnum = 0
  If tipnum > 6 Then tipnum = 6
  
  If tipnum = 6 Then Command1(1).Enabled = False Else Command1(1).Enabled = True
  If tipnum = 0 Then Command1(0).Enabled = False Else Command1(0).Enabled = True
  
  Label1.Caption = gstr(tipnum) '"First ensure you have all the pin numbers and configuration of the Ouroboros or VR installed correctly. See the seperate maunals for these products. Contact your dealer for PIN numbers."
  Select Case tipnum
  Case 0
     Image1.Picture = LoadPicture(App.Path & "\" & "main6" + ".bmp")
  Case 1
     Image1.Picture = LoadPicture(App.Path & "\" & "main5" + ".bmp")
  Case 2
     Image1.Picture = LoadPicture(App.Path & "\" & "gas3" + ".bmp")
  Case 3
     Image1.Picture = LoadPicture(App.Path & "\" & "diveplans" + ".bmp")
  Case 4
     Image1.Picture = LoadPicture(App.Path & "\" & "diveseriess" + ".bmp")
  Case 5
     Image1.Picture = LoadPicture(App.Path & "\" & "decopic" + ".bmp")
  
  Case 6
     Image1.Picture = LoadPicture(App.Path & "\" & "diveseq" + ".bmp")
  End Select
  Label4.Caption = "Tip " + CStr(tipnum + 1) + " of 7"
  Label2.Caption = gtitle(tipnum)
  
End Sub

Private Sub Command2_Click()
  'main.Show
  Unload Me
End Sub

Private Sub Form_Load()
Dim v
If App.PrevInstance Then
   MsgBox "Application running"
   End
End If
  Me.Left = 0
  Me.Top = 0
Dim OldName
Dim NewName
'On Error Resume Next
dbfilefound = False

File1.Path = App.Path
For i = 1 To File1.ListCount
   File1.ListIndex = i - 1
   tempfileselected = File1.FileName
   If tempfileselected = "planmain.mdb" Then
      dbfilefound = True
   End If
Next i
If dbfilefound = True Then
 If systemstarted = False Then
  Source = App.Path & "\planmain.mdb"
   destinationsource = App.Path & "\planmain2.mdb"
   FileCopy Source, destinationsource
   'DBEngine.CompactDatabase App.Path & "\RB.mdb", App.Path & "\RB2.mdb" 'nickrel2
   Kill App.Path & "\planmain2.mdb" 'nickrel2
   DBEngine.CompactDatabase App.Path & "\planmain.mdb", App.Path & "\planmain2.mdb"
   Kill App.Path & "\planmain.mdb"
   OldName = App.Path & "\planmain2.mdb": NewName = App.Path & "\planmain.mdb" ' Define filenames.
   Name OldName As NewName
   Source = App.Path & "\planmain.mdb"
   destinationsource = App.Path & "\backup.mdb"
   FileCopy Source, destinationsource
 End If
Else
   ans = MsgBox("Main Database not found, Would you like to duplicate from the backup database?", vbYesNo + vbCritical + vbDefaultButton2, "Comfirmation")
   Select Case ans
   Case vbYes
      Source = App.Path & "\backup.mdb"
      destinationsource = App.Path & "\planmain.mdb"
      FileCopy Source, destinationsource
   Case Else
     Unload Me
     End
   End Select
End If
  
  
  Set DB = OpenDatabase(App.Path & "/planmain.mdb")
  SQL = "SELECT * FROM dpSERIALNO"
  Set RS = DB.OpenRecordset(SQL)
  If RS.EOF = True And RS.BOF = True Then
     Unload Me
  Else
     v = RS("getstarted")
     If InStr(1, v, "Close") Then
        Unload Me
     Else
        Check1.Value = vbChecked
        tipnum = 0
        Command1_Click (2)
     End If
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
 '   DB.Close
    Splanmain.Show
End Sub

