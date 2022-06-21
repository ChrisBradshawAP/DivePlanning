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
   Begin VB.CommandButton Command3 
      Caption         =   "VGM Info"
      Height          =   375
      Left            =   9000
      TabIndex        =   9
      Top             =   8160
      Width           =   1335
   End
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
         Name            =   "Arial"
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
      Left            =   2280
      TabIndex        =   4
      Top             =   8160
      Visible         =   0   'False
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5535
      Left            =   360
      TabIndex        =   0
      Top             =   2160
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

Private Sub Command1_Click(index As Integer)
Dim gstr(10) As String
Dim gtitle(10) As String

gtitle(0) = "Terms of Use"
gtitle(1) = "Main Menu"
gtitle(2) = "Dive Planning"
gtitle(3) = "Gas adjust"
gtitle(4) = "Decompression Result"


gstr(0) = "Thank you for trying the VGM DivePlan software." + vbCrLf + vbCrLf + "This software can be freely distributed on a use as is basis. The authors accept no responsibility for the accuracy of any schedules planned on this software. All dive plans and decompression should only be undertaken if the diver is clear and happy with the schedules formulated. Variations in diver physique, fitness, preparation, hydration etc.. make it impossible to make a universal dive plan. So always use your experience and judgement in choosing a decompression schedule that is safe for you." + vbCrLf + "If you agree to these terms, click Accept. If not, Click Not accept."
gstr(0) = "Decompression and no stop limit diving put a considerable strain on the human body even when diving in the limits of decompression schedules." + vbcrl + "This software can be freely distributed on a 'use as is' basis. The authors accept no responsibility for the accuracy of any schedules planned on this software. All dive plans and decompression should only be undertaken if the diver is clear and happy with the schedules formulated. Variations in diver physique, fitness, preparation, hydration etc.. make it impossible to make a universal dive plan. So always use your experience and judgement in choosing a decompression schedule that is safe for you. Always refer to the recommendations of your diver training agency." + vbCrLf + "If you agree to these terms, click Accept. If not, Click Not accept."
gstr(1) = "From this screen you can plan a new dive, or edit an existing dive." + vbCrLf + vbCrLf + vbCrLf + "Click Plan New Dive to start a new dive plan." + vbCrLf + vbCrLf + "Double click an existing dive to edit that dive."
gstr(2) = "With this free software, a single depth dive can be planned." + vbCrLf + vbCrLf + "See detailed walk through of dive plan tasks in subsequent section." + vbCrLf + vbCrLf + "Users can adjust the decompression by changing the safety bubble control associated with Fast, Middle and Slow tissues. The approximate Equivalent Gradient Factor is shown for comparrison with other dive plans."
gstr(3) = "This section of the diveplanning screen allows specific gasses to be adjusted from the default values. Use the +- button to increase change the oxygen and helium content"
gstr(4) = "Items in yellow are the calculated decompression result."
 
  If index = 1 Then tipnum = tipnum + 1
  If index = 0 And tipnum > 0 Then tipnum = tipnum - 1
  If tipnum < 0 Then tipnum = 0
  If tipnum > 4 Then tipnum = 4
  
  If tipnum = 0 Then
    Command1(0).Caption = "Not Accept"
    Command1(1).Caption = "Accept"
  Else
    Command1(0).Caption = "Previous"
    Command1(1).Caption = "Next"
  End If
  
  If tipnum = 4 Then Command1(1).Enabled = False Else Command1(1).Enabled = True
  If tipnum = 0 Then Command1(0).Enabled = False Else Command1(0).Enabled = True
  Command1(0).Enabled = True
  If tipnum = 4 Then
    Command2.Visible = True
    Command2.SetFocus
  Else
    Command2.Visible = False
  End If
  Label1.Caption = gstr(tipnum) '"First ensure you have all the pin numbers and configuration of the Ouroboros or VR installed correctly. See the seperate maunals for these products. Contact your dealer for PIN numbers."
  Select Case tipnum
  Case 0
     Image1.Picture = LoadPicture(App.Path & "\" & "main6" + ".bmp")
  Case 1
     Image1.Picture = LoadPicture(App.Path & "\" & "main9" + ".bmp")
  Case 2
     Image1.Picture = LoadPicture(App.Path & "\" & "diveplans" + ".bmp")
  Case 3
     Image1.Picture = LoadPicture(App.Path & "\" & "diveplans2" + ".bmp")
  Case 4
     Image1.Picture = LoadPicture(App.Path & "\" & "decopic" + ".bmp")
  End Select
  Label4.Caption = "Page " + CStr(tipnum + 1) + " of 5"
  Label2.Caption = gtitle(tipnum)
  
If index = 0 Then
  If tipnum = 0 Then
    MsgBox "Exiting VGM DivePlan as terms not accepted"
    Unload Me
  End If
End If

End Sub

Private Sub Command2_Click()
  'main.Show
  Unload Me
  frmintro.Show
End Sub

Private Sub Command3_Click()
frmTip.Show
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

'MsgBox App.Path


For i = 1 To File1.ListCount
   File1.ListIndex = i - 1
   tempfileselected = File1.FileName
   If tempfileselected = "planmain.mdb" Or tempfileselected = "planmain.MDB" Then
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
   On Error Resume Next
'   Kill App.Path & "\planmain2.mdb" 'nickrel2
'   Kill App.Path & "\backup9.mdb" 'nickrel2
   FileCopy App.Path & "\backup8.mdb", App.Path & "\backup9.mdb"
'   Kill App.Path & "\backup8.mdb" 'nickrel2
   FileCopy App.Path & "\backup7.mdb", App.Path & "\backup8.mdb"
'   Kill App.Path & "\backup7.mdb" 'nickrel2
   FileCopy App.Path & "\backup6.mdb", App.Path & "\backup7.mdb"
'   Kill App.Path & "\backup6.mdb" 'nickrel2
   FileCopy App.Path & "\backup5.mdb", App.Path & "\backup6.mdb"
'   Kill App.Path & "\backup5.mdb" 'nickrel2
   FileCopy App.Path & "\backup4.mdb", App.Path & "\backup5.mdb"
'   Kill App.Path & "\backup4.mdb" 'nickrel2
   FileCopy App.Path & "\backup3.mdb", App.Path & "\backup4.mdb"
'   Kill App.Path & "\backup3.mdb" 'nickrel2
   FileCopy App.Path & "\backup2.mdb", App.Path & "\backup3.mdb"
'   Kill App.Path & "\backup2.mdb" 'nickrel2
   FileCopy App.Path & "\backup.mdb", App.Path & "\backup2.mdb"
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
        Splanmain.Show
        Unload Me
     Else
        Check1.Value = vbChecked
        tipnum = 0
        Command1_Click (2)
     End If
  End If
  
'  If IsNumeric(RS("inc_depth")) Then inc_depth = CDbl(RS("inc_depth")) Else inc_depth = 3#
  inc_time = 1

End Sub

Private Sub Form_Unload(Cancel As Integer)
 '   DB.Close
 If tipnum = 0 Then
 Else
    Splanmain.Show
 End If
End Sub

